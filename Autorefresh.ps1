<#
.SYNOPSIS
Abre Google Chrome controlado por Selenium y refresca una URL cada X segundos.

.DESCRIPTION
- Detecta la ruta de chrome.exe y su versión instalada.
- Garantiza (opcionalmente) que el ChromeDriver coincida con la versión mayor de Chrome, descargándolo si es necesario.
- Lanza un navegador Chrome (opcionalmente en modo headless) y recarga la página en un intervalo configurable.
- Permite reutilizar un perfil de usuario de Chrome para evitar tener que iniciar sesión cada vez.

.PARAMETER Url
URL a abrir y refrescar periódicamente.

.PARAMETER Intervalo
Intervalo de refresco en segundos (mínimo 3).

.PARAMETER DriverDir
Directorio donde se almacenará/leerá chromedriver.exe.

.PARAMETER Headless
Ejecuta Chrome en modo headless.

.PARAMETER UserDataDir
Ruta a un directorio de datos de usuario de Chrome para reutilizar sesión (por ejemplo, C:\Users\<usuario>\AppData\Local\Google\Chrome\User Data).

.PARAMETER ProfileDirectory
Nombre del perfil dentro de UserDataDir (por ejemplo, "Default", "Profile 1", etc.).

.PARAMETER NoDriverDownload
Si se especifica, no intentará descargar/actualizar ChromeDriver automáticamente.

.PARAMETER VerboseLogging
Muestra trazas detalladas de lo que hace el script.

.EXAMPLE
./Autorefresh.ps1 -Url "https://midominio/mi-pagina" -Intervalo 60 -Headless

.EXAMPLE
./Autorefresh.ps1 -Url "https://..." -UserDataDir "$env:LOCALAPPDATA\Google\Chrome\User Data" -ProfileDirectory "Default"
#>

param (
  [Parameter(Mandatory = $false)]
  [string]$Url = 'https://viewnext.atlassian.net/jira/servicedesk/projects/SIS/queues/custom/221',

  [Parameter(Mandatory = $false)]
  [ValidateRange(3, 86400)]
  [int]$Intervalo = 60,

  [Parameter(Mandatory = $false)]
  [string]$DriverDir = 'C:\Tools\ChromeDriver',

  [switch]$Headless,

  [Parameter(Mandatory = $false)]
  [string]$UserDataDir,

  [Parameter(Mandatory = $false)]
  [string]$ProfileDirectory,

  [switch]$NoDriverDownload,

  [switch]$VerboseLogging
)

if ($VerboseLogging) { $VerbosePreference = 'Continue' }
$ErrorActionPreference = 'Stop'

function Write-Info($msg) { Write-Host "[INFO] $msg" }
function Write-Warn($msg) { Write-Warning $msg }

# ===== Encontrar chrome.exe y su versión =====
function Get-ChromePath {
  $cands = @(
    'C:\Program Files\Google\Chrome\Application\chrome.exe',
    'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
    "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
  )
  foreach ($p in $cands) {
    if (Test-Path $p) { return $p }
  }
  $appPathHKLM = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe'
  $appPathHKCU = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe'
  foreach ($k in @($appPathHKLM, $appPathHKCU)) {
    try {
      $v = (Get-ItemProperty $k -ErrorAction Stop).'(default)'
      if ($v -and (Test-Path $v)) { return $v }
    } catch {}
  }
  return $null
}

function Get-ChromeVersion([string]$ChromeExe) {
  foreach ($k in @('HKLM:\SOFTWARE\Google\Chrome\BLBeacon','HKCU:\SOFTWARE\Google\Chrome\BLBeacon')) {
    try {
      $v = (Get-ItemProperty $k -ErrorAction Stop).version
      if ($v) { return $v }
    } catch {}
  }
  if ($ChromeExe -and (Test-Path $ChromeExe)) {
    try {
      $v = (& $ChromeExe --version) 2>$null  # ej: "Google Chrome 128.0.XXXX"
      if ($v) { return ($v -replace '.*?(\d+\.\d+\.\d+\.\d+).*', '$1') }
    } catch {}
  }
  return $null
}

# ===== Asegurar chromedriver.exe correcto =====
New-Item -ItemType Directory -Path $DriverDir -Force | Out-Null
$DriverExe = Join-Path $DriverDir 'chromedriver.exe'

function Get-DriverMajor([string]$DriverPath) {
  if (Test-Path $DriverPath) {
    try {
      $drv = (& $DriverPath --version) 2>$null
      if ($drv) { return ($drv -replace '.*?(\d+)\..*', '$1') }
    } catch {}
  }
  return $null
}

# ===== Flujo principal =====
$chromePath = Get-ChromePath
if (-not $chromePath) { throw 'No se encontró chrome.exe en el sistema.' }
Write-Verbose "Chrome en: $chromePath"

$chromeVer = Get-ChromeVersion -ChromeExe $chromePath
if (-not $chromeVer) { throw 'No pude obtener la versión de Chrome.' }
$chromeMajor = ($chromeVer -split '\.')[0]
Write-Info "Chrome $chromeVer (major $chromeMajor)"

$driverMajor = Get-DriverMajor -DriverPath $DriverExe
if ($driverMajor) { Write-Info "ChromeDriver detectado en $DriverExe (major $driverMajor)" }
else { Write-Verbose 'No hay ChromeDriver presente.' }

if (-not $NoDriverDownload) {
  $needDownload = $true
  if ($driverMajor -and ($driverMajor -eq $chromeMajor)) { $needDownload = $false }

  if ($needDownload) {
    $platform = if ([Environment]::Is64BitOperatingSystem) { 'win64' } else { 'win32' }
    $jsonUrl  = 'https://googlechromelabs.github.io/chrome-for-testing/known-good-versions-with-downloads.json'

    Write-Info 'Buscando versión compatible de ChromeDriver...'
    try {
      $data = Invoke-RestMethod -Uri $jsonUrl -UseBasicParsing -TimeoutSec 30
    } catch {
      throw "No pude consultar $jsonUrl. Asegúrate de tener conexión a Internet."
    }

    $match = $data.versions |
             Where-Object { $_.version -like "$chromeMajor.*" } |
             Sort-Object { [Version]$_.version } |
             Select-Object -Last 1
    if (-not $match) { throw "No hay ChromeDriver para major $chromeMajor" }

    $dl = $match.downloads.chromedriver | Where-Object { $_.platform -eq $platform } | Select-Object -First 1
    if (-not $dl) { throw "No hay build de ChromeDriver para $platform" }

    $zipPath = Join-Path $env:TEMP "chromedriver-$($match.version)-$platform.zip"
    Write-Info "Descargando ChromeDriver $($match.version) ($platform)..."
    Invoke-WebRequest -Uri $dl.url -OutFile $zipPath -UseBasicParsing

    $tmpDir = Join-Path $env:TEMP "chromedriver-$($match.version)-$platform"
    if (Test-Path $tmpDir) { Remove-Item $tmpDir -Recurse -Force }
    Expand-Archive -Path $zipPath -DestinationPath $tmpDir -Force

    $found = Get-ChildItem -Path $tmpDir -Recurse -Filter 'chromedriver.exe' | Select-Object -First 1
    if (-not $found) { throw 'El ZIP no contiene chromedriver.exe' }
    Copy-Item $found.FullName $DriverExe -Force

    Remove-Item $zipPath -Force
    Remove-Item $tmpDir -Recurse -Force

    $driverMajor = Get-DriverMajor -DriverPath $DriverExe
    Write-Info "ChromeDriver actualizado (major $driverMajor)"
  } else {
    Write-Verbose 'ChromeDriver ya coincide con la versión de Chrome.'
  }
} else {
  if (-not (Test-Path $DriverExe)) {
    throw "NoDriverDownload especificado pero no existe $DriverExe"
  }
  if ($driverMajor -and ($driverMajor -ne $chromeMajor)) {
    Write-Warn "Advertencia: ChromeDriver (major $driverMajor) no coincide con Chrome (major $chromeMajor). Puede fallar."
  }
}

# ===== Lanzar Selenium + autorefresh =====
try {
  try {
    Import-Module Selenium -ErrorAction Stop
  } catch {
    Write-Warn "No se pudo importar el módulo 'Selenium'. Instálalo si es necesario: Install-Module Selenium -Scope CurrentUser"
    throw
  }

  $service = [OpenQA.Selenium.Chrome.ChromeDriverService]::CreateDefaultService($DriverDir)
  $service.HideCommandPromptWindow = $true
  $service.SuppressInitialDiagnosticInformation = $true

  $options = New-Object OpenQA.Selenium.Chrome.ChromeOptions
  $options.BinaryLocation = $chromePath
  $options.AddArgument('--remote-allow-origins=*')
  $options.AddArgument('--disable-gpu')
  if ($Headless) {
    $options.AddArgument('--headless=new')
    $options.AddArgument('--window-size=1920,1080')
  }

  # Establecer valores por defecto para reutilizar sesión si no se proporcionan
  try {
    if (-not $UserDataDir) {
      $defaultUD = Join-Path $env:LOCALAPPDATA 'Google\Chrome\User Data'
      if (Test-Path $defaultUD) {
        $UserDataDir = $defaultUD
        Write-Verbose "Usando UserDataDir por defecto: $UserDataDir"
      }
    }
    if (-not $ProfileDirectory -and $UserDataDir) {
      $candidateProfile = 'Default'
      $candidatePath = Join-Path $UserDataDir $candidateProfile
      if (Test-Path $candidatePath) {
        $ProfileDirectory = $candidateProfile
        Write-Verbose "Usando ProfileDirectory por defecto: $ProfileDirectory"
      }
    }
  } catch {}

  if ($UserDataDir) { $options.AddArgument("--user-data-dir=$UserDataDir") }
  if ($ProfileDirectory) { $options.AddArgument("--profile-directory=$ProfileDirectory") }

  $driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($service, $options)
  if (-not $driver) { throw 'No se pudo crear WebDriver.' }

  Write-Info "Abriendo: $Url"
  $driver.Url = $Url

  while ($true) {
    Start-Sleep -Seconds $Intervalo
    ([OpenQA.Selenium.IJavaScriptExecutor]$driver).ExecuteScript('location.reload(true);') | Out-Null
    Write-Host ("🔄 Refresco: {0:yyyy-MM-dd HH:mm:ss}" -f (Get-Date))
  }
}
finally {
  if ($driver) {
    try { $driver.Quit() } catch {}
    try { $driver.Dispose() } catch {}
  }
  if ($service) {
    try { $service.Dispose() } catch {}
  }
}
