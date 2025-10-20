# ===== Config =====
$driverDir  = "C:\Tools\ChromeDriver"
$intervalo  = 60
$headless   = $false

# ===== Encontrar chrome.exe y su versión =====
function Get-ChromePath {
  $cands = @(
    'C:\Program Files\Google\Chrome\Application\chrome.exe',
    'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
    "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
  )
  foreach ($p in $cands) { if (Test-Path $p) { return $p } }
  $appPathHKLM = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe'
  $appPathHKCU = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe'
  foreach ($k in @($appPathHKLM,$appPathHKCU)) {
    try { $v=(Get-ItemProperty $k -ErrorAction Stop).'(default)'; if ($v -and (Test-Path $v)) { return $v } } catch {}
  }
  return $null
}

$chromePath = Get-ChromePath
if (-not $chromePath) { Write-Error "No se encontró chrome.exe"; exit 1 }

# Intenta leer versión desde registro BLBeacon (más rápido)
function Get-ChromeVersion {
  foreach ($k in @('HKLM:\SOFTWARE\Google\Chrome\BLBeacon','HKCU:\SOFTWARE\Google\Chrome\BLBeacon')) {
    try { $v=(Get-ItemProperty $k -ErrorAction Stop).version; if ($v) { return $v } } catch {}
  }
  # fallback ejecutando chrome
  $v = (& $chromePath --version) 2>$null  # ej: "Google Chrome 128.0.XXXX"
  if ($v) { return ($v -replace '.*?(\d+\.\d+\.\d+\.\d+).*','$1') }
  return $null
}

$chromeVer = Get-ChromeVersion
if (-not $chromeVer) { Write-Error "No pude obtener versión de Chrome"; exit 1 }
$chromeMajor = ($chromeVer -split '\.')[0]

# ===== Asegurar chromedriver.exe correcto =====
New-Item -ItemType Directory -Path $driverDir -Force | Out-Null
$driverExe = Join-Path $driverDir 'chromedriver.exe'

function Get-DriverMajor {
  if (Test-Path $driverExe) {
    $drv = (& $driverExe --version) 2>$null
    if ($drv) { return ($drv -replace '.*?(\d+)\..*','$1') }
  }
  return $null
}

$needDownload = $true
$driverMajor = Get-DriverMajor
if ($driverMajor -and ($driverMajor -eq $chromeMajor)) { $needDownload = $false }

if ($needDownload) {
  $platform = if ([Environment]::Is64BitOperatingSystem) { 'win64' } else { 'win32' }
  $jsonUrl  = 'https://googlechromelabs.github.io/chrome-for-testing/known-good-versions-with-downloads.json'
  try { $data = Invoke-RestMethod -Uri $jsonUrl -UseBasicParsing -TimeoutSec 30 }
  catch { Write-Error "No pude consultar $jsonUrl (¿sin Internet?)"; exit 1 }

  $match = $data.versions |
           Where-Object { $_.version -like "$chromeMajor.*" } |
           Sort-Object { [Version]$_.version } |
           Select-Object -Last 1
  if (-not $match) { Write-Error "No hay ChromeDriver para major $chromeMajor"; exit 1 }

  $dl = $match.downloads.chromedriver | Where-Object { $_.platform -eq $platform } | Select-Object -First 1
  if (-not $dl) { Write-Error "No hay build de ChromeDriver para $platform"; exit 1 }

  $zipPath = Join-Path $env:TEMP "chromedriver-$($match.version)-$platform.zip"
  Invoke-WebRequest -Uri $dl.url -OutFile $zipPath -UseBasicParsing
  $tmpDir = Join-Path $env:TEMP "chromedriver-$($match.version)-$platform"
  if (Test-Path $tmpDir) { Remove-Item $tmpDir -Recurse -Force }
  Expand-Archive -Path $zipPath -DestinationPath $tmpDir -Force
  $found = Get-ChildItem -Path $tmpDir -Recurse -Filter 'chromedriver.exe' | Select-Object -First 1
  if (-not $found) { Write-Error "El ZIP no contiene chromedriver.exe"; exit 1 }
  Copy-Item $found.FullName $driverExe -Force
  Remove-Item $zipPath -Force
  Remove-Item $tmpDir -Recurse -Force
}

# ===== Lanzar Selenium + autorefresh =====
Import-Module Selenium

$service = [OpenQA.Selenium.Chrome.ChromeDriverService]::CreateDefaultService($driverDir)
$options = New-Object OpenQA.Selenium.Chrome.ChromeOptions
$options.BinaryLocation = $chromePath
if ($headless) { $options.AddArgument('--headless=new') }

$driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($service, $options)
if (-not $driver) { Write-Error "No se pudo crear WebDriver"; exit 1 }

$baseUrl = "https://viewnext.atlassian.net/jira/servicedesk/projects/SIS/queues/custom/221"
$driver.Url = $baseUrl

try {
  while ($true) {
    Start-Sleep -Seconds $intervalo
    ([OpenQA.Selenium.IJavaScriptExecutor]$driver).ExecuteScript("location.reload(true);")
    Write-Host ("🔄 Refresco: {0:yyyy-MM-dd HH:mm:ss}" -f (Get-Date))
  }
}
finally {
  
}