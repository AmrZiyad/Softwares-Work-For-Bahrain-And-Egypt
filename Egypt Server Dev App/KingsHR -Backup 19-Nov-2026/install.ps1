# Kings School Egypt - HR Web App Installer
# Run as Administrator: powershell -ExecutionPolicy Bypass -File install.ps1

param(
    [string]$AppDir      = "C:\KingsHR",
    [string]$AppPort     = "80",
    [string]$ServiceName = "KingsHRApp"
)

$ErrorActionPreference = "Stop"

$PythonVersion   = "3.11.9"
$PythonInstaller = "python-$PythonVersion-amd64.exe"
$PythonUrl       = "https://www.python.org/ftp/python/$PythonVersion/$PythonInstaller"
$NSSMUrl         = "https://nssm.cc/release/nssm-2.24.zip"
$NSSMDir         = "C:\nssm"

function Write-Step([string]$msg) { Write-Host "" ; Write-Host "==> $msg" -ForegroundColor Cyan }
function Write-OK([string]$msg)   { Write-Host "    OK: $msg" -ForegroundColor Green }

# ----------------------------------------------------------------
# 1. Python
# ----------------------------------------------------------------
Write-Step "Checking Python..."

$pyVer = ""
try { $pyVer = (& python --version 2>&1).ToString() } catch {}

if ($pyVer -notmatch "Python 3") {
    Write-Host "    Downloading Python $PythonVersion ..." -ForegroundColor Yellow
    $tmp = "$env:TEMP\$PythonInstaller"
    Invoke-WebRequest -Uri $PythonUrl -OutFile $tmp -UseBasicParsing
    Start-Process -FilePath $tmp -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1 Include_test=0" -Wait
    Remove-Item $tmp -Force
    $mPath = [System.Environment]::GetEnvironmentVariable("Path", "Machine")
    $uPath = [System.Environment]::GetEnvironmentVariable("Path", "User")
    $env:Path = $mPath + ";" + $uPath
    Write-OK "Python $PythonVersion installed."
} else {
    Write-OK "$pyVer already installed."
}

# ----------------------------------------------------------------
# 2. Copy files (skip if already in place)
# ----------------------------------------------------------------
Write-Step "Checking app files in $AppDir ..."

$scriptDir = $PSScriptRoot.TrimEnd("\")
$destDir   = $AppDir.TrimEnd("\")

if ($scriptDir.ToLower() -eq $destDir.ToLower()) {
    Write-OK "Script is already running from $AppDir - no copy needed."
} else {
    if (-not (Test-Path $AppDir)) { New-Item -ItemType Directory -Path $AppDir | Out-Null }
    Get-ChildItem -Path $scriptDir | ForEach-Object {
        Copy-Item -Path $_.FullName -Destination $AppDir -Recurse -Force
    }
    Write-OK "Files copied to $AppDir."
}

# ----------------------------------------------------------------
# 3. pip packages
# ----------------------------------------------------------------
Write-Step "Installing Python packages..."
Set-Location $AppDir
& python -m pip install --upgrade pip --quiet
& python -m pip install -r requirements.txt --quiet
Write-OK "Packages installed."

# ----------------------------------------------------------------
# 4. ODBC Driver 18
# ----------------------------------------------------------------
Write-Step "Checking ODBC Driver 18 for SQL Server..."
$odbcKey = "HKLM:\SOFTWARE\ODBC\ODBCINST.INI\ODBC Driver 18 for SQL Server"
if (-not (Test-Path $odbcKey)) {
    Write-Host "    Downloading ODBC Driver 18..." -ForegroundColor Yellow
    $odbcTmp = "$env:TEMP\msodbcsql18.msi"
    Invoke-WebRequest -Uri "https://go.microsoft.com/fwlink/?linkid=2249006" -OutFile $odbcTmp -UseBasicParsing
    Start-Process msiexec -ArgumentList "/i `"$odbcTmp`" /quiet IACCEPTMSODBCSQLLICENSETERMS=YES" -Wait
    Remove-Item $odbcTmp -Force
    Write-OK "ODBC Driver 18 installed."
} else {
    Write-OK "ODBC Driver 18 already present."
}

# ----------------------------------------------------------------
# 5. Firewall
# ----------------------------------------------------------------
Write-Step "Adding firewall rule for port $AppPort ..."
$ruleName = "KingsHRApp-Port$AppPort"
if (-not (Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue)) {
    New-NetFirewallRule -DisplayName $ruleName -Direction Inbound -Protocol TCP -LocalPort $AppPort -Action Allow | Out-Null
    Write-OK "Firewall rule added."
} else {
    Write-OK "Firewall rule already exists."
}

# ----------------------------------------------------------------
# 6. NSSM
# ----------------------------------------------------------------
Write-Step "Setting up NSSM service manager..."
$nssmExe = "$NSSMDir\nssm.exe"
if (-not (Test-Path $nssmExe)) {
    if (-not (Test-Path $NSSMDir)) { New-Item -ItemType Directory -Path $NSSMDir | Out-Null }
    $nssmZip = "$env:TEMP\nssm.zip"
    $nssmTmp = "$env:TEMP\nssm_ext"
    Invoke-WebRequest -Uri $NSSMUrl -OutFile $nssmZip -UseBasicParsing
    Expand-Archive -Path $nssmZip -DestinationPath $nssmTmp -Force
    Copy-Item "$nssmTmp\nssm-2.24\win64\nssm.exe" $nssmExe -Force
    Remove-Item $nssmZip -Force
    Remove-Item $nssmTmp -Recurse -Force
    Write-OK "NSSM ready."
} else {
    Write-OK "NSSM already present."
}

# ----------------------------------------------------------------
# 7. Windows service
# ----------------------------------------------------------------
Write-Step "Registering service $ServiceName ..."

# Find python.exe - search common install paths
$pythonExe = ""
$candidates = @(
    (Get-Command python -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source),
    "C:\Python311\python.exe",
    "C:\Program Files\Python311\python.exe",
    "$env:LOCALAPPDATA\Programs\Python\Python311\python.exe"
)
foreach ($c in $candidates) {
    if ($c -and (Test-Path $c)) { $pythonExe = $c; break }
}
if (-not $pythonExe) {
    Write-Host "    ERROR: python.exe not found. Reopen PowerShell as Admin and retry." -ForegroundColor Red
    exit 1
}
Write-Host "    Using: $pythonExe" -ForegroundColor Gray

# Find gunicorn.exe next to python.exe
$pythonDir  = Split-Path $pythonExe
$gunicornExe = "$pythonDir\Scripts\gunicorn.exe"
if (-not (Test-Path $gunicornExe)) {
    # Try to find it anywhere on PATH
    $g = Get-Command gunicorn -ErrorAction SilentlyContinue
    if ($g) { $gunicornExe = $g.Source }
}
Write-Host "    Gunicorn: $gunicornExe" -ForegroundColor Gray

$existing = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($existing) {
    Write-Host "    Removing old service..." -ForegroundColor Yellow
    & $nssmExe stop $ServiceName 2>$null
    Start-Sleep -Seconds 2
    & $nssmExe remove $ServiceName confirm 2>$null
}

$logsDir = "$AppDir\logs"
if (-not (Test-Path $logsDir)) { New-Item -ItemType Directory -Path $logsDir | Out-Null }

if (Test-Path $gunicornExe) {
    # Preferred: use gunicorn.exe directly
    & $nssmExe install $ServiceName $gunicornExe
    & $nssmExe set $ServiceName AppParameters "-w 2 -b 0.0.0.0:$AppPort app:app"
} else {
    # Fallback: python -m gunicorn
    & $nssmExe install $ServiceName $pythonExe
    & $nssmExe set $ServiceName AppParameters "-m gunicorn -w 2 -b 0.0.0.0:$AppPort app:app"
}

& $nssmExe set $ServiceName AppDirectory  $AppDir
& $nssmExe set $ServiceName DisplayName   "Kings School Egypt HR App"
& $nssmExe set $ServiceName Start         SERVICE_AUTO_START
& $nssmExe set $ServiceName AppStdout     "$logsDir\service.log"
& $nssmExe set $ServiceName AppStderr     "$logsDir\service.log"

Write-Step "Starting service..."
& $nssmExe start $ServiceName
Start-Sleep -Seconds 5

$status = (Get-Service -Name $ServiceName).Status
if ($status -eq "Paused" -or $status -eq "Stopped") {
    Write-Host "    Service did not start cleanly. Checking log..." -ForegroundColor Yellow
    if (Test-Path "$logsDir\service.log") {
        Get-Content "$logsDir\service.log" -Tail 20 | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkYellow }
    }
} else {
    Write-OK "Service status: $status"
}

# ----------------------------------------------------------------
# 8. Done
# ----------------------------------------------------------------
$ip = ""
try {
    $addr = Get-NetIPAddress -AddressFamily IPv4 -InterfaceAlias "*Ethernet*" -ErrorAction SilentlyContinue |
            Select-Object -First 1
    if ($addr) { $ip = $addr.IPAddress }
} catch {}
if (-not $ip) { $ip = "10.4.172.15" }

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  INSTALLATION COMPLETE" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  App URL  : http://$ip/" -ForegroundColor Green
Write-Host "           : http://localhost/" -ForegroundColor Green
Write-Host "  Service  : $ServiceName  ($status)" -ForegroundColor Green
Write-Host "  App Dir  : $AppDir" -ForegroundColor Green
Write-Host "  Logs     : $logsDir\service.log" -ForegroundColor Green
Write-Host ""
Write-Host "  HR staff open http://10.4.172.15/ from their laptops." -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
