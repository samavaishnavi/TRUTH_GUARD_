# TruthGuard Mobile Test Runner & Orchestrator
# This script helps you run the Appium E2E Python mobile tests locally.

$Host.UI.RawUI.WindowTitle = "TruthGuard Mobile Test Orchestrator"

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "     🛡️  TRUTHGUARD PYTHON APPIUM RUNNER & ORCHESTRATOR  🛡️" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "This script will orchestrate local Android E2E testing."
Write-Host ""

function Ensure-Path {
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
}

Ensure-Path

# 1. Check for connected Android emulator / device
Write-Host ">>> Checking connected Android devices..." -ForegroundColor Yellow
$devices = & "C:\Users\HP\AppData\Local\Android\Sdk\platform-tools\adb.exe" devices
if ($devices -match "device$") {
    Write-Host "Connected Android device/emulator detected." -ForegroundColor Green
} else {
    Write-Host "No connected device. Attempting to start emulator 'Pixel_6'..." -ForegroundColor Yellow
    Start-Process "C:\Users\HP\AppData\Local\Android\Sdk\emulator\emulator.exe" -ArgumentList "-avd Pixel_6 -no-audio" -WindowStyle Minimized
    Write-Host "Waiting for emulator to register via ADB (this may take up to 45 seconds)..." -ForegroundColor Yellow
    
    # Wait loop
    $booted = $false
    for ($i = 0; $i -lt 30; $i++) {
        Start-Sleep -Seconds 2
        $devices = & "C:\Users\HP\AppData\Local\Android\Sdk\platform-tools\adb.exe" devices
        if ($devices -match "emulator-5554`tdevice") {
            $booted = $true
            break
        }
    }
    if ($booted) {
        Write-Host "Emulator 'Pixel_6' successfully booted and online." -ForegroundColor Green
    } else {
        Write-Host "Warning: Could not auto-detect booted emulator. Attempting to run anyway..." -ForegroundColor Red
    }
}

# 2. Check for Appium server
Write-Host "`n>>> Checking Appium server connection on port 4723..." -ForegroundColor Yellow
$appiumRunning = Test-NetConnection -ComputerName 127.0.0.1 -Port 4723 -WarningAction SilentlyContinue
if ($appiumRunning.TcpTestSucceeded) {
    Write-Host "Appium server is already running." -ForegroundColor Green
} else {
    Write-Host "Launching Appium server in a minimized window..." -ForegroundColor Yellow
    $env:ANDROID_HOME = "C:\Users\HP\AppData\Local\Android\Sdk"
    $env:JAVA_HOME = "C:\Program Files\Android\Android Studio\jbr"
    Start-Process cmd -ArgumentList "/c appium" -WindowStyle Minimized
    Start-Sleep -Seconds 8
}

# 3. Verify Python and pip requirements
Write-Host "`n>>> Verifying Python dependencies..." -ForegroundColor Yellow
$pythonCheck = Get-Command python -ErrorAction SilentlyContinue
if ($null -eq $pythonCheck) {
    Write-Host "Error: 'python' command was not found. Please ensure Python is installed and in your environment PATH." -ForegroundColor Red
    Write-Host "Press any key to exit..."
    [void]$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

Push-Location "appium-testing"
Write-Host "Installing/Updating packages from requirements.txt..." -ForegroundColor Yellow
pip install -r requirements.txt

# 4. Run tests
Write-Host "`n>>> Running Appium Python E2E Tests..." -ForegroundColor Green
$env:ANDROID_HOME = "C:\Users\HP\AppData\Local\Android\Sdk"
$env:JAVA_HOME = "C:\Program Files\Android\Android Studio\jbr"
python test.py
Pop-Location

Write-Host "`n>>> Mobile E2E tests finished. Check 'appium-testing/Appium_Test_Report.xlsx' for details." -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "Press any key to close this runner..."
[void]$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
