# Installs Invoice Converter on this Windows PC.
# Run it by double-clicking "Install Invoice Converter (Windows).bat".
# It installs Python and Tesseract if needed, builds the app, puts it in
# your programs folder and adds a Desktop shortcut. Safe to run again to update.

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot
$AppName = "Invoice Converter"

function Step($msg) { Write-Host ""; Write-Host "==> $msg" -ForegroundColor Cyan }
function Fail($msg) {
    Write-Host ""; Write-Host "PROBLEM: $msg" -ForegroundColor Red
    Read-Host "Press Enter to close"; exit 1
}
function Have-Winget { return [bool](Get-Command winget -ErrorAction SilentlyContinue) }

# ---- 1. Python 3.12 -------------------------------------------------------
Step "Checking for Python 3.12"
function Find-Python {
    $cands = @(
        "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe",
        "C:\Program Files\Python312\python.exe"
    )
    foreach ($c in $cands) { if (Test-Path $c) { return $c } }
    if (Get-Command py -ErrorAction SilentlyContinue) {
        try {
            $ErrorActionPreference = "Continue"
            $p = & py -3.12 -c "import sys; print(sys.executable)" 2>$null
            if ($LASTEXITCODE -eq 0 -and $p) { return "$p".Trim() }
        } catch { }
    }
    return $null
}
$py = Find-Python
if (-not $py) {
    if (-not (Have-Winget)) {
        Fail "Python 3.12 isn't installed. Install it from https://www.python.org/downloads/ (version 3.12), then run this again."
    }
    Write-Host "Installing Python 3.12 (a few minutes)..."
    winget install -e --id Python.Python.3.12 --scope user --silent --accept-package-agreements --accept-source-agreements
    $py = Find-Python
    if (-not $py) { Fail "Python didn't install. Install Python 3.12 from https://www.python.org/downloads/ and run this again." }
}
Write-Host "Using $py"

# ---- 2. Tesseract (the text reader; gets bundled inside the app) -----------
Step "Checking for Tesseract (the text reader)"
$tess = "C:\Program Files\Tesseract-OCR"
if (-not (Test-Path "$tess\tesseract.exe")) {
    if (-not (Have-Winget)) {
        Fail "Tesseract isn't installed. Download the Windows installer from https://github.com/UB-Mannheim/tesseract/wiki , install it with the default options, then run this again."
    }
    Write-Host "Installing Tesseract. If Windows asks for permission, click Yes."
    winget install -e --id UB-Mannheim.TesseractOCR --silent --accept-package-agreements --accept-source-agreements
    if (-not (Test-Path "$tess\tesseract.exe")) {
        Fail "Tesseract didn't install to $tess. Install it from https://github.com/UB-Mannheim/tesseract/wiki (default options), then run this again."
    }
}
Write-Host "Found $tess\tesseract.exe"

# ---- 3. Build ---------------------------------------------------------------
Step "Setting up the build (first time takes a few minutes)"
foreach ($d in ".build-venv", "build", "dist") { if (Test-Path $d) { Remove-Item $d -Recurse -Force } }
& $py -m venv .build-venv
if ($LASTEXITCODE -ne 0) { Fail "Couldn't set up Python." }
$vpy = ".\.build-venv\Scripts\python.exe"
& $vpy -m pip install --upgrade pip -q
& $vpy -m pip install -q -r requirements.txt "pyinstaller==6.11.1" "pyinstaller-hooks-contrib==2024.10"
if ($LASTEXITCODE -ne 0) { Fail "Couldn't download the Python packages. Check the internet connection and run this again." }

Step "Building the app (a few minutes)"
& $vpy -m PyInstaller --noconfirm --windowed --name $AppName `
    --add-data "ui.html;." --add-data "template.xlsx;." --add-data "labels_template.docx;." `
    --add-data "$tess;tesseract" `
    --collect-all webview --collect-all pymupdf `
    --hidden-import webview.platforms.winforms `
    app.py
if ($LASTEXITCODE -ne 0) { Fail "The build failed. Send the text above to Claude." }

Step "Checking the app starts"
$exe = "dist\$AppName\$AppName.exe"
$report = Join-Path $env:TEMP "invoice-converter-selftest.txt"
if (Test-Path $report) { Remove-Item $report -Force }
$p = Start-Process -FilePath $exe -ArgumentList "--selftest", "--report=`"$report`"" -Wait -PassThru
if (Test-Path $report) { Get-Content $report }
if ($p.ExitCode -ne 0) { Fail "The app built but didn't pass its self-test. Send the text above to Claude." }

# ---- 4. Install + shortcuts --------------------------------------------------
Step "Installing"
$dest = Join-Path $env:LOCALAPPDATA "Programs\$AppName"
Get-Process -Name $AppName -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 500
if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
New-Item -ItemType Directory -Path (Split-Path $dest) -Force | Out-Null
Copy-Item "dist\$AppName" $dest -Recurse

$shell = New-Object -ComObject WScript.Shell
$targets = @(
    (Join-Path ([Environment]::GetFolderPath("Desktop")) "$AppName.lnk"),
    (Join-Path ([Environment]::GetFolderPath("Programs")) "$AppName.lnk")
)
foreach ($lnk in $targets) {
    $s = $shell.CreateShortcut($lnk)
    $s.TargetPath = Join-Path $dest "$AppName.exe"
    $s.WorkingDirectory = $dest
    $s.Save()
}

# tidy up the build files
foreach ($d in ".build-venv", "build", "dist") { if (Test-Path $d) { Remove-Item $d -Recurse -Force -ErrorAction SilentlyContinue } }
Get-ChildItem -Filter "*.spec" | Remove-Item -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "Done! Invoice Converter is on your Desktop and in the Start menu." -ForegroundColor Green
Write-Host "You can delete this folder now."
Read-Host "Press Enter to close"
