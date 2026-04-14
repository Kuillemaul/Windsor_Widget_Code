param(
    [string]$RepoRoot = '',
    [string]$PythonExe = 'python',
    [string]$InnoSetupCompiler = '',
    [string]$OdbcInstallerPath = '',
    [switch]$SkipPyInstaller,
    [switch]$Clean
)

$ErrorActionPreference = 'Stop'

function Resolve-IsccPath {
    param([string]$Preferred)
    $candidates = @()
    if ($Preferred) { $candidates += $Preferred }
    $cmd = Get-Command ISCC.exe -ErrorAction SilentlyContinue
    if ($cmd) { $candidates += $cmd.Source }
    $candidates += @(
        "$env:ProgramFiles(x86)\Inno Setup 6\ISCC.exe",
        "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
    )
    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path $candidate)) {
            return (Resolve-Path $candidate).Path
        }
    }
    return $null
}

function Find-FirstFile {
    param(
        [Parameter(Mandatory=$true)][string]$Root,
        [Parameter(Mandatory=$true)][string[]]$Names
    )
    foreach ($name in $Names) {
        $match = Get-ChildItem -Path $Root -Recurse -File -Filter $name -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($match) { return $match.FullName }
    }
    return $null
}

function Ensure-RepoFile {
    param(
        [Parameter(Mandatory=$true)][string]$Root,
        [Parameter(Mandatory=$true)][string[]]$Names,
        [Parameter(Mandatory=$true)][string]$Label
    )
    $found = Find-FirstFile -Root $Root -Names $Names
    if (-not $found) {
        throw "Could not find $Label. Looked for: $($Names -join ', ') under $Root"
    }
    return (Resolve-Path $found).Path
}

function Stage-FileFlat {
    param(
        [Parameter(Mandatory=$true)][string]$SourcePath,
        [Parameter(Mandatory=$true)][string]$DestDir,
        [Parameter(Mandatory=$true)][string]$DestName
    )
    New-Item -ItemType Directory -Path $DestDir -Force | Out-Null
    Copy-Item $SourcePath (Join-Path $DestDir $DestName) -Force
}

$builderRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $RepoRoot) {
    $RepoRoot = Split-Path -Parent $builderRoot
}
$RepoRoot = (Resolve-Path $RepoRoot).Path

$stageDir = Join-Path $builderRoot 'staged_app'
$distDir = Join-Path $builderRoot 'dist'
$buildTempDir = Join-Path $builderRoot 'build-temp'
$outputDir = Join-Path $builderRoot 'output'
$specPath = Join-Path $builderRoot 'scripts\WindsorWidget.auto.spec'
$issPath = Join-Path $builderRoot 'installer\WindsorWidget.client.configurable.iss'

if ($Clean) {
    Remove-Item -Recurse -Force $stageDir -ErrorAction SilentlyContinue
    Remove-Item -Recurse -Force $distDir -ErrorAction SilentlyContinue
    Remove-Item -Recurse -Force $buildTempDir -ErrorAction SilentlyContinue
    Remove-Item -Recurse -Force $outputDir -ErrorAction SilentlyContinue
}

New-Item -ItemType Directory -Path $stageDir -Force | Out-Null
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

$mainPy = Ensure-RepoFile -Root $RepoRoot -Names @('main_patched_status_yu.py', 'main.py', 'main_patched_status_yu_v9.py') -Label 'main application file'
$uiMain = Ensure-RepoFile -Root $RepoRoot -Names @('ui_mainwindow.py') -Label 'ui_mainwindow.py'
$uiShip = Ensure-RepoFile -Root $RepoRoot -Names @('shipments_window_ui.py', 'ui_shipments_window.py', 'ui_shipments_window(1).py') -Label 'shipments window UI'
$yuWorkflow = Ensure-RepoFile -Root $RepoRoot -Names @('yu_order_workflow.py', 'yu_order_workflow_v3.py') -Label 'YU workflow module'
$yuReview = Ensure-RepoFile -Root $RepoRoot -Names @('yu_order_review_export_test_window.py') -Label 'YU review/export window'
$monthPicker = Find-FirstFile -Root $RepoRoot -Names @('month_year_picker.py')
$iconFile = Find-FirstFile -Root $RepoRoot -Names @('windsor_icon.ico', 'windsor_logo.ico', '*.ico')
$requirements = Find-FirstFile -Root $RepoRoot -Names @('requirements.txt')

Write-Host "Staging discovered source files..." -ForegroundColor Cyan
Stage-FileFlat -SourcePath $mainPy -DestDir $stageDir -DestName 'main_patched_status_yu.py'
Stage-FileFlat -SourcePath $uiMain -DestDir $stageDir -DestName 'ui_mainwindow.py'
Stage-FileFlat -SourcePath $uiShip -DestDir $stageDir -DestName 'shipments_window_ui.py'
Stage-FileFlat -SourcePath $yuWorkflow -DestDir $stageDir -DestName 'yu_order_workflow.py'
Stage-FileFlat -SourcePath $yuReview -DestDir $stageDir -DestName 'yu_order_review_export_test_window.py'
if ($monthPicker) {
    Stage-FileFlat -SourcePath $monthPicker -DestDir $stageDir -DestName 'month_year_picker.py'
} else {
    Write-Warning "month_year_picker.py was not found. Build may fail if the app imports it."
}

$assetsStage = Join-Path $stageDir 'assets'
New-Item -ItemType Directory -Path $assetsStage -Force | Out-Null

if ($iconFile) {
    Copy-Item $iconFile (Join-Path $builderRoot 'assets.ico') -Force
    Write-Host "Using icon: $iconFile"
} else {
    Write-Warning "No .ico file found in repo. Installer/app will build without a custom icon."
}

if ($requirements) {
    Write-Host "requirements.txt found at: $requirements" -ForegroundColor Green
    Write-Host "Installing requirements from repo..." -ForegroundColor Cyan
    & $PythonExe -m pip install -r $requirements
} else {
    Write-Warning "No requirements.txt found in repo. Installing fallback packages."
    & $PythonExe -m pip install PyInstaller PySide6 pyodbc openpyxl
}

if ($OdbcInstallerPath) {
    if (-not (Test-Path $OdbcInstallerPath)) {
        throw "ODBC installer not found: $OdbcInstallerPath"
    }
    New-Item -ItemType Directory -Path (Join-Path $builderRoot 'prereqs') -Force | Out-Null
    $ext = [System.IO.Path]::GetExtension($OdbcInstallerPath)
    $dest = Join-Path $builderRoot ("prereqs\msodbcsql18" + $ext)
    Copy-Item $OdbcInstallerPath $dest -Force
    Write-Host "Staged ODBC installer: $dest"
}

if (-not $SkipPyInstaller) {
    Write-Host "Building app with PyInstaller..." -ForegroundColor Cyan
    & $PythonExe -m PyInstaller --noconfirm --clean $specPath
}

$exePath = Join-Path $distDir 'WindsorWidget\WindsorWidget.exe'
if (-not (Test-Path $exePath)) {
    throw "Expected EXE not found after build: $exePath"
}

$iscc = Resolve-IsccPath -Preferred $InnoSetupCompiler
if (-not $iscc) {
    throw "Could not find ISCC.exe. Install Inno Setup 6 or pass -InnoSetupCompiler <full path>."
}

Write-Host "Compiling installer..." -ForegroundColor Cyan
& $iscc "/O$outputDir" $issPath

Write-Host ""
Write-Host "Build complete." -ForegroundColor Green
Write-Host "EXE: $exePath"
Write-Host "Installer: $(Join-Path $outputDir 'WindsorWidget_Client_1_0_0.exe')"
