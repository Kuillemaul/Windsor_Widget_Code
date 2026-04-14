param(
    [string]$PythonExe = 'python',
    [string]$InnoSetupCompiler = '',
    [switch]$SkipPyInstaller,
    [switch]$Clean,
    [string]$OdbcInstallerPath = ''
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

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $root
try {
    if ($Clean) {
        Remove-Item -Recurse -Force .\dist -ErrorAction SilentlyContinue
        Remove-Item -Recurse -Force .\build-temp -ErrorAction SilentlyContinue
        Remove-Item -Recurse -Force .\output -ErrorAction SilentlyContinue
    }

if ($OdbcInstallerPath) {
    if (-not (Test-Path $OdbcInstallerPath)) {
        throw "ODBC installer not found: $OdbcInstallerPath"
    }
    New-Item -ItemType Directory -Path .\prereqs -Force | Out-Null
    $ext = [System.IO.Path]::GetExtension($OdbcInstallerPath)
    $dest = Join-Path $root ("prereqs\msodbcsql18" + $ext)
    Copy-Item $OdbcInstallerPath $dest -Force
    Write-Host "Staged ODBC installer: $dest"
}

    if (-not $SkipPyInstaller) {
        & .\build\build_client.ps1 -PythonExe $PythonExe
    }

    $iscc = Resolve-IsccPath -Preferred $InnoSetupCompiler
    if (-not $iscc) {
        throw "Could not find ISCC.exe. Install Inno Setup 6 or pass -InnoSetupCompiler <full path>."
    }

    New-Item -ItemType Directory -Path .\output -Force | Out-Null
    & $iscc "/O$root\output" .\installer\WindsorWidget.client.configurable.iss

    Write-Host ''
    Write-Host 'Client installer build complete.' -ForegroundColor Green
    Write-Host "Installer: $root\output\WindsorWidget_Client_1_0_0.exe"
    Write-Host "App icon: $root\assets\windsor_icon.ico"
}
finally {
    Pop-Location
}
