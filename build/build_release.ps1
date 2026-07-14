param(
    [string]$Version = "1.0.0",
    [string]$UpdateSharedRoot = "",
    [string]$PythonExe = "python",
    [switch]$SkipTests
)

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $projectRoot

if (-not $SkipTests) {
    & $PythonExe -m pytest -q
    if ($LASTEXITCODE -ne 0) {
        throw "Las pruebas fallaron."
    }
}

& $PythonExe -m PyInstaller --clean --noconfirm build/pyinstaller/AutomatizacionDocumental.spec
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller fall?."
}

$templatePath = Join-Path $projectRoot "build/installer/update_settings.template.json"
$settingsPath = Join-Path $projectRoot "build/installer/update_settings.json"
$publishRoot = $UpdateSharedRoot
if ([string]::IsNullOrWhiteSpace($publishRoot)) {
    $publishRoot = Join-Path $projectRoot "dist/published"
}

$templateData = Get-Content $templatePath -Raw -Encoding UTF8 | ConvertFrom-Json
$templateData.shared_root = $publishRoot
$templateData | ConvertTo-Json -Depth 4 | Set-Content -Path $settingsPath -Encoding UTF8

$iscc = Get-Command iscc -ErrorAction SilentlyContinue
if (-not $iscc) {
    $fallbackIscc = "C:\Users\wandica\AppData\Local\Programs\Inno Setup 6\ISCC.exe"
    if (Test-Path $fallbackIscc) {
        $iscc = @{ Source = $fallbackIscc }
    }
}
if (-not $iscc) {
    Write-Warning "Inno Setup no est? instalado. Se dej? preparado el script build/installer/AutomatizacionDocumental.iss."
    return
}

New-Item -ItemType Directory -Force -Path (Join-Path $projectRoot "dist/installer") | Out-Null

$isccArgs = @(
    "/DMyAppVersion=$Version",
    "/DMySourceDir=$projectRoot\dist\AutomatizacionDocumental",
    "/DMyOutputDir=$projectRoot\dist\installer",
    "/DMyUpdateSettingsFile=$settingsPath",
    "$projectRootuild\installer\AutomatizacionDocumental.iss"
)
$innoProcess = Start-Process -FilePath $iscc.Source -ArgumentList $isccArgs -Wait -PassThru -NoNewWindow
if ($innoProcess.ExitCode -ne 0) {
    throw "La compilaci?n del instalador fall?."
}

$installerName = "AutomatizacionDocumentalSetup_$Version.exe"
$installerSource = Join-Path $projectRoot "dist/installer/$installerName"
New-Item -ItemType Directory -Force -Path $publishRoot | Out-Null
Copy-Item -Path $installerSource -Destination (Join-Path $publishRoot $installerName) -Force

$manifest = @{
    version = $Version
    installer_name = $installerName
    published_at = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
    notes = "Versi?n publicada desde build_release.ps1"
    mandatory = $false
} | ConvertTo-Json -Depth 4

Set-Content -Path (Join-Path $publishRoot "latest.json") -Value $manifest -Encoding UTF8
