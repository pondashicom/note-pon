[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$projectPath = Join-Path $repositoryRoot 'NotePon.csproj'
$installerScript = Join-Path $PSScriptRoot 'NOTE-PON.iss'
$outputPath = Join-Path $repositoryRoot 'publish\installer\NOTE-PON-Setup-0.2.3.exe'

$compilerCandidates = @(
    (Join-Path $env:LOCALAPPDATA 'Programs\Inno Setup 6\ISCC.exe'),
    'C:\Program Files (x86)\Inno Setup 6\ISCC.exe',
    'C:\Program Files\Inno Setup 6\ISCC.exe'
)

$compilerPath = $compilerCandidates |
    Where-Object { Test-Path -LiteralPath $_ } |
    Select-Object -First 1

if (-not $compilerPath) {
    throw 'Inno Setup 6 was not found. Run: winget install --id JRSoftware.InnoSetup --exact'
}

$notePonPreviousLib = $env:LIB
try {
    $env:LIB = ''
    & dotnet build $projectPath --configuration Release
    if ($LASTEXITCODE -ne 0) {
        throw "NOTE-PON build failed with exit code $LASTEXITCODE."
    }
}
finally {
    $env:LIB = $notePonPreviousLib
}

$requiredPayload = @(
    'bin\Release\net10.0-windows\NOTE-PON.exe',
    'bin\Release\net10.0-windows\NOTE-PON.dll',
    'bin\Release\net10.0-windows\NOTE-PON.deps.json',
    'bin\Release\net10.0-windows\NOTE-PON.runtimeconfig.json',
    'README.md',
    'LICENSE'
)

foreach ($relativePath in $requiredPayload) {
    $fullPath = Join-Path $repositoryRoot $relativePath
    if (-not (Test-Path -LiteralPath $fullPath)) {
        throw "Installer payload file was not found: $fullPath"
    }
}

& $compilerPath $installerScript
if ($LASTEXITCODE -ne 0) {
    throw "Installer build failed with exit code $LASTEXITCODE."
}

if (-not (Test-Path -LiteralPath $outputPath)) {
    throw "Installer output was not generated: $outputPath"
}

Get-Item -LiteralPath $outputPath |
    Select-Object FullName, Length, LastWriteTime
