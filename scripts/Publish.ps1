#Requires -Version 7.0
param([switch]$SkipTests)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
Push-Location -LiteralPath $root
try {
    dotnet restore SharePointDiff.sln --locked-mode
    if ($LASTEXITCODE -ne 0) { throw 'Dependency restore failed.' }
    dotnet build SharePointDiff.sln -c Release --no-restore
    if ($LASTEXITCODE -ne 0) { throw 'Build failed.' }
    if (-not $SkipTests) {
        dotnet test tests/SharePointDiff.Tests -c Release --no-build --no-restore
        if ($LASTEXITCODE -ne 0) { throw 'Tests failed. Nothing published.' }
    }
    dotnet publish src/SharePointDiff.App -c Release -r win-x64 --self-contained false --no-restore -o publish/win-x64
    if ($LASTEXITCODE -ne 0) { throw 'Publish failed.' }
    foreach ($document in @('README.md', 'LICENSE', 'THIRD-PARTY-NOTICES.md')) {
        Copy-Item -LiteralPath (Join-Path $root $document) -Destination (Join-Path $root "publish/win-x64/$document") -Force
    }
    foreach ($folder in @('docs', 'samples')) {
        Copy-Item -LiteralPath (Join-Path $root $folder) -Destination (Join-Path $root 'publish/win-x64') -Recurse -Force
    }
    Write-Output (Join-Path $root 'publish/win-x64/SharePointDiff.exe')
} finally { Pop-Location }
