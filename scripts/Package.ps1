#Requires -Version 7.0
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
[xml]$project = Get-Content -LiteralPath (Join-Path $root 'src/SharePointDiff.App/SharePointDiff.App.csproj') -Raw
$version = [string]$project.Project.PropertyGroup.Version
if ($version -notmatch '^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$') { throw 'Invalid release version.' }

# Explicit allowlist: never include settings, caches, project records or unrelated publish output.
$files = @('SharePointDiff.exe', 'README.md', 'LICENSE', 'THIRD-PARTY-NOTICES.md', 'docs/Verification.md', 'samples/settings.example.json', "docs/releases/v$version.md")
$notes = Join-Path $root "docs/releases/v$version.md"
if (-not (Test-Path -LiteralPath $notes -PathType Leaf)) { throw "Release notes are missing: $notes" }
& (Join-Path $PSScriptRoot 'Publish.ps1')

$published = Join-Path $root 'publish/win-x64'
foreach ($relative in $files) {
    if (-not (Test-Path -LiteralPath (Join-Path $published $relative) -PathType Leaf)) { throw "Package file is missing: $relative" }
}
$output = Join-Path $root "dist/v$version"
[IO.Directory]::CreateDirectory($output) | Out-Null
$name = "SharePointDiff-$version-win-x64.zip"
$destination = Join-Path $output $name
if (Test-Path -LiteralPath $destination) { throw "Package already exists; review it before rebuilding: $destination" }
$temporary = Join-Path $output ('.package-' + [Guid]::NewGuid().ToString('N') + '.tmp')

try {
    $archive = [IO.Compression.ZipFile]::Open($temporary, [IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($relative in $files) {
            [IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, (Join-Path $published $relative), $relative, [IO.Compression.CompressionLevel]::SmallestSize) | Out-Null
        }
    } finally { $archive.Dispose() }

    $archive = [IO.Compression.ZipFile]::OpenRead($temporary)
    try {
        if ($archive.Entries.Count -ne $files.Count) { throw 'Unexpected package entry count.' }
        $entry = $archive.GetEntry('SharePointDiff.exe')
        $stream = $entry.Open()
        try { $insideHash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($stream)) }
        finally { $stream.Dispose() }
        if ($insideHash -ne (Get-FileHash -LiteralPath (Join-Path $published 'SharePointDiff.exe') -Algorithm SHA256).Hash) { throw 'Packaged executable does not match the published build.' }
    } finally { $archive.Dispose() }

    [IO.File]::Move($temporary, $destination)
    $hash = (Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash.ToLowerInvariant()
    [IO.File]::WriteAllText((Join-Path $output 'SHA256SUMS.txt'), "$hash  $name`r`n", [Text.UTF8Encoding]::new($false))
    Write-Output $destination
    Write-Output (Join-Path $output 'SHA256SUMS.txt')
} finally {
    if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force }
}
