$ErrorActionPreference = 'Stop'
& (Join-Path $PSScriptRoot 'build.ps1')
$exe = Join-Path $PSScriptRoot 'dist\DualSlide.exe'
$metadata = [Diagnostics.FileVersionInfo]::GetVersionInfo($exe)
$version = '{0}.{1}.{2}' -f $metadata.FileMajorPart, $metadata.FileMinorPart, $metadata.FileBuildPart
$staging = Join-Path $PSScriptRoot ('.build\package-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path (Join-Path $staging 'assets') -Force | Out-Null
Copy-Item -LiteralPath $exe -Destination $staging
foreach ($name in @('README.md', 'USER_GUIDE.md', 'IMPLEMENTATION_PLAN.md', 'VERIFICATION.md', 'RELEASE_NOTES.md')) {
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination $staging
}
foreach ($name in @('DualSlide.svg', 'DualSlide.ico')) {
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot "assets\$name") -Destination (Join-Path $staging 'assets')
}
$archive = Join-Path $PSScriptRoot "dist\DualSlide-$version.zip"
Compress-Archive -Path (Join-Path $staging '*') -DestinationPath $archive -Force
$hash = (Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash.ToLowerInvariant()
[IO.File]::WriteAllText("$archive.sha256", "$hash  DualSlide-$version.zip`n", [Text.Encoding]::ASCII)
Write-Output "Created $archive and its SHA-256 checksum."
