param([string]$MakeAppx = '', [string]$SourceRoot = $PSScriptRoot)
$ErrorActionPreference = 'Stop'
if (-not $MakeAppx) {
    $command = Get-Command MakeAppx.exe -ErrorAction SilentlyContinue
    if ($command) { $MakeAppx = $command.Source }
    else {
        $sdk = Join-Path ${env:ProgramFiles(x86)} 'Windows Kits\10\bin'
        $tool = Get-ChildItem -LiteralPath $sdk -Directory | Sort-Object Name -Descending |
            ForEach-Object { Join-Path $_.FullName 'x64\makeappx.exe' } | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
        if (-not $tool) { throw 'MakeAppx.exe is required (install the Windows SDK).' }
        $MakeAppx = $tool
    }
}
$SourceRoot = (Resolve-Path -LiteralPath $SourceRoot).Path
& (Join-Path $SourceRoot 'build.ps1')
[xml]$manifest = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'store\AppxManifest.xml')
$version = [Diagnostics.FileVersionInfo]::GetVersionInfo((Join-Path $SourceRoot 'dist\DualSlide.exe')).FileVersion
$manifest.Package.Identity.Version = $version
New-Item -ItemType Directory -Force -Path (Join-Path $PSScriptRoot 'dist') | Out-Null
$staging = Join-Path $PSScriptRoot ('.build\msix-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $staging | Out-Null
Copy-Item -LiteralPath (Join-Path $SourceRoot 'dist\DualSlide.exe') -Destination $staging
$manifest.Save((Join-Path $staging 'AppxManifest.xml'))
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'store\Assets') -Destination $staging -Recurse
foreach ($name in @('README.md','USER_GUIDE.md')) { Copy-Item -LiteralPath (Join-Path $SourceRoot $name) -Destination $staging }
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'PRIVACY.md') -Destination $staging
$package = Join-Path $PSScriptRoot "dist\DualSlide-$version-x64.msix"
& $MakeAppx pack /d $staging /p $package /o
if ($LASTEXITCODE -ne 0) { throw 'MSIX validation/packaging failed.' }
$hash = (Get-FileHash -LiteralPath $package -Algorithm SHA256).Hash.ToLowerInvariant()
[IO.File]::WriteAllText("$package.sha256", "$hash  $([IO.Path]::GetFileName($package))`n", [Text.Encoding]::ASCII)
Write-Output "Store submission package: $package (unsigned; Microsoft Store signs after certification)."
