param([Parameter(Mandatory)][string]$ReleaseTag, [string]$OutputDirectory = '.build/store-candidate')
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot/StorePublishing.ps1"
$root = Split-Path $PSScriptRoot
Set-Location $root
if ($ReleaseTag -cnotmatch '^v(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)$') { throw 'Specify an exact stable release tag, such as v0.5.4.' }
$version = $ReleaseTag.Substring(1)
$releaseJson = gh release view $ReleaseTag --repo $env:GH_REPO --json tagName,isDraft,isPrerelease
if ($LASTEXITCODE -ne 0) { throw "Published release $ReleaseTag was not found." }
$release = $releaseJson | ConvertFrom-Json
if ($release.tagName -cne $ReleaseTag -or $release.isDraft -or $release.isPrerelease) { throw 'Only a published stable release can be submitted.' }
git fetch --no-tags origin "refs/tags/${ReleaseTag}:refs/tags/${ReleaseTag}"
if ($LASTEXITCODE -ne 0) { throw 'Cannot fetch release source.' }
$sourceSha = (git rev-parse --verify "${ReleaseTag}^{commit}").Trim()
if ($LASTEXITCODE -ne 0) { throw 'Cannot resolve release source.' }
git merge-base --is-ancestor $sourceSha origin/main
if ($LASTEXITCODE -ne 0) { throw 'The release must belong to the main branch history.' }
$notes = (git show "${sourceSha}:RELEASE_NOTES.md") -join "`n"
if ($LASTEXITCODE -ne 0) { throw 'The selected release has no release notes.' }
$notes = Get-DualSlideReleaseNotes -Changelog $notes -Version $version
$shots = @(Get-DualSlideStoreScreenshots -Directory "$root/store/Screenshots" -Version $version)
$source = Join-Path $root ('.build/store-source-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $source -Force | Out-Null
$archive = "$source.zip"
git archive --format=zip "--output=$archive" $sourceSha
if ($LASTEXITCODE -ne 0) { throw 'Cannot archive exact release source.' }
Expand-Archive -LiteralPath $archive -DestinationPath $source
& "$root/build-msix.ps1" -SourceRoot $source
$package = Join-Path $root "dist/DualSlide-$version.0-x64.msix"
$null = Get-DualSlidePackageIdentity -Path $package -Version $version
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Candidate directory must be new, to prevent stale upload files.' }
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
Copy-Item -LiteralPath $package -Destination $OutputDirectory
foreach ($shot in $shots) { Copy-Item -LiteralPath $shot.path -Destination $OutputDirectory }
Copy-Item -LiteralPath "$root/store/Screenshots/screenshots.json" -Destination $OutputDirectory
Copy-Item -LiteralPath "$root/store/listing-en-US.json" -Destination $OutputDirectory
Set-Content -LiteralPath (Join-Path $OutputDirectory 'release-notes.txt') -Value $notes -Encoding utf8
$files = @(Get-ChildItem -LiteralPath $OutputDirectory -File | ForEach-Object {
    [ordered]@{ fileName=$_.Name; sha256=(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant() }
})
[ordered]@{ releaseTag=$ReleaseTag; version=$version; sourceSha=$sourceSha; productId='9NP8S6HC2252'; files=$files } |
    ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'candidate.json') -Encoding utf8
if ($env:GITHUB_OUTPUT) { "version=$version" >> $env:GITHUB_OUTPUT; "source_sha=$sourceSha" >> $env:GITHUB_OUTPUT }
if ($env:GITHUB_STEP_SUMMARY) {
    "## DualSlide Store candidate`n`n- Release: $ReleaseTag`n- Source: $sourceSha`n- Package: DualSlide-$version.0-x64.msix`n- Screenshots: $($shots.Count)`n- No Store changes made by preparation." >> $env:GITHUB_STEP_SUMMARY
}
