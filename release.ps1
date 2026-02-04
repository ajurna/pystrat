$ErrorActionPreference = "Stop"

$version = & uv version --short
$version = $version.Trim()
if (-not $version) {
    throw "Could not read version from pyproject.toml"
}

$tag = "v$version"

& uv run -- python build_exe.py

$exePath = Join-Path $PSScriptRoot "dist\\pystrat.exe"
if (-not (Test-Path $exePath)) {
    throw "Build output not found: $exePath"
}

$zipPath = Join-Path $PSScriptRoot "dist\\pystrat-$version.zip"
if (Test-Path $zipPath) {
    Remove-Item $zipPath -Force
}
Compress-Archive -Path $exePath -DestinationPath $zipPath

& git tag $tag
& git push origin $tag
& gh release create $tag $exePath $zipPath -t $tag -F "$PSScriptRoot\\RELEASE.md"
