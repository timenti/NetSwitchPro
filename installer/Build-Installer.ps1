param(
    [Parameter(Mandatory)] [string]$Compiler,
    [string]$Dotnet = 'dotnet',
    [string]$Output = (Join-Path $PSScriptRoot '../dist')
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$build = Join-Path $env:TEMP ('NetSwitch-build-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $Output -Force | Out-Null
$outputPath = (Resolve-Path -LiteralPath $Output).Path

& $Dotnet publish (Join-Path $root 'NetworkAdapterSwitcher.csproj') -c Release -r win-x64 --self-contained true -o $build
if ($LASTEXITCODE -ne 0) { throw 'NetSwitch Pro build failed.' }
$exe = Get-Item -LiteralPath (Join-Path $build 'NetworkAdapterSwitcher.exe')
if ($exe.VersionInfo.ProductName -ne 'NetSwitch Pro' -or $exe.VersionInfo.FileVersion -notmatch '^\d+\.\d+\.\d+\.0$') {
    throw 'The published executable has unexpected product metadata.'
}
$fileVersion = [version]$exe.VersionInfo.FileVersion
if ($fileVersion.Minor -gt 9 -or $fileVersion.Build -gt 9) {
    throw 'Minor and patch must be single digits. Carry to the next component after 9.'
}
$version = $fileVersion.ToString(3)
$display = $version -replace '(\.0){1,2}$', ''
& $Compiler "/DAppVersion=$version" "/DDisplayVersion=$display" "/DPayload=$($exe.FullName)" "/O$outputPath" (Join-Path $PSScriptRoot 'NetSwitchPro.iss')
if ($LASTEXITCODE -ne 0) { throw 'NetSwitch Pro installer build failed.' }
$asset = "NetSwitchPro-$version.exe"
$hash = (Get-FileHash -LiteralPath (Join-Path $outputPath $asset) -Algorithm SHA256).Hash.ToLowerInvariant()
Set-Content -LiteralPath (Join-Path $outputPath 'SHA256SUMS.txt') -Value "$hash  $asset" -Encoding ascii
