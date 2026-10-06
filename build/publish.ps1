<#
.SYNOPSIS
    Builds self-contained single-file Windows executables for the app and the CLI and zips them.
.EXAMPLE
    ./build/publish.ps1 -Version v3.0.0
#>
param(
    [string]$Version = "0.0.0-dev",
    [string]$Runtime = "win-x64"
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$out = Join-Path $root "artifacts/publish"
$semver = $Version.TrimStart('v')
$numeric = ($semver -split '-')[0]
Remove-Item $out -Recurse -Force -ErrorAction SilentlyContinue

$common = @(
    '-c', 'Release', '-r', $Runtime, '--self-contained',
    '-p:PublishSingleFile=true', '-p:IncludeNativeLibrariesForSelfExtract=true',
    '-p:EnableCompressionInSingleFile=true', '-p:DebugType=none',
    "-p:Version=$numeric", "-p:InformationalVersion=$semver"
)
dotnet publish (Join-Path $root 'src/HypervisorExplorer.App') @common -o (Join-Path $out 'app')
dotnet publish (Join-Path $root 'src/HypervisorExplorer.Cli') @common -o (Join-Path $out 'app')
Get-ChildItem (Join-Path $out 'app') -Filter *.pdb | Remove-Item

$zip = Join-Path $out "HypervisorExplorer-$semver-$Runtime.zip"
Compress-Archive -Path (Join-Path $out 'app/*') -DestinationPath $zip
Write-Host "Created $zip"
