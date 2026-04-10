param(
    [string]$Configuration = "Release",
    [string]$RuntimeIdentifier = "win-x64"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$payloadDir = Join-Path $repoRoot "release-assets\win-x64"
$projectPath = Join-Path $repoRoot "AudioRoute.csproj"
$projectXml = [xml](Get-Content -LiteralPath $projectPath)
$targetFramework = $projectXml.Project.PropertyGroup.TargetFramework | Select-Object -First 1

if ([string]::IsNullOrWhiteSpace($targetFramework)) {
    throw "Unable to resolve TargetFramework from $projectPath"
}

$buildOutputDir = Join-Path $repoRoot "bin\$Configuration\$targetFramework\$RuntimeIdentifier"

$resolvedRepoRoot = (Resolve-Path $repoRoot).Path
if (Test-Path $payloadDir) {
    $resolvedPayloadDir = (Resolve-Path $payloadDir).Path
    if (-not $resolvedPayloadDir.StartsWith($resolvedRepoRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to clean unexpected path: $resolvedPayloadDir"
    }

    Get-ChildItem -LiteralPath $resolvedPayloadDir -Force | Where-Object { $_.Name -ne ".gitkeep" } | Remove-Item -Recurse -Force
}
else {
    New-Item -ItemType Directory -Path $payloadDir -Force | Out-Null
}

Get-ChildItem -LiteralPath (Join-Path $repoRoot "release-assets") -Filter "win-x64*.xbf" -Force -ErrorAction SilentlyContinue | Remove-Item -Force

Write-Host "Building release payload using default output directory"
dotnet build $projectPath -c $Configuration -r $RuntimeIdentifier

if ($LASTEXITCODE -ne 0) {
    throw "dotnet build failed with exit code $LASTEXITCODE."
}

if (-not (Test-Path $buildOutputDir)) {
    throw "Build output directory not found: $buildOutputDir"
}

Copy-Item (Join-Path $buildOutputDir "*") -Destination $payloadDir -Recurse -Force

Write-Host "Release payload updated: $payloadDir"
