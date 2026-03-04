# Builds V2XAmbilight.exe as a self-contained single-file Windows executable.
# Requires .NET 8 SDK.
$ErrorActionPreference = "Stop"

dotnet publish "$PSScriptRoot\V2XAmbilight.csproj" `
    -c Release `
    -r win-x64 `
    --self-contained `
    -p:PublishSingleFile=true `
    -p:PublishReadyToRun=true `
    -o "$PSScriptRoot\dist"

Write-Host "Built: $PSScriptRoot\dist\V2XAmbilight.exe" -ForegroundColor Green
