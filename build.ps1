# build.ps1 - Generate ResolutionMonitor.vbs with embedded PowerShell script

$ErrorActionPreference = "Stop"

$sourceFile = Join-Path $PSScriptRoot "ResolutionMonitor.ps1"
$outputFile = Join-Path $PSScriptRoot "ResolutionMonitor.vbs"

if (-not (Test-Path $sourceFile)) {
    Write-Error "Source file not found: $sourceFile"
    exit 1
}

Write-Host "Reading source PowerShell script..."
$psScript = Get-Content $sourceFile -Raw

# Escape quotes for VBS string (double every quote)
$escaped = $psScript -replace '"', '""'

# Remove any trailing whitespace/newlines to keep it clean
$escaped = $escaped.TrimEnd()

# Generate VBS wrapper with embedded PowerShell
$vbsContent = @"
CreateObject("Wscript.Shell").Run "powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command ""$escaped""", 0, False
"@

Write-Host "Writing VBS file: $outputFile"
Set-Content -Path $outputFile -Value $vbsContent -Encoding ASCII

Write-Host "Build complete! Generated: $outputFile" -ForegroundColor Green
Write-Host "File size: $((Get-Item $outputFile).Length) bytes"
