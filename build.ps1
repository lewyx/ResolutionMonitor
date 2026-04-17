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

# Split into lines, escape quotes for VBS (double every quote), then
# join as VBS string with vbCrLf between lines. VBS doesn't support
# multi-line string literals, so each line must be a separate "..." literal.
$psLines = $psScript -split "`r?`n"

# Build VBS string expression: "line1" & vbCrLf & "line2" & vbCrLf & ...
# Use _ for line continuation in VBS to keep it readable
$vbsStringParts = [System.Collections.ArrayList]::new()
foreach ($line in $psLines) {
    $escapedLine = $line -replace '"', '""'
    [void]$vbsStringParts.Add("""$escapedLine""")
}
$vbsConcat = $vbsStringParts -join " & vbCrLf & _`r`n    "

# Generate VBS wrapper that:
# 1. Writes the embedded PS1 source to a temp file
# 2. Launches PowerShell with -File (avoids command-line length limit)
$vbsContent = @"
Dim fso, tmp, f
Set fso = CreateObject("Scripting.FileSystemObject")
tmp = fso.BuildPath(fso.GetSpecialFolder(2), "ResMonitor_" & fso.GetTempName & ".ps1")
Set f = fso.CreateTextFile(tmp, True)
f.Write $vbsConcat
f.Close
CreateObject("Wscript.Shell").Run "powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & tmp & """", 0, False
"@

Write-Host "Writing VBS file: $outputFile"
Set-Content -Path $outputFile -Value $vbsContent -Encoding ASCII

Write-Host "Build complete! Generated: $outputFile" -ForegroundColor Green
Write-Host "File size: $((Get-Item $outputFile).Length) bytes"
