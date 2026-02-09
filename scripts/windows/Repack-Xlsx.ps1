<#
.SYNOPSIS
    Repack XML directory back to Excel (.xlsx)

.DESCRIPTION
    Repacks a directory of XML files back into an Excel file.

.PARAMETER InputDir
    Directory containing the extracted Excel XML

.PARAMETER OutputFile
    Path for the output .xlsx file

.EXAMPLE
    .\Repack-Xlsx.ps1 -InputDir .\workbook -OutputFile output.xlsx

.EXAMPLE
    .\Repack-Xlsx.ps1 -InputDir .\report_xml -OutputFile report_modified.xlsx
#>

param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$InputDir,
    
    [Parameter(Mandatory=$true, Position=1)]
    [string]$OutputFile
)

# Validate input directory exists
if (-not (Test-Path $InputDir)) {
    Write-Host "Error: Directory not found: $InputDir" -ForegroundColor Red
    exit 1
}

# Validate it contains Excel XML structure
$contentTypesPath = Join-Path $InputDir "[Content_Types].xml"
if (-not (Test-Path $contentTypesPath)) {
    Write-Host "Error: Not a valid Excel XML directory (missing [Content_Types].xml)" -ForegroundColor Red
    exit 1
}

# Check if output file exists
if (Test-Path $OutputFile) {
    Write-Host "Warning: Output file already exists: $OutputFile" -ForegroundColor Yellow
    $confirm = Read-Host "Overwrite? (y/n)"
    if ($confirm -ne "y" -and $confirm -ne "Y") {
        Write-Host "Aborted."
        exit 1
    }
    Remove-Item -Path $OutputFile -Force
}

# Get absolute paths
$InputDirAbs = (Resolve-Path $InputDir).Path

# Ensure output directory exists and get absolute path
$OutputFileDir = Split-Path -Parent $OutputFile
if ($OutputFileDir -and -not (Test-Path $OutputFileDir)) {
    New-Item -ItemType Directory -Path $OutputFileDir -Force | Out-Null
}
if ($OutputFileDir) {
    $OutputFileAbs = Join-Path (Resolve-Path $OutputFileDir).Path (Split-Path -Leaf $OutputFile)
} else {
    $OutputFileAbs = Join-Path (Get-Location).Path $OutputFile
}

# Create xlsx using .NET
Write-Host "Repacking: $InputDir" -ForegroundColor Green
Write-Host "To: $OutputFile" -ForegroundColor Green

try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    
    # Create the zip file
    [System.IO.Compression.ZipFile]::CreateFromDirectory(
        $InputDirAbs,
        $OutputFileAbs,
        [System.IO.Compression.CompressionLevel]::Optimal,
        $false  # Don't include base directory name
    )
}
catch {
    Write-Host "Error during repacking: $_" -ForegroundColor Red
    exit 1
}

# Verify
if (Test-Path $OutputFile) {
    $fileInfo = Get-Item $OutputFile
    $fileSize = "{0:N2} KB" -f ($fileInfo.Length / 1KB)
    
    Write-Host ""
    Write-Host "Repack complete" -ForegroundColor Green
    Write-Host "  Output: $OutputFile"
    Write-Host "  Size: $fileSize"
}
else {
    Write-Host "Error: Failed to create output file" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Tip: Open in Excel to verify the file works correctly." -ForegroundColor Yellow
