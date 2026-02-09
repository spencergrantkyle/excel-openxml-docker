<#
.SYNOPSIS
    Extract Excel (.xlsx) to XML for editing

.DESCRIPTION
    Extracts an Excel file to a directory of XML files for programmatic editing.

.PARAMETER InputFile
    Path to the .xlsx file to extract

.PARAMETER OutputDir
    Directory to extract to (default: <filename>_xml)

.EXAMPLE
    .\Extract-Xlsx.ps1 -InputFile report.xlsx

.EXAMPLE
    .\Extract-Xlsx.ps1 -InputFile report.xlsx -OutputDir .\workbook
#>

param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$InputFile,
    
    [Parameter(Mandatory=$false, Position=1)]
    [string]$OutputDir
)

# Set default output directory
if (-not $OutputDir) {
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)
    $OutputDir = "${baseName}_xml"
}

# Validate input file exists
if (-not (Test-Path $InputFile)) {
    Write-Host "Error: File not found: $InputFile" -ForegroundColor Red
    exit 1
}

# Validate it's an xlsx file
if (-not $InputFile.EndsWith(".xlsx")) {
    Write-Host "Warning: File doesn't have .xlsx extension" -ForegroundColor Yellow
}

# Check if output directory exists
if (Test-Path $OutputDir) {
    Write-Host "Warning: Output directory already exists: $OutputDir" -ForegroundColor Yellow
    $confirm = Read-Host "Overwrite? (y/n)"
    if ($confirm -ne "y" -and $confirm -ne "Y") {
        Write-Host "Aborted."
        exit 1
    }
    Remove-Item -Path $OutputDir -Recurse -Force
}

# Create output directory
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

# Get absolute paths
$InputFileAbs = (Resolve-Path $InputFile).Path
$OutputDirAbs = (Resolve-Path $OutputDir).Path

# Extract using .NET
Write-Host "Extracting: $InputFile" -ForegroundColor Green
Write-Host "To: $OutputDir" -ForegroundColor Green

try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($InputFileAbs, $OutputDirAbs)
}
catch {
    Write-Host "Error during extraction: $_" -ForegroundColor Red
    exit 1
}

# Show structure
Write-Host ""
Write-Host "Extraction complete" -ForegroundColor Green
Write-Host ""
Write-Host "Structure:"

Get-ChildItem -Path $OutputDir -Recurse -File | Select-Object -First 20 | ForEach-Object {
    Write-Host "  $($_.FullName.Replace($OutputDirAbs, ''))"
}

# Count files
$fileCount = (Get-ChildItem -Path $OutputDir -Recurse -File).Count
Write-Host ""
Write-Host "Total files: $fileCount" -ForegroundColor Green

# Reminder
Write-Host ""
Write-Host "Key files for editing:" -ForegroundColor Yellow
Write-Host "  $OutputDir\xl\workbook.xml       - Sheet names, named ranges"
Write-Host "  $OutputDir\xl\worksheets\*.xml   - Cell data and formulas"
Write-Host "  $OutputDir\xl\sharedStrings.xml  - Text content"
Write-Host ""
Write-Host "To repack: .\Repack-Xlsx.ps1 -InputDir $OutputDir -OutputFile output.xlsx" -ForegroundColor Yellow
