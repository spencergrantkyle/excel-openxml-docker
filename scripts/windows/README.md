# Excel OpenXML Scripts — Windows

PowerShell scripts for extracting and repacking Excel files for programmatic XML editing.

## Quick Start

```powershell
# Extract Excel to XML
.\Extract-Xlsx.ps1 -InputFile input.xlsx

# Edit the XML files...

# Repack back to Excel
.\Repack-Xlsx.ps1 -InputDir input_xml -OutputFile output.xlsx
```

## Scripts

### Extract-Xlsx.ps1

Extracts an Excel file to a directory of XML files.

```powershell
.\Extract-Xlsx.ps1 -InputFile <path> [-OutputDir <path>]
```

**Parameters:**
- `-InputFile` — The Excel file to extract (required)
- `-OutputDir` — Where to put the XML (default: `<filename>_xml\`)

**Examples:**
```powershell
.\Extract-Xlsx.ps1 -InputFile report.xlsx
# Creates: report_xml\

.\Extract-Xlsx.ps1 -InputFile report.xlsx -OutputDir .\workbook
# Creates: .\workbook\
```

### Repack-Xlsx.ps1

Repacks a directory of XML files back into an Excel file.

```powershell
.\Repack-Xlsx.ps1 -InputDir <path> -OutputFile <path>
```

**Parameters:**
- `-InputDir` — The directory containing Excel XML (required)
- `-OutputFile` — The output Excel file (required)

**Examples:**
```powershell
.\Repack-Xlsx.ps1 -InputDir report_xml -OutputFile report_modified.xlsx
.\Repack-Xlsx.ps1 -InputDir .\workbook -OutputFile output.xlsx
```

## Execution Policy

If you get an execution policy error, run:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

Or run the script with bypass:

```powershell
powershell -ExecutionPolicy Bypass -File .\Extract-Xlsx.ps1 -InputFile input.xlsx
```

## Claude Code Integration

Add this to your project's `CLAUDE.md` or prompt:

```markdown
## Excel File Editing

To edit Excel files programmatically:

1. Extract: `.\scripts\windows\Extract-Xlsx.ps1 -InputFile input.xlsx -OutputDir .\workbook`
2. Edit XML files in `.\workbook\xl\`
3. Repack: `.\scripts\windows\Repack-Xlsx.ps1 -InputDir .\workbook -OutputFile output.xlsx`

Key XML files:
- `xl\workbook.xml` — Sheet names, named ranges
- `xl\worksheets\sheet1.xml` — Cell data (first sheet)
- `xl\sharedStrings.xml` — Text content

NEVER use openpyxl to save complex workbooks. Use direct XML editing only.
```

## Excel XML Structure

After extraction, you'll have:

```
workbook\
├── [Content_Types].xml
├── _rels\
│   └── .rels
├── docProps\
│   ├── app.xml
│   └── core.xml
└── xl\
    ├── workbook.xml          ← Sheet list, named ranges
    ├── sharedStrings.xml     ← All text content
    ├── styles.xml            ← Formatting
    ├── worksheets\
    │   ├── sheet1.xml        ← First sheet cells
    │   ├── sheet2.xml
    │   └── ...
    └── _rels\
        └── workbook.xml.rels ← Sheet relationships
```

## Common Edits

### Change Cell Value (worksheet XML)

Find in `xl\worksheets\sheet1.xml`:
```xml
<c r="B5" t="s"><v>12</v></c>
```

- `r="B5"` — Cell reference
- `t="s"` — Type is shared string (index 12 in sharedStrings.xml)
- `<v>12</v>` — Value (or formula in `<f>` tag)

### Add Named Range (workbook.xml)

Add to `<definedNames>` section:
```xml
<definedName name="MyRange">'Sheet1'!$A$1:$A$10</definedName>
```

## Troubleshooting

**"Execution policy" error**
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

**"Access denied" error**
- Ensure you have write permissions to the output directory
- Close Excel if the file is open

**Corrupted output file**
- Check that all XML is valid (no unclosed tags)
- Verify `[Content_Types].xml` exists in the input directory

**Excel won't open the file**
- Open one of the XML files in a text editor to check for errors
- Ensure no binary files were accidentally modified
