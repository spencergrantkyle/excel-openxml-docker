# Excel OpenXML Scripts — macOS

Scripts for extracting and repacking Excel files for programmatic XML editing.

## Quick Start

```bash
# Make scripts executable (one-time)
chmod +x extract-xlsx.sh repack-xlsx.sh

# Extract Excel to XML
./extract-xlsx.sh input.xlsx

# Edit the XML files...

# Repack back to Excel
./repack-xlsx.sh input_xml/ output.xlsx
```

## Scripts

### extract-xlsx.sh

Extracts an Excel file to a directory of XML files.

```bash
./extract-xlsx.sh <input.xlsx> [output_dir]
```

**Arguments:**
- `input.xlsx` — The Excel file to extract (required)
- `output_dir` — Where to put the XML (default: `<filename>_xml/`)

**Examples:**
```bash
./extract-xlsx.sh report.xlsx
# Creates: report_xml/

./extract-xlsx.sh report.xlsx ./workbook
# Creates: ./workbook/
```

### repack-xlsx.sh

Repacks a directory of XML files back into an Excel file.

```bash
./repack-xlsx.sh <input_dir> <output.xlsx>
```

**Arguments:**
- `input_dir` — The directory containing Excel XML (required)
- `output.xlsx` — The output Excel file (required)

**Examples:**
```bash
./repack-xlsx.sh report_xml/ report_modified.xlsx
./repack-xlsx.sh ./workbook output.xlsx
```

## Claude Code Integration

Add this to your project's `CLAUDE.md` or prompt:

```markdown
## Excel File Editing

To edit Excel files programmatically:

1. Extract: `./scripts/mac/extract-xlsx.sh input.xlsx ./workbook`
2. Edit XML files in `./workbook/xl/`
3. Repack: `./scripts/mac/repack-xlsx.sh ./workbook output.xlsx`

Key XML files:
- `xl/workbook.xml` — Sheet names, named ranges
- `xl/worksheets/sheet1.xml` — Cell data (first sheet)
- `xl/sharedStrings.xml` — Text content

NEVER use openpyxl to save complex workbooks. Use direct XML editing only.
```

## Excel XML Structure

After extraction, you'll have:

```
workbook/
├── [Content_Types].xml
├── _rels/
│   └── .rels
├── docProps/
│   ├── app.xml
│   └── core.xml
└── xl/
    ├── workbook.xml          ← Sheet list, named ranges
    ├── sharedStrings.xml     ← All text content
    ├── styles.xml            ← Formatting
    ├── worksheets/
    │   ├── sheet1.xml        ← First sheet cells
    │   ├── sheet2.xml
    │   └── ...
    └── _rels/
        └── workbook.xml.rels ← Sheet relationships
```

## Common Edits

### Change Cell Value (worksheet XML)

Find in `xl/worksheets/sheet1.xml`:
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

**"Permission denied"**
```bash
chmod +x extract-xlsx.sh repack-xlsx.sh
```

**Corrupted output file**
- Ensure no `.DS_Store` files (script excludes these)
- Check that all XML is valid (no unclosed tags)
- Verify `[Content_Types].xml` exists

**Excel won't open the file**
- Try: `unzip -t output.xlsx` to check archive integrity
- Open in text editor to check for XML errors
