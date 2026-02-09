# Excel OpenXML Docker Test

Test project combining Docker development workflow with OpenXML Excel editing.

## Quick Start

### macOS

```bash
# Make scripts executable (one-time)
chmod +x scripts/mac/*.sh

# Extract Excel to XML
./scripts/mac/extract-xlsx.sh myfile.xlsx ./workbook

# Edit XML files in ./workbook/xl/...

# Repack back to Excel
./scripts/mac/repack-xlsx.sh ./workbook output.xlsx
```

### Windows

```powershell
# Extract Excel to XML
.\scripts\windows\Extract-Xlsx.ps1 -InputFile myfile.xlsx -OutputDir .\workbook

# Edit XML files in .\workbook\xl\...

# Repack back to Excel
.\scripts\windows\Repack-Xlsx.ps1 -InputDir .\workbook -OutputFile output.xlsx
```

## Scripts

| Platform | Extract | Repack |
|----------|---------|--------|
| macOS | `scripts/mac/extract-xlsx.sh` | `scripts/mac/repack-xlsx.sh` |
| Windows | `scripts/windows/Extract-Xlsx.ps1` | `scripts/windows/Repack-Xlsx.ps1` |

See the README in each platform folder for detailed usage.

## Excel XML Structure

After extraction:

```
workbook/
├── [Content_Types].xml
├── _rels/
│   └── .rels
├── docProps/
│   ├── app.xml
│   └── core.xml
└── xl/
    ├── workbook.xml          ← Sheet names, named ranges
    ├── sharedStrings.xml     ← All text content
    ├── styles.xml            ← Formatting
    ├── worksheets/
    │   ├── sheet1.xml        ← First sheet cells
    │   ├── sheet2.xml
    │   └── ...
    └── _rels/
        └── workbook.xml.rels ← Sheet relationships
```

## Workflow

```
┌─────────────┐     ┌─────────────┐     ┌─────────────┐
│   .xlsx     │ ──▶ │    XML      │ ──▶ │   .xlsx     │
│  (input)    │     │  (edit)     │     │  (output)   │
└─────────────┘     └─────────────┘     └─────────────┘
    extract            Python/          repack
                      manual edit
```

## Claude Code Integration

Add to your project's `CLAUDE.md`:

```markdown
## Excel File Editing

To edit Excel files programmatically:

### macOS
1. Extract: `./scripts/mac/extract-xlsx.sh input.xlsx ./workbook`
2. Edit XML files in `./workbook/xl/`
3. Repack: `./scripts/mac/repack-xlsx.sh ./workbook output.xlsx`

### Windows
1. Extract: `.\scripts\windows\Extract-Xlsx.ps1 -InputFile input.xlsx -OutputDir .\workbook`
2. Edit XML files in `.\workbook\xl\`
3. Repack: `.\scripts\windows\Repack-Xlsx.ps1 -InputDir .\workbook -OutputFile output.xlsx`

Key XML files:
- `xl/workbook.xml` — Sheet names, named ranges
- `xl/worksheets/sheet1.xml` — Cell data (first sheet)
- `xl/sharedStrings.xml` — Text content

NEVER use openpyxl to save complex workbooks. Use direct XML editing only.
```

## Key Rules for Editing

1. **Never use openpyxl/pandas to SAVE** — They corrupt complex formulas
2. **Read-only analysis is OK** — `openpyxl.load_workbook(file, data_only=True)`
3. **Edit XML directly** — Use Python's `xml.etree.ElementTree`
4. **Always work on a copy** — Keep the original safe
5. **Verify after repacking** — Open in Excel to confirm it works

## Docker Integration (Coming Soon)

Docker container for Python-based XML editing workflow.
