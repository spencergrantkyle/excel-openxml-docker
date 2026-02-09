# Excel OpenXML Docker Test

Test project combining Docker development workflow with OpenXML Excel editing.

## Structure

```
workbook/              # Unzipped Excel file (XML)
├── [Content_Types].xml
├── _rels/
├── xl/
│   ├── workbook.xml
│   ├── sharedStrings.xml
│   ├── worksheets/
│   │   ├── sheet1.xml
│   │   └── ...
│   └── _rels/
└── docProps/
```

## Workflow

1. Unzip your `.xlsx` file: `unzip myfile.xlsx -d workbook/`
2. Push the XML to this repo
3. Docker container mounts the workbook directory
4. Python scripts edit the XML
5. Repack to `.xlsx`: `cd workbook && zip -r ../output.xlsx .`

## Quick Start

```bash
docker compose up
docker compose exec app python scripts/analyze.py
```
