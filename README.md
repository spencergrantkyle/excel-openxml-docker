# Excel OpenXML Toolkit

**A secure, safe tool for exploring and surgically editing complex Excel workbooks using OpenXML manipulation.**

## Table of Contents

- [Overview](#overview)
- [Security Protocol](#security-protocol)
- [Prerequisites](#prerequisites)
- [Setup Instructions](#setup-instructions)
- [How It Works](#how-it-works)
- [Usage](#usage)
- [Understanding the Output](#understanding-the-output)
- [OpenXML Surgery Workflow](#openxml-surgery-workflow)
- [Security Best Practices](#security-best-practices)
- [Troubleshooting](#troubleshooting)

---

## Overview

This toolkit provides a **safe, read-only exploration** and **surgical editing** capability for complex Excel workbooks (.xlsx files) using the OpenXML format.

### What Makes This Different

Traditional Python Excel libraries (openpyxl, xlsxwriter, pandas) **reconstruct** Excel files when saving, which:
- ❌ Destroys complex shared formulas
- ❌ Replaces formulas with cached values
- ❌ Corrupts cross-sheet references
- ❌ Breaks conditional formatting and data bindings

**This toolkit uses OpenXML (ZIP + XML manipulation)**, which:
- ✅ Preserves all formulas exactly as-is
- ✅ Makes surgical edits to specific XML elements
- ✅ Maintains all workbook integrity
- ✅ Suitable for production-critical templates

### Use Cases

- **Exploring** complex Excel templates with thousands of named ranges
- **Adding** new named ranges without touching existing formulas
- **Analyzing** workbook structure (sheets, strings, styles)
- **Verifying** round-trip integrity (unzip → rezip)
- **Working with** encrypted Excel files (CDFV2 encryption)

---

## Security Protocol

### Password Protection Model

This toolkit implements a **zero-knowledge password model** where:

1. **Passwords are stored locally** in a `.env` file that is **never committed to git**
2. **Scripts read from environment variables** using `python-dotenv`
3. **AI assistants can run scripts** without seeing the actual password value
4. **You configure once** and the password is automatically used for all operations

### How It Keeps Passwords Secure

```
┌─────────────────────────────────────────────────────────┐
│                    Your Local Machine                   │
│                                                           │
│  .env file                                               │
│  ┌──────────────────────────────────┐                   │
│  │ EXCEL_PASSWORD=your_secret_here  │  ← NEVER in git   │
│  └──────────────────────────────────┘                   │
│                    ↓                                      │
│  Python Script (explore_excel_xml.py)                    │
│  ┌──────────────────────────────────┐                   │
│  │ load_dotenv()                    │                   │
│  │ pw = os.getenv('EXCEL_PASSWORD') │  ← Reads from env │
│  │ decrypt_file(password=pw)        │                   │
│  └──────────────────────────────────┘                   │
│                    ↓                                      │
│  Decrypted Excel → XML → Analysis                        │
└─────────────────────────────────────────────────────────┘

                          ↓ git commit

┌─────────────────────────────────────────────────────────┐
│                   Git Repository                        │
│                                                           │
│  ✅ explore_excel_xml.py   (contains no secrets)         │
│  ✅ README.md             (this file)                    │
│  ✅ .env.example          (template only)                │
│  ✅ .gitignore            (blocks .env)                  │
│                                                           │
│  ❌ .env                  (EXCLUDED by .gitignore)        │
│  ❌ *_DECRYPTED.xlsx      (EXCLUDED by .gitignore)        │
└─────────────────────────────────────────────────────────┘
```

### What AI Assistants See

When an AI assistant (like Claude Code) runs the script:

**❌ AI CANNOT see:**
- The actual password value
- Contents of `.env` file
- Decrypted workbook contents (unless explicitly read)

**✅ AI CAN see:**
- Script output: "Attempting decryption using EXCEL_PASSWORD from environment"
- Analysis results: number of sheets, named ranges, structure
- Error messages if password is wrong or missing

**✅ AI CAN do:**
- Run the script automatically
- Read the XML structure
- Make surgical edits using OpenXML
- Generate reports and analysis

---

## Prerequisites

### Required Software

- **Python 3.7+**
- **pip** (Python package manager)

### Required Python Packages

```bash
msoffcrypto-tool   # For decrypting password-protected Excel files
python-dotenv      # For loading environment variables from .env file
```

These will be installed during setup.

---

## Setup Instructions

### Step 1: Create Virtual Environment (Recommended)

```bash
# Create virtual environment
python3 -m venv venv

# Activate it
# On Windows (Git Bash/MINGW):
source venv/Scripts/activate

# On macOS/Linux:
source venv/bin/activate
```

### Step 2: Install Dependencies

```bash
pip install msoffcrypto-tool python-dotenv
```

### Step 3: Configure Password

Create a `.env` file in your project root (or wherever you run the script from):

```bash
# Create .env file
cat > .env << 'EOF'
# Excel Workbook Encryption Password
EXCEL_PASSWORD=your_actual_password_here
EOF
```

**Important:** Replace `your_actual_password_here` with your actual Excel password.

### Step 4: Protect Your .env File

Add to `.gitignore`:

```bash
echo ".env" >> .gitignore
echo "*_DECRYPTED.xlsx" >> .gitignore
echo "*_REZIPPED.xlsx" >> .gitignore
echo "workbook_xml/" >> .gitignore
```

### Step 5: Verify Setup

```bash
# Check that .env is ignored
git check-ignore .env
# Should output: .env

# Check that .env exists
ls -la .env
# Should show the file
```

---

## How It Works

### The .xlsx File Format

Excel files (.xlsx) are actually **ZIP archives** containing XML files:

```
MyWorkbook.xlsx (ZIP archive)
├── [Content_Types].xml          # MIME types
├── _rels/
│   └── .rels                    # Relationships
└── xl/
    ├── workbook.xml             # Sheet list, named ranges ⭐
    ├── sharedStrings.xml        # All text strings
    ├── styles.xml               # Formatting
    ├── worksheets/
    │   ├── sheet1.xml           # Sheet 1 data
    │   ├── sheet2.xml           # Sheet 2 data
    │   └── ...
    └── ...
```

### Two-Phase Workflow

#### Phase 1: Read-Only Exploration (Safe)

```python
# 1. Decrypt if needed
decrypted_file = decrypt_if_needed("MyWorkbook.xlsx")

# 2. Unzip to XML directory
unzip_excel(decrypted_file, "workbook_xml/")

# 3. Read and analyze (openpyxl with data_only=True)
# - Count sheets, named ranges
# - Scan structure
# - Generate reports

# 4. Close without saving (NO changes made)
```

**Why this is safe:**
- Original file is never modified
- `data_only=True` means read-only mode
- No save operation occurs

#### Phase 2: Surgical OpenXML Edit (Targeted)

```python
# 1. Copy original file byte-for-byte
shutil.copy2("original.xlsx", "modified.xlsx")

# 2. Open as ZIP archive
with ZipFile("modified.xlsx", "r") as zip_in:
    workbook_xml = zip_in.read("xl/workbook.xml")

# 3. Parse specific XML file
root = ET.fromstring(workbook_xml)

# 4. Make minimal, targeted edit
# Example: Add a new named range
defined_names = root.find("{...}definedNames")
new_range = ET.SubElement(defined_names, "{...}definedName")
new_range.set("name", "MyNewRange")
new_range.text = "'Sheet1'!$A$1:$A$10"

# 5. Repack ZIP with modified XML, copy everything else untouched
new_xml = ET.tostring(root, encoding="unicode")
with ZipFile("modified.xlsx", "w") as zip_out:
    for item in zip_in.namelist():
        if item == "xl/workbook.xml":
            zip_out.writestr(item, new_xml)
        else:
            zip_out.writestr(item, zip_in.read(item))

# 6. Verify by reading back
```

**Why this is safe:**
- Only the specific XML element is modified
- All other content is copied byte-for-byte
- Formulas, formatting, and bindings remain intact

---

## Usage

### Basic Usage

Place your Excel file in the same directory as `explore_excel_xml.py` and run:

```bash
python explore_excel_xml.py
```

By default, it looks for a file named:
```
Financials+_RepeatHeaderNamedRanges SPG.xlsx
```

### Custom File Path

Edit the script to change the file path:

```python
def main():
    # Change this line:
    xlsx_file = Path("YourFileName.xlsx")
    # ...
```

Or modify the script to accept command-line arguments.

---

## Understanding the Output

### Decryption Phase

```
[*] File appears to be encrypted
[*] Attempting decryption using EXCEL_PASSWORD from environment
[+] Successfully decrypted to: YourFile_DECRYPTED.xlsx
```

**What this means:**
- Script detected the file is encrypted (CDFV2)
- Used `EXCEL_PASSWORD` from your `.env` file
- Created a temporary decrypted copy
- **Note:** The password value never appears in output

### Extraction Phase

```
[*] Unzipping: YourFile_DECRYPTED.xlsx
[*] Output directory: workbook_xml
[+] Extracted 97 files and directories
```

**What this means:**
- Unzipped the .xlsx (ZIP archive) to `workbook_xml/` directory
- 97 total files = XML files, relationships, media, etc.

### Structure Analysis

```
[DIR] Directory Structure (max depth 2):

+-- _rels
|   +-- .rels
+-- xl
|   +-- _rels
|   +-- workbook.xml
|   +-- worksheets/
|       +-- sheet1.xml
|       +-- sheet2.xml
|       ...
```

**What this means:**
- Shows the internal XML structure
- `xl/workbook.xml` = most important file (contains named ranges)
- `xl/worksheets/` = individual sheet data

### Key XML Files Report

```
[1] xl/workbook.xml
   Worksheets: 66
   Named Ranges: 11325
   First 3 sheets:
      - Status (sheetId=1)
      - Cover (sheetId=2)
      - Index (sheetId=3)
   First 3 named ranges:
      - MyRange1 = 'Sheet1'!$A$1:$A$10
      - MyRange2 = 'Sheet2'!$B$5:$C$20
```

**What this means:**
- Found 66 worksheets
- Found 11,325 named ranges (this is the data you'll be working with)
- Shows sample sheets and named ranges

### Round-Trip Verification

```
[VERIFY] ROUND-TRIP VALIDATION
   Original: 12,269,588 bytes
   Rezipped: 10,279,005 bytes
   [OK] Size difference: 16.22% (acceptable)
   Original files: 88
   Rezipped files: 88
   [OK] File list matches perfectly
```

**What this means:**
- Tested unzip → rezip to verify the process works
- Size difference is normal (compression algorithms differ)
- **16% difference is acceptable** for unmodified round-trip
- File count matches = all XML files preserved

---

## OpenXML Surgery Workflow

### Example: Adding a New Named Range

This is the **safe pattern** for modifying complex Excel files:

```python
import shutil
import xml.etree.ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED
from pathlib import Path
import tempfile

# Namespace for Excel XML
SSML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

def add_named_range(input_xlsx, output_xlsx, range_name, range_ref):
    """
    Safely add a named range to an Excel file using OpenXML.

    Args:
        input_xlsx: Source Excel file (decrypted)
        range_name: Name of the new range (e.g., "MyRange")
        range_ref: Range reference (e.g., "'Sheet1'!$A$1:$A$10")
    """

    # Step 1: Copy original file
    shutil.copy2(input_xlsx, output_xlsx)

    # Step 2: Read workbook.xml from the ZIP
    with ZipFile(output_xlsx, "r") as zin:
        workbook_xml = zin.read("xl/workbook.xml")

    # Step 3: Parse XML
    ET.register_namespace('', SSML_NS)  # Preserve namespace
    root = ET.fromstring(workbook_xml)

    # Step 4: Find or create definedNames section
    defined_names_el = root.find(f"{{{SSML_NS}}}definedNames")
    if defined_names_el is None:
        defined_names_el = ET.SubElement(root, f"{{{SSML_NS}}}definedNames")

    # Step 5: Check for duplicate name
    existing = defined_names_el.find(f".//{{{SSML_NS}}}definedName[@name='{range_name}']")
    if existing is not None:
        print(f"[WARN] Named range '{range_name}' already exists, skipping")
        return False

    # Step 6: Add new named range
    new_dn = ET.SubElement(defined_names_el, f"{{{SSML_NS}}}definedName")
    new_dn.set("name", range_name)
    new_dn.text = range_ref

    # Step 7: Serialize back to XML
    new_xml = ET.tostring(root, encoding="unicode", xml_declaration=True)

    # Step 8: Repack ZIP (copy everything else untouched)
    tmp = Path(tempfile.mktemp(suffix=".xlsx"))
    with ZipFile(output_xlsx, "r") as zin, ZipFile(tmp, "w", ZIP_DEFLATED) as zout:
        for item in zin.namelist():
            if item == "xl/workbook.xml":
                zout.writestr(item, new_xml)
            else:
                zout.writestr(item, zin.read(item))

    tmp.replace(output_xlsx)
    print(f"[+] Added named range: {range_name} = {range_ref}")
    return True

# Usage example:
add_named_range(
    "MyWorkbook_DECRYPTED.xlsx",
    "MyWorkbook_MODIFIED.xlsx",
    "SalesData",
    "'Data'!$A$2:$A$100"
)
```

### Why This Pattern Is Safe

| Action | What Happens | Risk Level |
|--------|--------------|------------|
| Copy original file | Creates a working copy | ✅ None - original preserved |
| Read as ZIP | Opens archive, reads specific XML | ✅ None - read-only |
| Parse XML | Loads into memory | ✅ None - doesn't affect formulas |
| Add XML element | Appends to `<definedNames>` | ✅ None - surgical change |
| Repack ZIP | Copies all files except modified XML | ✅ None - byte-for-byte copy |

### What NOT to Do

❌ **Never use openpyxl to save:**
```python
# DANGER - This destroys formulas!
wb = openpyxl.load_workbook("file.xlsx")
wb.save("file.xlsx")  # ← All formulas replaced with values!
```

❌ **Never use pandas to write:**
```python
# DANGER - This only writes data, loses formulas!
df.to_excel("file.xlsx")  # ← Formulas gone!
```

✅ **Always use OpenXML (ZIP manipulation):**
```python
# SAFE - Only modifies specific XML element
with ZipFile(...) as z:
    modify_specific_xml(...)
```

---

## Security Best Practices

### 1. Never Commit Sensitive Files

**Add to `.gitignore`:**
```gitignore
# Password and credentials
.env

# Decrypted Excel files
*_DECRYPTED.xlsx
*_REZIPPED.xlsx

# Working directories
workbook_xml/
```

### 2. Use Environment Variables

**✅ DO:**
```python
from dotenv import load_dotenv
import os

load_dotenv()
password = os.getenv('EXCEL_PASSWORD')
```

**❌ DON'T:**
```python
# Hardcoded password visible to everyone!
password = "P4ssW0rX"
```

### 3. Share Only Safe Files

**Safe to share:**
- ✅ `explore_excel_xml.py` (script)
- ✅ `README.md` (this document)
- ✅ `.env.example` (template)

**NEVER share:**
- ❌ `.env` (contains actual password)
- ❌ `*_DECRYPTED.xlsx` (unencrypted workbook)
- ❌ Screenshots showing password values

### 4. Re-encrypt Before Delivery

If you modify the Excel file and need to deliver it:

```bash
# Re-encrypt using msoffcrypto-tool
msoffcrypto-tool \
  -p "YourPassword" \
  MyWorkbook_MODIFIED.xlsx \
  MyWorkbook_ENCRYPTED.xlsx

# Delete the decrypted version
rm MyWorkbook_MODIFIED.xlsx
```

### 5. Emergency: Password Leaked

If you accidentally commit the `.env` file:

```bash
# 1. Remove from git history
git filter-branch --force --index-filter \
  "git rm --cached --ignore-unmatch .env" \
  --prune-empty --tag-name-filter cat -- --all

# 2. Force push (coordinate with team!)
git push origin --force --all

# 3. Change the Excel password immediately
# 4. Update .env with new password
```

---

## Troubleshooting

### Error: "File is not a zip file"

**Symptom:**
```
zipfile.BadZipFile: File is not a zip file
```

**Cause:** The Excel file is encrypted

**Solution:**
- Ensure `EXCEL_PASSWORD` is set in `.env`
- Verify the password is correct
- Check that `python-dotenv` is installed

### Error: "EXCEL_PASSWORD environment variable not set"

**Symptom:**
```
[ERROR] No password provided and EXCEL_PASSWORD environment variable not set
```

**Cause:** `.env` file is missing or not loaded

**Solution:**
```bash
# Create .env file
echo "EXCEL_PASSWORD=your_password_here" > .env

# Verify it exists
cat .env
```

### Error: Decryption failed

**Symptom:**
```
[ERROR] Decryption failed: ...
```

**Cause:** Wrong password

**Solution:**
- Double-check the password in `.env`
- Ensure there are no extra spaces or quotes
- Try opening the Excel file manually to verify the password

### Error: Size difference > 20%

**Symptom:**
```
[WARN] Size difference: 25.00% (review recommended)
```

**Cause:** Something may have been lost during rezip

**Solution:**
- Check `Original files` vs `Rezipped files` count
- If counts match, size difference is likely just compression
- Open the rezipped file in Excel to verify it works
- If file is corrupted, review your XML modifications

### Script Can't Find the Excel File

**Symptom:**
```
[ERROR] File not found: MyFile.xlsx
```

**Solution:**
- Ensure the file is in the same directory as the script
- Or edit the script to use absolute path:
  ```python
  xlsx_file = Path(r"C:\full\path\to\file.xlsx")
  ```

### AI Assistant Asks for Password

**Symptom:** AI keeps asking "What's the password?"

**Cause:** `.env` not set up properly

**Solution:**
- Verify `.env` exists in the working directory
- Ensure `python-dotenv` is installed
- Check that script has `load_dotenv()` call
- Confirm password is on a line like: `EXCEL_PASSWORD=value`

---

## File Manifest

This toolkit consists of exactly **2 files**:

1. **`explore_excel_xml.py`** - The OpenXML exploration script
2. **`README.md`** - This documentation (you are here)

### Additional Files You'll Create

When you use this toolkit, you'll create:

- `.env` - Your local password file (never commit)
- `workbook_xml/` - Unzipped XML directory (working directory)
- `*_DECRYPTED.xlsx` - Temporary decrypted files (never commit)
- `*_REZIPPED.xlsx` - Test output files (never commit)

---

## Quick Start Checklist

- [ ] Install Python 3.7+
- [ ] Create virtual environment: `python3 -m venv venv`
- [ ] Activate venv: `source venv/Scripts/activate` (Windows) or `source venv/bin/activate` (Mac/Linux)
- [ ] Install dependencies: `pip install msoffcrypto-tool python-dotenv`
- [ ] Create `.env` file with `EXCEL_PASSWORD=your_password`
- [ ] Add `.env` to `.gitignore`
- [ ] Place your Excel file in the same directory
- [ ] Run: `python explore_excel_xml.py`
- [ ] Review output and explore `workbook_xml/` directory

---

## License & Attribution

This toolkit is designed for working with Draftworx UK Financials templates and similar complex Excel workbooks that cannot be safely modified using traditional Python Excel libraries.

**Key Principle:** Preserve everything, change only what is explicitly required.

---

## Questions for AI Assistants

If you're an AI assistant using this toolkit, here's what you need to know:

### Can I see the password?
**No.** The password is in the `.env` file, which you should not read directly. You can run scripts that use `os.getenv('EXCEL_PASSWORD')`, and the script will handle it.

### Can I modify the Excel file directly?
**No.** Never use `openpyxl.save()` or similar methods. Always use OpenXML (ZIP manipulation) as shown in the examples above.

### How do I add a named range?
Follow the **OpenXML Surgery Workflow** example above. Always:
1. Copy the original file
2. Modify specific XML
3. Repack the ZIP
4. Verify the output

### What if I need to read cell values?
Use `openpyxl` with `data_only=True` in **read-only mode**:
```python
from openpyxl import load_workbook
wb = load_workbook("file.xlsx", data_only=True)
# Read values
# NEVER call wb.save()
```

### The file is encrypted. What do I do?
Run `explore_excel_xml.py` - it automatically handles decryption using the `EXCEL_PASSWORD` from the environment.

---

**End of Documentation**
