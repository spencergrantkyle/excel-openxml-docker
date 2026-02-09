#!/usr/bin/env python3
"""
Excel OpenXML Explorer
Safely unzips an Excel file to examine its XML structure
and demonstrates the round-trip (unzip → rezip) workflow.
"""

import os
import sys
import shutil
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED, BadZipFile
import xml.etree.ElementTree as ET
from datetime import datetime
import msoffcrypto
import io
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()


def decrypt_if_needed(xlsx_path, password=None):
    """Decrypt Excel file if it's encrypted. Returns path to usable file.

    Args:
        xlsx_path: Path to the Excel file
        password: Optional password override. If not provided, reads from EXCEL_PASSWORD env var
    """
    xlsx_path = Path(xlsx_path)

    # Get password from environment variable if not provided
    if password is None:
        password = os.getenv('EXCEL_PASSWORD')
        if not password:
            print("[ERROR] No password provided and EXCEL_PASSWORD environment variable not set")
            print("[ERROR] Please set EXCEL_PASSWORD in your .env file or environment")
            return None

    # Try to open as ZIP first - if it works, it's not encrypted
    try:
        with ZipFile(xlsx_path, 'r') as test_zip:
            test_zip.namelist()
        print(f"[*] File is not encrypted, using original")
        return xlsx_path
    except BadZipFile:
        # File is likely encrypted
        print(f"[*] File appears to be encrypted")
        print(f"[*] Attempting decryption using EXCEL_PASSWORD from environment")

        decrypted_path = xlsx_path.with_name(xlsx_path.stem + "_DECRYPTED.xlsx")

        try:
            with open(xlsx_path, 'rb') as encrypted_file:
                office_file = msoffcrypto.OfficeFile(encrypted_file)
                office_file.load_key(password=password)

                with open(decrypted_path, 'wb') as decrypted_file:
                    office_file.decrypt(decrypted_file)

            print(f"[+] Successfully decrypted to: {decrypted_path.name}")
            return decrypted_path

        except Exception as e:
            print(f"[ERROR] Decryption failed: {e}")
            return None


def unzip_excel(xlsx_path, output_dir):
    """Unzip Excel file to examine XML structure."""
    xlsx_path = Path(xlsx_path)
    output_dir = Path(output_dir)

    if not xlsx_path.exists():
        print(f"[ERROR] File not found: {xlsx_path}")
        return False

    print(f"[*] Unzipping: {xlsx_path.name}")
    print(f"[*] Output directory: {output_dir}")

    # Remove existing output directory
    if output_dir.exists():
        shutil.rmtree(output_dir)

    # Extract all files
    try:
        with ZipFile(xlsx_path, 'r') as zip_ref:
            zip_ref.extractall(output_dir)
    except BadZipFile as e:
        print(f"[ERROR] Cannot unzip file: {e}")
        return False

    print(f"[+] Extracted {len(list(output_dir.rglob('*')))} files and directories")
    return True


def show_structure(root_dir, max_depth=3):
    """Display directory structure."""
    root_dir = Path(root_dir)

    print(f"\n[DIR] Directory Structure (max depth {max_depth}):\n")

    def print_tree(directory, prefix="", depth=0):
        if depth > max_depth:
            return

        try:
            contents = sorted(directory.iterdir(), key=lambda x: (not x.is_dir(), x.name))
        except PermissionError:
            return

        for i, path in enumerate(contents):
            is_last = i == len(contents) - 1
            current_prefix = "+-- " if is_last else "+-- "
            print(f"{prefix}{current_prefix}{path.name}")

            if path.is_dir() and depth < max_depth:
                extension_prefix = "    " if is_last else "|   "
                print_tree(path, prefix + extension_prefix, depth + 1)

    print_tree(root_dir)


def examine_key_files(xml_dir):
    """Examine key XML files."""
    xml_dir = Path(xml_dir)

    print("\n" + "="*70)
    print("[INSPECT] KEY XML FILES")
    print("="*70)

    # 1. workbook.xml - contains sheet list and defined names
    workbook_xml = xml_dir / "xl" / "workbook.xml"
    if workbook_xml.exists():
        print(f"\n[1] xl/workbook.xml")
        print("   Contains: worksheet list, named ranges (definedNames)")
        tree = ET.parse(workbook_xml)
        root = tree.getroot()

        # Count sheets
        sheets = root.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet")
        print(f"   Worksheets: {len(sheets)}")

        # Count named ranges
        defined_names = root.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}definedName")
        print(f"   Named Ranges: {len(defined_names)}")

        # Show first 3 sheets
        if sheets:
            print(f"   First 3 sheets:")
            for sheet in sheets[:3]:
                name = sheet.get('name')
                sheet_id = sheet.get('sheetId')
                rid = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                print(f"      - {name} (sheetId={sheet_id}, r:id={rid})")

        # Show first 3 named ranges
        if defined_names:
            print(f"   First 3 named ranges:")
            for dn in defined_names[:3]:
                name = dn.get('name')
                value = dn.text[:50] if dn.text else ""
                print(f"      - {name} = {value}...")

    # 2. Count worksheet files
    worksheets_dir = xml_dir / "xl" / "worksheets"
    if worksheets_dir.exists():
        worksheet_files = list(worksheets_dir.glob("sheet*.xml"))
        print(f"\n[2] xl/worksheets/")
        print(f"   Contains: {len(worksheet_files)} worksheet XML files")
        print(f"   Example: {worksheet_files[0].name if worksheet_files else 'None'}")

    # 3. Shared strings
    shared_strings = xml_dir / "xl" / "sharedStrings.xml"
    if shared_strings.exists():
        tree = ET.parse(shared_strings)
        root = tree.getroot()
        si_elements = root.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si")
        print(f"\n[3] xl/sharedStrings.xml")
        print(f"   Contains: {len(si_elements)} shared string entries")

    # 4. Styles
    styles = xml_dir / "xl" / "styles.xml"
    if styles.exists():
        print(f"\n[4] xl/styles.xml")
        print(f"   Contains: cell formats, fonts, fills, borders")

    # 5. Content Types
    content_types = xml_dir / "[Content_Types].xml"
    if content_types.exists():
        print(f"\n[5] [Content_Types].xml")
        print(f"   Contains: MIME types for all XML parts")


def rezip_excel(xml_dir, output_xlsx):
    """Rezip the XML directory back into an Excel file."""
    xml_dir = Path(xml_dir)
    output_xlsx = Path(output_xlsx)

    print(f"\n[*] Rezipping XML directory...")
    print(f"[*] Source: {xml_dir}")
    print(f"[*] Output: {output_xlsx.name}")

    with ZipFile(output_xlsx, 'w', ZIP_DEFLATED) as zip_out:
        for file_path in xml_dir.rglob('*'):
            if file_path.is_file():
                arcname = file_path.relative_to(xml_dir)
                zip_out.write(file_path, arcname)

    file_size = output_xlsx.stat().st_size
    print(f"[+] Created: {output_xlsx.name} ({file_size:,} bytes)")
    return True


def verify_round_trip(original_xlsx, rezipped_xlsx):
    """Verify the round-trip worked by comparing file sizes."""
    original_size = Path(original_xlsx).stat().st_size
    rezipped_size = Path(rezipped_xlsx).stat().st_size

    print(f"\n[VERIFY] ROUND-TRIP VALIDATION")
    print(f"   Original: {original_size:,} bytes")
    print(f"   Rezipped: {rezipped_size:,} bytes")

    size_diff_pct = abs(rezipped_size - original_size) / original_size * 100

    if size_diff_pct < 5:
        print(f"   [OK] Size difference: {size_diff_pct:.2f}% (acceptable)")
    else:
        print(f"   [WARN] Size difference: {size_diff_pct:.2f}% (review recommended)")

    # Try to open both with ZipFile
    try:
        with ZipFile(original_xlsx, 'r') as z:
            original_files = set(z.namelist())
        with ZipFile(rezipped_xlsx, 'r') as z:
            rezipped_files = set(z.namelist())

        print(f"   Original files: {len(original_files)}")
        print(f"   Rezipped files: {len(rezipped_files)}")

        if original_files == rezipped_files:
            print(f"   [OK] File list matches perfectly")
        else:
            missing = original_files - rezipped_files
            extra = rezipped_files - original_files
            if missing:
                print(f"   [WARN] Missing files: {missing}")
            if extra:
                print(f"   [WARN] Extra files: {extra}")

    except Exception as e:
        print(f"   [ERROR] Verification failed: {e}")


def main():
    # Configuration
    xlsx_file = Path("Financials+_RepeatHeaderNamedRanges SPG.xlsx")
    xml_output_dir = Path("workbook_xml")
    rezipped_file = Path("Financials+_RepeatHeaderNamedRanges_REZIPPED.xlsx")

    print("="*70)
    print("EXCEL OPENXML EXPLORER")
    print("="*70)
    print(f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

    # Step 0: Decrypt if needed
    working_file = decrypt_if_needed(xlsx_file)
    if working_file is None:
        print("[ERROR] Could not prepare file for processing")
        return 1

    # Step 1: Unzip
    if not unzip_excel(working_file, xml_output_dir):
        return 1

    # Step 2: Show structure
    show_structure(xml_output_dir, max_depth=2)

    # Step 3: Examine key files
    examine_key_files(xml_output_dir)

    # Step 4: Rezip
    print("\n" + "="*70)
    print("[TEST] ROUND-TRIP (UNZIP -> REZIP)")
    print("="*70)
    rezip_excel(xml_output_dir, rezipped_file)

    # Step 5: Verify
    verify_round_trip(working_file, rezipped_file)

    print("\n" + "="*70)
    print("[SUCCESS] COMPLETE")
    print("="*70)
    print(f"\nXML directory: {xml_output_dir.absolute()}")
    print(f"Rezipped file: {rezipped_file.absolute()}")
    print(f"\nNext steps:")
    print(f"   1. Explore {xml_output_dir}/xl/workbook.xml to see named ranges")
    print(f"   2. Explore {xml_output_dir}/xl/worksheets/ to see sheet data")
    print(f"   3. Use OpenXML techniques to surgically edit specific XML files")
    print(f"   4. Test the rezipped file in Excel to verify it opens correctly")

    return 0


if __name__ == "__main__":
    sys.exit(main())
