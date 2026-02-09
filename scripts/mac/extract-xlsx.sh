#!/bin/bash
# ============================================
# Extract Excel (.xlsx) to XML for editing
# Usage: ./extract-xlsx.sh input.xlsx [output_dir]
# ============================================

set -e

# Colors for output
GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Check arguments
if [ -z "$1" ]; then
    echo -e "${RED}Error: No input file specified${NC}"
    echo "Usage: ./extract-xlsx.sh input.xlsx [output_dir]"
    echo ""
    echo "Examples:"
    echo "  ./extract-xlsx.sh myfile.xlsx"
    echo "  ./extract-xlsx.sh myfile.xlsx ./workbook"
    exit 1
fi

INPUT_FILE="$1"
OUTPUT_DIR="${2:-${INPUT_FILE%.xlsx}_xml}"

# Validate input file exists
if [ ! -f "$INPUT_FILE" ]; then
    echo -e "${RED}Error: File not found: $INPUT_FILE${NC}"
    exit 1
fi

# Validate it's an xlsx file
if [[ ! "$INPUT_FILE" == *.xlsx ]]; then
    echo -e "${YELLOW}Warning: File doesn't have .xlsx extension${NC}"
fi

# Check if output directory exists
if [ -d "$OUTPUT_DIR" ]; then
    echo -e "${YELLOW}Warning: Output directory already exists: $OUTPUT_DIR${NC}"
    read -p "Overwrite? (y/n) " -n 1 -r
    echo
    if [[ ! $REPLY =~ ^[Yy]$ ]]; then
        echo "Aborted."
        exit 1
    fi
    rm -rf "$OUTPUT_DIR"
fi

# Create output directory
mkdir -p "$OUTPUT_DIR"

# Extract
echo -e "${GREEN}Extracting:${NC} $INPUT_FILE"
echo -e "${GREEN}To:${NC} $OUTPUT_DIR"

unzip -q "$INPUT_FILE" -d "$OUTPUT_DIR"

# Show structure
echo ""
echo -e "${GREEN}✓ Extraction complete${NC}"
echo ""
echo "Structure:"
find "$OUTPUT_DIR" -type f | head -20

# Count files
FILE_COUNT=$(find "$OUTPUT_DIR" -type f | wc -l | tr -d ' ')
echo ""
echo -e "${GREEN}Total files: $FILE_COUNT${NC}"

# Reminder
echo ""
echo -e "${YELLOW}Key files for editing:${NC}"
echo "  $OUTPUT_DIR/xl/workbook.xml       - Sheet names, named ranges"
echo "  $OUTPUT_DIR/xl/worksheets/*.xml   - Cell data and formulas"
echo "  $OUTPUT_DIR/xl/sharedStrings.xml  - Text content"
echo ""
echo -e "${YELLOW}To repack:${NC} ./repack-xlsx.sh $OUTPUT_DIR output.xlsx"
