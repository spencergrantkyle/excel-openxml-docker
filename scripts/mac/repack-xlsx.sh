#!/bin/bash
# ============================================
# Repack XML directory back to Excel (.xlsx)
# Usage: ./repack-xlsx.sh input_dir output.xlsx
# ============================================

set -e

# Colors for output
GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Check arguments
if [ -z "$1" ] || [ -z "$2" ]; then
    echo -e "${RED}Error: Missing arguments${NC}"
    echo "Usage: ./repack-xlsx.sh input_dir output.xlsx"
    echo ""
    echo "Examples:"
    echo "  ./repack-xlsx.sh ./workbook output.xlsx"
    echo "  ./repack-xlsx.sh ./myfile_xml myfile_modified.xlsx"
    exit 1
fi

INPUT_DIR="$1"
OUTPUT_FILE="$2"

# Validate input directory exists
if [ ! -d "$INPUT_DIR" ]; then
    echo -e "${RED}Error: Directory not found: $INPUT_DIR${NC}"
    exit 1
fi

# Validate it contains Excel XML structure
if [ ! -f "$INPUT_DIR/[Content_Types].xml" ]; then
    echo -e "${RED}Error: Not a valid Excel XML directory (missing [Content_Types].xml)${NC}"
    exit 1
fi

# Check if output file exists
if [ -f "$OUTPUT_FILE" ]; then
    echo -e "${YELLOW}Warning: Output file already exists: $OUTPUT_FILE${NC}"
    read -p "Overwrite? (y/n) " -n 1 -r
    echo
    if [[ ! $REPLY =~ ^[Yy]$ ]]; then
        echo "Aborted."
        exit 1
    fi
    rm -f "$OUTPUT_FILE"
fi

# Get absolute path for output
OUTPUT_FILE_ABS="$(cd "$(dirname "$OUTPUT_FILE")" && pwd)/$(basename "$OUTPUT_FILE")"

# Create xlsx (zip with specific structure)
echo -e "${GREEN}Repacking:${NC} $INPUT_DIR"
echo -e "${GREEN}To:${NC} $OUTPUT_FILE"

cd "$INPUT_DIR"

# Create zip with proper structure
# -r recursive, -q quiet, -X no extra file attributes
zip -r -q -X "$OUTPUT_FILE_ABS" . -x "*.DS_Store" -x "__MACOSX/*"

cd - > /dev/null

# Verify
if [ -f "$OUTPUT_FILE" ]; then
    FILE_SIZE=$(ls -lh "$OUTPUT_FILE" | awk '{print $5}')
    echo ""
    echo -e "${GREEN}✓ Repack complete${NC}"
    echo "  Output: $OUTPUT_FILE"
    echo "  Size: $FILE_SIZE"
else
    echo -e "${RED}Error: Failed to create output file${NC}"
    exit 1
fi

echo ""
echo -e "${YELLOW}Tip:${NC} Open in Excel to verify the file works correctly."
