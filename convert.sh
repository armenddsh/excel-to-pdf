#!/bin/bash
# Excel to PDF Converter - Bash Script
# Usage: ./convert.sh [file.xlsx] [directory]

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

show_help() {
    echo -e "${YELLOW}Excel to PDF Converter${NC}"
    echo
    echo "Usage:"
    echo "  ./convert.sh file.xlsx              - Convert single file"
    echo "  ./convert.sh directory              - Convert all Excel files in directory"
    echo "  ./convert.sh                         - Show this help"
    echo
    echo "Examples:"
    echo "  ./convert.sh report.xlsx"
    echo "  ./convert.sh excel_files"
    echo
    echo "Note: This requires Windows with Microsoft Excel installed"
}

# Check if uv is installed
if ! command -v uv &> /dev/null; then
    echo -e "${RED}Error: 'uv' is not installed or not in PATH${NC}"
    echo "Please install uv first: https://github.com/astral-sh/uv"
    exit 1
fi

# Check if running on Windows (required for Excel COM automation)
if [[ "$OSTYPE" != "msys" && "$OSTYPE" != "cygwin" && "$OSTYPE" != "win32" ]]; then
    echo -e "${RED}Error: This script requires Windows to use Excel COM automation${NC}"
    exit 1
fi

if [ $# -eq 0 ]; then
    show_help
    exit 0
fi

echo -e "${GREEN}Converting Excel files to PDF...${NC}"

# Pass all arguments to the Python script
uv run main.py "$@"

if [ $? -eq 0 ]; then
    echo -e "${GREEN}Conversion completed successfully!${NC}"
else
    echo -e "${RED}Conversion failed. Check the error messages above.${NC}"
    exit 1
fi
