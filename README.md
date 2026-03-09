# Excel to PDF Converter

A Python utility to convert Excel files (.xlsx) to PDF format using Microsoft Excel COM automation on Windows.

## Features

- Convert single Excel files to PDF
- Batch convert all Excel files in a directory
- Simple command-line interface
- Uses Microsoft Excel's built-in PDF export functionality

## Requirements

- Windows operating system
- Microsoft Excel installed
- Python 3.13+
- pywin32 package

## Installation

1. Clone this repository:
```bash
git clone <repository-url>
cd excel-to-pdf
```

2. Install dependencies using uv:
```bash
uv sync
```

## Usage

### Convert a single Excel file

```bash
uv run main.py -i input.xlsx
```

This will create `input.pdf` in the same directory.

### Convert all Excel files in a directory

```bash
uv run main.py -d "path/to/excel/files"
```

This will convert all `.xlsx` files in the specified directory to PDF format.

## Command Line Options

- `-i, --input`: Path to a single Excel file to convert
- `-d, --directory`: Path to a directory containing Excel files to convert in batch

## How It Works

This utility uses the `pywin32` library to interface with Microsoft Excel's COM automation:

1. Creates an instance of Excel application
2. Opens the specified Excel workbook
3. Saves the workbook as PDF using Excel's built-in export functionality
4. Closes the workbook and quits Excel

The conversion preserves the formatting, layout, and content of the original Excel file.

## Notes

- Microsoft Excel must be installed on the system
- The Excel application runs invisibly in the background during conversion
- Output PDF files will be created in the same directory as the input files
- Existing PDF files will not be overwritten

## Example

```bash
# Convert a single file
uv run main.py -i financial_report.xlsx

# Convert all files in a directory
uv run main.py -d "C:\Documents\Spreadsheets"
```

## License

This project is open source. Please refer to the license file for details.