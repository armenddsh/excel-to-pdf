import argparse
from pathlib import Path
from glob import glob

from excel_com import ExcelConversionError, convert_excel_to_pdf

pdf_directory = Path("pdfs")
if not pdf_directory.exists():
    pdf_directory.mkdir()

def main():
    """
    Convert Excel files to PDF
    
    Usage:
        uv run main.py -i input.xlsx
        uv run main.py -d "pdfs"
    """
    
    parser = argparse.ArgumentParser(description="Convert Excel files to PDF")
    parser.add_argument("-i", "--input", help="Input Excel file", type=str)
    parser.add_argument("-d", "--directory", help="Directory containing Excel files", type=str)
    
    args = parser.parse_args()
    
    print(f"Converting {args}")
    
    return args

def convert(input_file: str, output_file: str):
    """
    Convert Excel file to PDF
    """
    input_path = Path(input_file).resolve()
    output_path = Path(output_file).resolve()

    print(f"Input file: {input_path}")
    print(f"Output file: {output_path}")

    if not input_path.exists():
        print(f"Error: Input file '{input_path}' not found")
        return False

    if output_path.exists():
        print(f"Skipping existing PDF: {output_path}")
        return True

    try:
        convert_excel_to_pdf(input_path, output_path)
        return True
    except ExcelConversionError as exc:
        print(f"Error converting file: {exc}")
        return False
    
def convert_directory(directory: str):
    """
    Convert all Excel files in a directory to PDF
    """
    excel_files = glob(f"{directory}/*.xlsx") + glob(f"{directory}/*.xls")
    for excel_file in excel_files:
        output_file = Path.cwd() / pdf_directory / Path(excel_file).with_suffix(".pdf").name
        convert(excel_file, str(output_file))

if __name__ == "__main__":
   
    args = main()
   
    if args.input:
        input_path = Path.joinpath(Path.cwd(), Path(args.input))
        if not input_path.exists():
            print(f"Error: Input file '{args.input}' not found")
            exit(1)
        output_path = pdf_directory / input_path.with_suffix(".pdf").name
        if not convert(str(input_path), str(output_path)):
            exit(1)
    elif args.directory:
        input_path = Path.joinpath(Path.cwd(), Path(args.directory))
        if not input_path.exists():
            print(f"Error: Directory '{args.directory}' not found")
            exit(1)
        convert_directory(str(input_path))
   
   
   
   
