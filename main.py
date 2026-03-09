import argparse
from win32com.client import Dispatch
from pathlib import Path
from glob import glob

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
    
    print(f"Input file: {input_file}")
    print(f"Output file: {output_file}")
    
    if not Path(input_file).exists():
        return
    
    if Path(output_file).exists():
        return
    
    excel = Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    workbook = excel.Workbooks.Open(input_file)
    workbook.SaveAs(output_file, FileFormat=57)  # 57 = PDF format
    workbook.Close()
    
    excel.Quit()
    
def convert_directory(directory: str):
    """
    Convert all Excel files in a directory to PDF
    """
    excel_files = glob(f"{directory}/*.xlsx")
    for excel_file in excel_files:
        output_file = Path.cwd() / pdf_directory / Path(excel_file).name.replace(".xlsx", ".pdf")
        convert(excel_file, str(output_file))

if __name__ == "__main__":
   
    args = main()
   
    if args.input:
        input_path = Path.joinpath(Path.cwd(), Path(args.input))
        if not input_path.exists():
            print(f"Error: Input file '{args.input}' not found")
            exit(1)
        output_path = pdf_directory / input_path.name.replace(".xlsx", ".pdf")
        convert(str(input_path), str(output_path))
    elif args.directory:
        input_path = Path.joinpath(Path.cwd(), Path(args.directory))
        if not input_path.exists():
            print(f"Error: Directory '{args.directory}' not found")
            exit(1)
        convert_directory(str(input_path))
   
   
   
   
