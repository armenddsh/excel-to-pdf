import argparse
from win32com.client import Dispatch
from pathlib import Path
from glob import glob

def main():
    """
    Convert Excel files to PDF
    
    Usage:
        uv run main.py -i input.xlsx
        uv run main.py -d "directory_path"
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
    
    if not Path(input_file).exists():
        print(f"Input file '{input_file}' does not exist")
        return
    
    if Path(output_file).exists():
        print(f"Output file '{output_file}' already exists")
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
        convert(excel_file, excel_file.replace(".xlsx", ".pdf"))

if __name__ == "__main__":
   
    args = main()
   
    if args.input:
        input_path = Path.joinpath(Path.cwd(), Path(args.input))
        if not input_path.exists():
            print(f"Error: Input file '{args.input}' not found")
            exit(1)
        print(f"Input path: {input_path}")
        output_path = input_path.with_suffix('.pdf')
        print(f"Output path: {output_path}")
        convert(str(input_path), str(output_path))
        print(f"Successfully converted {input_path} to {output_path}")
    elif args.directory:
        input_path = Path.joinpath(Path.cwd(), Path(args.directory))
        if not input_path.exists():
            print(f"Error: Directory '{args.directory}' not found")
            exit(1)
        convert_directory(str(input_path))
   
   
   
   
