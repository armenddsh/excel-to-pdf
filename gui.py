#!/usr/bin/env python3
"""
Excel to PDF Converter GUI
Double-click this file to run the application
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from glob import glob
import threading
import sys
import os

# Use absolute paths for reliability
CURRENT_DIR = Path(__file__).parent.resolve()
sys.path.insert(0, str(CURRENT_DIR))

try:
    from excel_com import ExcelConversionError, convert_excel_to_pdf
except ImportError:
    print("Error: required dependencies are not installed")
    print("Please run: uv sync")
    input("Press Enter to exit...")
    sys.exit(1)

def convert(input_file: str, output_file: str):
    """
    Convert Excel file to PDF
    """
    # Use absolute paths
    input_path = Path(input_file).resolve()
    output_path = Path(output_file).resolve()
    
    print(f"Input file: {input_path}")
    print(f"Output file: {output_path}")
    
    if not input_path.exists():
        print(f"Error: Input file '{input_path}' not found")
        return False
    
    # Check if output file already exists and remove it
    if output_path.exists():
        try:
            output_path.unlink()
        except Exception as e:
            print(f"Warning: Could not remove existing file '{output_path}': {e}")
            return False
    
    try:
        convert_excel_to_pdf(input_path, output_path)
        return True
    except ExcelConversionError as e:
        error_msg = str(e)
        print(f"Error converting file: {error_msg}")

        # Check for specific Excel errors
        if "Document not saved" in error_msg:
            print("This usually means:")
            print("1. The Excel file is already open")
            print("2. You don't have permission to save to the output location")
            print("3. The Excel file is corrupted or protected")

        return False

def convert_directory(directory: str):
    """
    Convert all Excel files in a directory to PDF
    """
    pdf_directory = CURRENT_DIR / "pdfs"
    if not pdf_directory.exists():
        pdf_directory.mkdir()
    
    dir_path = Path(directory).resolve()
    excel_files = list(dir_path.glob("*.xlsx")) + list(dir_path.glob("*.xls"))
    converted_count = 0
    
    for excel_file in excel_files:
        output_file = pdf_directory / excel_file.name.replace(excel_file.suffix, ".pdf")
        if convert(str(excel_file), str(output_file)):
            converted_count += 1
    
    return converted_count

class ExcelToPDFGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to PDF Converter")
        self.root.geometry("600x400")
        self.root.configure(bg='#f0f0f0')
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = ttk.Label(main_frame, text="Excel to PDF Converter", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # File selection area
        selection_frame = ttk.LabelFrame(main_frame, text="Select Files to Convert", padding="20")
        selection_frame.grid(row=1, column=0, columnspan=2, pady=20, padx=10, 
                            sticky=(tk.W, tk.E))
        
        # Browse buttons
        self.file_button = ttk.Button(selection_frame, text="Browse Excel Files", 
                                      command=self.browse_files, width=20)
        self.file_button.grid(row=0, column=0, padx=10, pady=10)
        
        self.folder_button = ttk.Button(selection_frame, text="Browse Folder", 
                                       command=self.browse_folder, width=20)
        self.folder_button.grid(row=0, column=1, padx=10, pady=10)
        
        # Selected files display
        self.files_var = tk.StringVar(value="No files selected")
        self.files_label = ttk.Label(selection_frame, textvariable=self.files_var, 
                                     wraplength=500, justify='left')
        self.files_label.grid(row=1, column=0, columnspan=2, pady=10, padx=10)
        
        # Convert button
        self.convert_button = ttk.Button(main_frame, text="Convert to PDF", 
                                       command=self.convert_selected, state='disabled')
        self.convert_button.grid(row=2, column=0, columnspan=2, pady=20)
        
        # Status area
        self.status_var = tk.StringVar(value="Ready - Select files or folder to begin")
        self.status_label = ttk.Label(main_frame, textvariable=self.status_var, 
                                      font=('Arial', 10))
        self.status_label.grid(row=3, column=0, columnspan=2, pady=10)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=4, column=0, columnspan=2, pady=10, 
                          sticky=(tk.W, tk.E), padx=10)
        
        # Output folder info
        output_info = ttk.Label(main_frame, 
                               text=f"PDFs will be saved to: {CURRENT_DIR / 'pdfs'}",
                               font=('Arial', 9), foreground='#666')
        output_info.grid(row=5, column=0, columnspan=2, pady=5)
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        
        self.selected_files = []
        self.selected_folder = None
        
    def browse_files(self):
        files = filedialog.askopenfilenames(
            title="Select Excel files",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if files:
            self.selected_files = list(files)
            self.selected_folder = None
            self.update_selection_display()
            
    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select folder with Excel files")
        if folder:
            self.selected_folder = folder
            self.selected_files = []
            self.update_selection_display()
            
    def update_selection_display(self):
        if self.selected_files:
            count = len(self.selected_files)
            if count <= 3:
                display_text = f"Selected {count} file(s):\n" + "\n".join(
                    [Path(f).name for f in self.selected_files])
            else:
                display_text = f"Selected {count} files:\n" + "\n".join(
                    [Path(f).name for f in self.selected_files[:3]]) + f"\n... and {count-3} more"
            self.files_var.set(display_text)
            self.convert_button.config(state='normal')
        elif self.selected_folder:
            excel_files = glob(f"{self.selected_folder}/*.xlsx") + glob(f"{self.selected_folder}/*.xls")
            count = len(excel_files)
            folder_name = Path(self.selected_folder).name
            display_text = f"Selected folder: {folder_name}\nFound {count} Excel file(s)"
            self.files_var.set(display_text)
            self.convert_button.config(state='normal')
        else:
            self.files_var.set("No files selected")
            self.convert_button.config(state='disabled')
            
    def convert_selected(self):
        if self.selected_files:
            self.process_files(self.selected_files)
        elif self.selected_folder:
            self.process_folder(self.selected_folder)
            
    def process_files(self, files):
        if not files:
            return
            
        def convert_thread():
            try:
                self.status_var.set("Converting files...")
                self.progress.start()
                self.convert_button.config(state='disabled')
                
                pdf_directory = CURRENT_DIR / "pdfs"
                if not pdf_directory.exists():
                    pdf_directory.mkdir()
                
                converted_count = 0
                for file_path in files:
                    if file_path.lower().endswith(('.xlsx', '.xls')):
                        input_path = Path(file_path).resolve()
                        output_file = pdf_directory / input_path.name.replace(
                            input_path.suffix, '.pdf')
                        
                        self.status_var.set(f"Converting {input_path.name}...")
                        if convert(str(input_path), str(output_file)):
                            converted_count += 1
                
                self.status_var.set(f"Conversion complete! {converted_count} file(s) processed")
                messagebox.showinfo("Success", f"Successfully converted {converted_count} file(s) to PDF")
                
            except Exception as e:
                self.status_var.set(f"Error: {str(e)}")
                messagebox.showerror("Error", f"Conversion failed: {str(e)}")
            finally:
                self.progress.stop()
                self.convert_button.config(state='normal')
        
        threading.Thread(target=convert_thread, daemon=True).start()
        
    def process_folder(self, folder):
        def convert_thread():
            try:
                self.status_var.set("Scanning folder for Excel files...")
                self.progress.start()
                self.convert_button.config(state='disabled')
                
                excel_files = glob(f"{folder}/*.xlsx") + glob(f"{folder}/*.xls")
                
                if not excel_files:
                    self.status_var.set("No Excel files found in folder")
                    messagebox.showinfo("Info", "No Excel files found in the selected folder")
                    return
                
                self.status_var.set(f"Found {len(excel_files)} Excel files. Converting...")
                converted_count = convert_directory(folder)
                
                self.status_var.set(f"Conversion complete! {converted_count} file(s) processed")
                messagebox.showinfo("Success", f"Successfully converted {converted_count} file(s) to PDF")
                
            except Exception as e:
                self.status_var.set(f"Error: {str(e)}")
                messagebox.showerror("Error", f"Conversion failed: {str(e)}")
            finally:
                self.progress.stop()
                self.convert_button.config(state='normal')
        
        threading.Thread(target=convert_thread, daemon=True).start()

def main():
    root = tk.Tk()
    app = ExcelToPDFGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
