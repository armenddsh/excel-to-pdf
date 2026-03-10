#!/usr/bin/env python3
"""
Test Excel to PDF conversion - simplified version
"""
import sys
import os
from pathlib import Path

# Add current directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from win32com.client import Dispatch
except ImportError:
    print("Error: pywin32 is not installed")
    input("Press Enter to exit...")
    sys.exit(1)

def test_conversion():
    """Test conversion with a simple file"""
    
    # Create a simple test Excel file if it doesn't exist
    test_file = Path("test.xlsx")
    if not test_file.exists():
        print("Creating test Excel file...")
        try:
            excel = Dispatch("Excel.Application")
            excel.Visible = False
            
            # Add delay and retry logic
            import time
            time.sleep(1)
            
            workbook = excel.Workbooks.Add()
            worksheet = workbook.Worksheets(1)
            worksheet.Cells(1, 1).Value = "Test"
            worksheet.Cells(1, 2).Value = "Data"
            
            # Try different save methods
            try:
                workbook.SaveAs(str(test_file))
            except Exception as save_error:
                print(f"SaveAs failed: {save_error}")
                # Try alternative save
                workbook.Save()
                # Copy to desired location
                import shutil
                temp_path = Path.cwd() / "Book1.xlsx"
                if temp_path.exists():
                    shutil.move(str(temp_path), str(test_file))
            
            workbook.Close(SaveChanges=False)
            excel.Quit()
            
            # Force cleanup
            import gc
            gc.collect()
            time.sleep(1)
            
            print(f"Created test file: {test_file}")
        except Exception as e:
            print(f"Error creating test file: {e}")
            print("This suggests Excel COM automation is not working properly.")
            print("Possible solutions:")
            print("1. Restart your computer")
            print("2. Repair Microsoft Office installation")
            print("3. Run Excel as administrator once")
            return False
    
    # Try to convert the test file
    output_file = Path("test.pdf")
    
    print(f"Testing conversion: {test_file} -> {output_file}")
    
    excel = None
    try:
        excel = Dispatch("Excel.Application")
        excel.Visible = True  # Make Excel visible to see what's happening
        excel.DisplayAlerts = False
        
        print("Opening workbook...")
        workbook = excel.Workbooks.Open(str(test_file))
        
        print("Exporting to PDF...")
        try:
            # Try different export methods
            try:
                workbook.ExportAsFixedFormat(0, str(output_file))
                print("ExportAsFixedFormat completed")
            except Exception as export_error:
                print(f"ExportAsFixedFormat failed: {export_error}")
                print("Trying SaveAs method...")
                workbook.SaveAs(str(output_file), FileFormat=57)  # PDF format
                print("SaveAs completed")
        except Exception as pdf_error:
            print(f"PDF export error: {pdf_error}")
            print("Trying alternative approach...")
            # Try saving as XPS first, then convert
            xps_file = Path("test.xps")
            try:
                workbook.ExportAsFixedFormat(1, str(xps_file))  # 1 = xlTypeXPS
                print(f"XPS export worked: {xps_file}")
            except Exception as xps_error:
                print(f"XPS export also failed: {xps_error}")
        
        workbook.Close(SaveChanges=False)
        
        if output_file.exists():
            print(f"SUCCESS: PDF created at {output_file}")
            return True
        else:
            print("FAILED: PDF was not created")
            return False
            
    except Exception as e:
        print(f"Error during conversion: {e}")
        return False
    finally:
        if excel:
            try:
                excel.Quit()
            except:
                pass

if __name__ == "__main__":
    print("Excel to PDF Conversion Test")
    print("=" * 40)
    
    # Check Excel version
    try:
        excel = Dispatch("Excel.Application")
        print(f"Excel version: {excel.Version}")
        excel.Quit()
    except Exception as e:
        print(f"Error accessing Excel: {e}")
        input("Press Enter to exit...")
        sys.exit(1)
    
    # Test conversion
    success = test_conversion()
    
    if success:
        print("\nTest PASSED! The conversion system is working.")
    else:
        print("\nTest FAILED! Check the error messages above.")
    
    input("\nPress Enter to exit...")
