#!/usr/bin/env python3
"""
Alternative Excel to PDF conversion using different methods
"""
import sys
import os
from pathlib import Path

# Add current directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    import win32com.client as win32
except ImportError:
    print("Error: pywin32 is not installed")
    input("Press Enter to exit...")
    sys.exit(1)

def test_excel_methods():
    """Test different Excel automation methods"""
    
    print("Testing Excel automation methods...")
    
    # Method 1: Standard Dispatch
    print("\n1. Testing standard Dispatch...")
    try:
        excel = win32.Dispatch("Excel.Application")
        print(f"   Excel version: {excel.Version}")
        excel.Quit()
        print("   ✓ Standard Dispatch works")
    except Exception as e:
        print(f"   ✗ Standard Dispatch failed: {e}")
    
    # Method 2: Dispatch with Create
    print("\n2. Testing DispatchEx...")
    try:
        excel = win32.DispatchEx("Excel.Application")
        print(f"   Excel version: {excel.Version}")
        excel.Quit()
        print("   ✓ DispatchEx works")
    except Exception as e:
        print(f"   ✗ DispatchEx failed: {e}")
    
    # Method 3: CreateObject
    print("\n3. Testing CreateObject...")
    try:
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        print(f"   Excel version: {excel.Version}")
        excel.Quit()
        print("   ✓ CreateObject works")
    except Exception as e:
        print(f"   ✗ CreateObject failed: {e}")

def create_test_file_simple():
    """Create test file using simple method"""
    print("\nCreating test Excel file...")
    
    try:
        # Try different creation methods
        methods = [
            lambda: win32.Dispatch("Excel.Application"),
            lambda: win32.DispatchEx("Excel.Application"),
            lambda: win32.gencache.EnsureDispatch("Excel.Application")
        ]
        
        excel = None
        for i, method in enumerate(methods):
            try:
                print(f"   Trying method {i+1}...")
                excel = method()
                break
            except Exception as e:
                print(f"   Method {i+1} failed: {e}")
                continue
        
        if not excel:
            print("   All Excel creation methods failed")
            return False
        
        excel.Visible = False
        excel.DisplayAlerts = False
        
        # Create workbook
        workbook = excel.Workbooks.Add()
        
        # Add some data
        sheet = workbook.Worksheets(1)
        sheet.Range("A1").Value = "Test"
        sheet.Range("B1").Value = "Data"
        sheet.Range("A2").Value = "Row 2"
        
        # Save with full path
        test_file = Path.cwd() / "test_simple.xlsx"
        workbook.SaveAs(str(test_file))
        workbook.Close()
        excel.Quit()
        
        print(f"   ✓ Created: {test_file}")
        return True
        
    except Exception as e:
        print(f"   ✗ Error: {e}")
        return False

if __name__ == "__main__":
    print("Excel Automation Diagnostic Tool")
    print("=" * 50)
    
    test_excel_methods()
    
    if create_test_file_simple():
        print("\n✓ Excel automation is working!")
        print("You can now try the main GUI application.")
    else:
        print("\n✗ Excel automation has issues.")
        print("\nSuggested solutions:")
        print("1. Restart your computer")
        print("2. Close all Excel applications")
        print("3. Run Excel as administrator once")
        print("4. Repair Microsoft Office installation")
        print("5. Check Windows Defender/antivirus blocking")
    
    input("\nPress Enter to exit...")
