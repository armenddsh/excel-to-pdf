@echo off
setlocal enabledelayedexpansion

REM Excel to PDF Converter - Windows Batch Script
REM Double-click to run with file selection

echo Excel to PDF Converter
echo.

REM Check if arguments were passed (drag-and-drop or command line)
if "%~1"=="" (
    echo No file specified. Please:
    echo 1. Drag and drop an Excel file onto this script
    echo 2. Drag and drop a folder onto this script
    echo 3. Type: convert.bat filename.xlsx
    echo.
    pause
    exit /b
)

REM Process each dropped file/folder
:process_files
if "%~1"=="" goto :done

set "TARGET=%~1"
echo Processing: !TARGET!

REM Check if it's a directory
if exist "!TARGET!\*" (
    echo Converting all Excel files in directory...
    uv run main.py -d "!TARGET!"
) else (
    REM Check if it's an Excel file
    echo !TARGET! | findstr /i "\.xlsx$" >nul
    if !errorlevel! equ 0 (
        echo Converting single Excel file...
        uv run main.py -i "!TARGET!"
    ) else (
        echo Error: "!TARGET!" is not an Excel file (.xlsx)
    )
)

echo.
shift
goto :process_files

:done
echo Conversion completed!
echo.
pause
