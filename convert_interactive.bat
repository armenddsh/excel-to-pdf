@echo off
setlocal enabledelayedexpansion

REM Excel to PDF Converter - Interactive Windows Batch Script
REM Double-click to run with file/folder selection dialog

echo Excel to PDF Converter
echo.

REM Check if arguments were passed (drag-and-drop)
if "%~1"=="" (
    echo No file specified. Opening file/folder selection...
    echo.
    
    REM Create a temporary VBScript to show file/folder dialog
    echo Set shell = CreateObject("Shell.Application") > "%temp%\selectfile.vbs"
    echo Set folder = shell.BrowseForFolder(0, "Select Excel file or folder containing Excel files", 0) >> "%temp%\selectfile.vbs"
    echo If folder Is Nothing Then >> "%temp%\selectfile.vbs"
    echo     WScript.Echo "CANCELLED" >> "%temp%\selectfile.vbs"
    echo Else >> "%temp%\selectfile.vbs"
    echo     WScript.Echo folder.Self.Path >> "%temp%\selectfile.vbs"
    echo End If >> "%temp%\selectfile.vbs"
    
    REM Run the VBScript and get the result
    for /f "delims=" %%F in ('cscript //nologo "%temp%\selectfile.vbs"') do set "SELECTED=%%F"
    
    REM Clean up the temporary file
    del "%temp%\selectfile.vbs" 2>nul
    
    REM Check if user cancelled
    if "!SELECTED!"=="CANCELLED" (
        echo Operation cancelled by user.
        pause
        exit /b
    )
    
    set "TARGET=!SELECTED!"
) else (
    set "TARGET=%~1"
)

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
        echo.
        echo Please select either:
        echo - An Excel file (.xlsx)
        echo - A folder containing Excel files
        pause
        exit /b
    )
)

echo.
echo Conversion completed!
echo.
pause
