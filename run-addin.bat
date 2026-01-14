@echo off
cd /d "%~dp0bin\Debug\net10.0-windows"
echo Launching Excel with EXLSXS Add-in...
echo.

if not exist EXLSXS-AddIn.xll (
    if exist ExcelDna.xll (
        echo Renaming ExcelDna.xll to EXLSXS-AddIn.xll...
        ren ExcelDna.xll EXLSXS-AddIn.xll
    ) else (
        echo EXLSXS-AddIn.xll (or ExcelDna.xll) not found.
        echo Please download ExcelDna.xll and place it here:
        echo %CD%
        pause
        exit /b 1
    )
)

echo Starting Excel with EXLSXS-AddIn.xll...
start EXLSXS-AddIn.xll
