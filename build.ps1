# Build script for ACQUA Report Reviewer
# Run this script to create the executable

# Get version from the main script
$version = python -c "import sys; sys.path.insert(0, '.'); from process_acqua_reports import __version__; print(__version__)"

Write-Host "Building ACQUA Report Reviewer v$version..." -ForegroundColor Cyan

# Install dependencies if needed
Write-Host "Installing dependencies..." -ForegroundColor Yellow
pip install -r requirements.txt

# Find python-docx templates location
Write-Host "Locating python-docx templates..." -ForegroundColor Yellow
$docxPath = python -c "import docx; import os; print(os.path.dirname(docx.__file__))"
Write-Host "Found docx at: $docxPath" -ForegroundColor Gray

# Build the executable with python-docx templates included
Write-Host "Creating executable..." -ForegroundColor Yellow
pyinstaller --onefile --name "ACQUA_ReportReviewer_v$version" `
    --add-data "$docxPath/templates;docx/templates" `
    --hidden-import "docx" `
    --hidden-import "docx.document" `
    --hidden-import "docx.opc.constants" `
    --hidden-import "docx.opc.package" `
    --hidden-import "docx.opc.packuri" `
    --hidden-import "docx.opc.part" `
    --hidden-import "docx.opc.phys_pkg" `
    --hidden-import "docx.opc.rel" `
    --hidden-import "lxml._elementpath" `
    --hidden-import "lxml.etree" `
    --hidden-import "win32com" `
    --hidden-import "win32com.client" `
    --hidden-import "pythoncom" `
    --hidden-import "pywintypes" `
    process_acqua_reports.py

Write-Host ""
Write-Host "Build complete!" -ForegroundColor Green
Write-Host "Executable location: dist\ACQUA_ReportReviewer_v$version.exe" -ForegroundColor Green
