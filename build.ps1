# Build script for ACQUA Report Reviewer
# Run this script to create the executable

# Get version from the main script
$version = python -c "import sys; sys.path.insert(0, '.'); from process_acqua_reports import __version__; print(__version__)"

Write-Host "Building ACQUA Report Reviewer v$version..." -ForegroundColor Cyan

# Install dependencies if needed
Write-Host "Installing dependencies..." -ForegroundColor Yellow
pip install -r requirements.txt

# Build the executable
Write-Host "Creating executable..." -ForegroundColor Yellow
pyinstaller --onefile --name "ACQUA_ReportReviewer_v1.0.0" process_acqua_reports.py

Write-Host ""
Write-Host "Build complete!" -ForegroundColor Green
Write-Host "Executable location: dist\ACQUA_ReportReviewer_v$version.exe" -ForegroundColor Green
