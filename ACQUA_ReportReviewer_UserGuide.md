# ACQUA Report Reviewer - User Guide

**Version:** 1.1.0  
**Author:** Jian Zou  
**Last Updated:** January 30, 2026

---

## Table of Contents

1. [Overview](#overview)
2. [System Requirements](#system-requirements)
3. [Getting Started](#getting-started)
4. [How to Use](#how-to-use)
5. [Output Reports](#output-reports)
6. [Report Sections Explained](#report-sections-explained)
7. [Troubleshooting](#troubleshooting)
8. [FAQ](#faq)

---

## Overview

ACQUA Report Reviewer is a Python-based tool (also available as a standalone Windows .exe) that processes ACQUA audio test report Word documents (.docx) and extracts key test data, validation results, and status information into an organized summary report.

### Key Features

- **Multi-file Processing**: Select and process multiple ACQUA report files at once
- **Test Time Analysis**: Calculate test duration by category and overall
- **Test Case Validation**: Verify required test cases for Shared Speakerphone, Headset, Open Office Headset, Handset, and Personal/Desktop Speakerphone
- **ACQUA & Database Version Tracking**: Extract software and database versions from reports
- **Equipment Settings Extraction**: Display labCORE, HATS/HMS, and BEQ configuration
- **54dB Noise Scenario Analysis**: Compare NS ON vs NS OFF results
- **Status Overview**: Identify "Not OK" entries across all reports
- **Double Talk Performance**: Extract attenuation measurements
- **CSV Export**: All results exported to a single CSV file for further analysis
- **Long file name support**: Output tables now support up to 100-character file names
- **Packaged as .exe**: Build script included for easy packaging

---

## System Requirements

- **Operating System**: Windows 10 or later
- **Python 3.8+** (if running from source)
- **Required Python packages**: python-docx, tkinter
- **No Python required** if using the packaged .exe

---

## Getting Started

### Using the Standalone Executable

1. Run `build.ps1` in PowerShell to generate the .exe (requires Python and PyInstaller):
   ```
   .\build.ps1
   ```
2. Find the generated `process_acqua_reports.exe` in the `dist` folder.
3. Double-click the .exe or run it from the command line.

### Using Python Source

1. Clone or download the repository.
2. Open a terminal in the project directory.
3. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
4. Run the script:
   ```
   python process_acqua_reports.py
   ```

---

## How to Use

### Step 1: Select Files

When the application starts, a file selection dialog will appear:

1. Navigate to the folder containing your ACQUA report files (.docx)
2. Select one or more Word documents
   - Hold `Ctrl` to select multiple individual files
   - Hold `Shift` to select a range of files
3. Click **Open** to begin processing

### Step 2: View Results

The application will display progress and results in the console window:

- Processing status for each file
- Summary tables with extracted data
- Validation results for all supported device types

### Step 3: Review Output File

After processing completes, a CSV file named `Smd_Report_Output.csv` is automatically saved in the **same folder** as the first selected file.

---

## Output Reports

### Console Output

The console displays formatted tables with the following sections:

1. Test Time by Category
2. Overall Test Duration
3. Device Type Validation (Shared Speakerphone, Headset, Open Office Headset, Handset, Personal/Desktop Speakerphone)
4. ACQUA & Teams Database Information
5. Test Case Validation
6. 54dB Noise Scenario Results
7. Status Overview: 'Not OK' Entries
8. Double Talk Performance

### CSV Output (Smd_Report_Output.csv)

All console output is also saved to a CSV file for:
- Import into Excel for further analysis
- Archiving test results
- Sharing with team members

---

## Report Sections Explained

### Test Time by Category

Shows the duration of testing for each test category:

| Category | Description |
|----------|-------------|
| P-series (A) | AR (Audio Rendering) tests |
| P-series (R) | RR (Recording/Rendering) tests |
| Device-direct | Di series tests |
| Option codes | OpO, OpM, OpS, OpP tests |
| Custom | Non-standard test codes |

Includes:
- Number of tests per category
- Start and end times
- Total duration
- Daily breakdown for overnight testing

### Overall Test Duration

Aggregates all tests regardless of category, showing:
- Tests per day with time ranges
- Total test count and cumulative duration
- Multi-day testing summary

### ACQUA & Teams Database Information

Displays for each processed file:
- **ACQUA Version**: Software version (e.g., "ACQUA 6.0.200")
- **Database Version**: Teams database revision (e.g., "51_MS_Teams_Rev05_SP2")
- **File**: Up to 100 characters shown in output

### Device Type Validation

Validation output is provided for:
- Shared Speakerphone
- Headset
- Open Office Headset
- Handset
- Personal/Desktop Speakerphone

Each section lists required test cases for AR, RR, Di, and Op categories as appropriate, and highlights missing or complete sets.

### 54dB Noise Scenario Results

Compares noise suppression performance:

| Field | Description |
|-------|-------------|
| Device | Device name from filename (up to 100 chars) |
| Lab | Test lab (AST or PAL) |
| Report Time | Test date |
| NS Setting | Noise Suppression ON or OFF |
| SMOS/NMOS/GMOS | MOS scores for 2nd Talker and BGN scenarios |

### Status Overview: 'Not OK' Entries

Lists all test items that did not pass, including:
- File name (up to 100 chars)
- SMD (test description)
- Specific issues or failed criteria

### Double Talk Performance

Extracts attenuation measurements during double talk scenarios:
- SMD description
- Status (OK/Not OK)
- Single value in dB

---

## Troubleshooting

### "No files selected. Exiting."
- **Cause**: You clicked Cancel in the file dialog
- **Solution**: Re-run the application and select at least one .docx file

### "No matching 'SmdTitle' or 'SmdDate' styles found"
- **Cause**: The selected files may not be valid ACQUA reports
- **Solution**: Ensure you're selecting Word documents exported from ACQUA

### "Error: Could not write to Smd_Report_Output.csv"
- **Cause**: The CSV file is open in another program (e.g., Excel)
- **Solution**: Close Excel and run the application again

### "Failed to process file: [filename]"
- **Cause**: The file may be corrupted or in an unsupported format
- **Solution**: Verify the file opens correctly in Microsoft Word

### Console window closes immediately
- **Cause**: An unexpected error occurred
- **Solution**: Run from Command Prompt to see error details:
  ```cmd
  cd C:\path\to\folder
  process_acqua_reports.exe
  ```

---

## FAQ

**Q: Can I process files from different folders?**  
A: Yes, you can select files from any location. The output CSV will be saved in the folder of the first selected file.

**Q: What file formats are supported?**  
A: Only Microsoft Word documents (.docx) exported from ACQUA are supported. Legacy .doc files are not supported.

**Q: How do I compare results from different test runs?**  
A: Each run creates a new Smd_Report_Output.csv. Rename or move previous output files before running again to preserve them.

**Q: Can I automate this tool from command line?**  
A: The tool uses a graphical file picker by default. For automation, use the Python script with command-line arguments (see --help).

**Q: What if some tests show as "Custom" category?**  
A: This means the test code wasn't recognized as a standard ACQUA code. The tool will still extract and display the data.

---

## Support

For questions or issues, contact:
- **Author**: Jian Zou
- **Email**: jianzou@microsoft.com

---

*© 2026 Microsoft Corporation. For internal use only.*
