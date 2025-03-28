# PrinterImport-IPP PowerShell Script

## Description

This PowerShell script automates the process of adding printers using the **IPP (Internet Printing Protocol)**. It reads a CSV file containing printer names and IP addresses, then installs and renames them accordingly. A log file is generated to track the process.

## Features

- Imports printers from a CSV file
- Adds printers using IPP protocol
- Renames printers as per the CSV data
- Generates a log file at `C:\UP\PrinterScript.log`

## Requirements

- Windows operating system
- PowerShell 5.1 or later
- Administrator privileges
- CSV file containing printer details (Name, IP Address)

## Installation & Usage

### 1. Prepare the CSV File

Ensure you have a CSV file structured as follows: (delimiter is a semi-column!!)

```csv
PrinterName;IPAddress
Printer1;192.168.1.10
Printer2;192.168.1.11
```

### 2. Run the Script

1. Open **PowerShell as Administrator**.
2. Navigate to the script's directory:
   ```powershell
   cd C:\Path\To\Script
   ```
3. Execute the script:
   ```powershell
   .\PrinterImport-IPP.ps1 -CsvPath "C:\Path\To\printers.csv"
   ```

## Logging

The script logs its actions in `C:\UP\PrinterScript.log`, recording any errors or successful installations.

## Author

Created by **Andrii Zadorozhnyi** ([andrii.zadorozhnyi@wfp.org](mailto\:andrii.zadorozhnyi@wfp.org))

## Version

**1.3**

## License

This script is provided as-is. Use it at your own risk.

