# Default input file path
$defaultInput = "C:\temp\input.csv"
$inputPrompt = "Enter the full path to the CSV input file [$defaultInput]"

# Read input, use default if blank
$InputFile = Read-Host $inputPrompt
if ([string]::IsNullOrWhiteSpace($InputFile)) {
    $InputFile = $defaultInput
}

# Check if the file exists
if (-not (Test-Path $InputFile)) {
    Write-Host "❌ File not found: $InputFile" -ForegroundColor Red
    exit
}

# Default output file path
$defaultOutput = "C:\temp\printer_output.csv"

# Read input, use default if blank
$outputPrompt = "Enter the full path where the output CSV should be saved [$defaultOutput]"
$OutputFile = Read-Host $outputPrompt
if ([string]::IsNullOrWhiteSpace($OutputFile)) {
    $OutputFile = $defaultOutput
}

# Ask user if they want to skip the first row
$skipFirst = Read-Host "Do you want to skip the first row (e.g. example row)? [Y/N]"

# Import CSV accordingly
if ($skipFirst -match '^[Yy]') {
    $csv = Import-Csv -Path $InputFile | Select-Object -Skip 1
} else {
    $csv = Import-Csv -Path $InputFile
}

# Mapping logic for driver based on model keywords
function Get-PrinterDriver($model) {
    switch -Regex ($model) {
        "HP|Hewlett Packard" { return "HP Universal Printing PCL 6" }
        "Xerox"              { return "Xerox Global Print Driver PCL6" }
        "Kyocera"            { return "KX DRIVER for Universal Printing" }
        "Ricoh"              { return "PCL6 V4 Driver for Universal Print" }
        "Canon"              { return "Canon Generic Plus PCL6" }
        "Brother"            { return "Brother Mono Universal Printer (PCL)" }
        "Epson"              { return "EPSON Universal Print Driver" }
        "Konica Minolta"     { return "KONICA MINOLTA Universal v4 PCL" }
        default              { return "" }
    }
}

# Build output objects
$printers = foreach ($row in $csv) {
    if ($row.'IP address' -and $row.'Printer name' -and $row.'Printer model') {
        [PSCustomObject]@{
            IP     = $row.'IP address'.Trim()
            Name   = $row.'Printer name'.Trim()
            Driver = Get-PrinterDriver $row.'Printer model'
        }
    }
}

# Export to CSV with semicolon delimiter
$printers | Export-Csv -Path $OutputFile -NoTypeInformation -Delimiter ';'

Write-Host "`n✅ CSV created at: $OutputFile"
