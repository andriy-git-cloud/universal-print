param (
    [string]$InputFile = "C:\Path\To\Your\AzureUniversalPrint_RequestForm_template.csv",
    [string]$OutputFile = "C:\Path\To\printers_output.csv"
)

# Read the CSV, skipping the first row (example row)
$csv = Import-Csv -Path $InputFile | Select-Object -Skip 1

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

Write-Host "CSV created at: $OutputFile"
