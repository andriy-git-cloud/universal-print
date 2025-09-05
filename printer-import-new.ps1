# Define log file path
$LogFile = "C:\UP\PrinterScript.log"

# Function to log messages
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )

    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "$Timestamp [$Level] $Message"
	
	# Choose color based on log level
    switch ($Level) {
        "INFO"     { $Color = "White" }
        "SUCCESS"  { $Color = "Green" }
        "WARNING"  { $Color = "Yellow" }
        "ERROR"    { $Color = "Red" }
        default    { $Color = "Gray" }
    }
    
    # Write to console
    Write-Host $LogEntry -ForegroundColor $Color
    
    # Append to log file
    Add-Content -Path $LogFile -Value $LogEntry
}

# Start logging
Write-Log "Script execution started."

# Record script start time
$scriptStartTime = Get-Date

$ImportFile = "C:\UP\Printers.csv"

# Import the printers that are going to be imported

$printers = import-csv $ImportFile -Delimiter ";"

Write-Log "Printers creation... Started" -Level "SUCCESS"
foreach ($printer in $printers) {

	# Check if port already exist

	try {
		$portName = $printer.IP
		$checkPortExists = Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue
		if (-not $checkPortExists) {
			Add-PrinterPort -Name $portName -PrinterHostAddress $portName
			Write-Log "TCP/IP Port '$portName' created." -Level "SUCCESS"
		} else {
			Write-Log "TCP/IP Port '$portName' already exists." -Level "ERROR"
		}

	# Add the printer using the TCP/IP port
		Add-Printer -Name $printer.Name -DriverName $printer.Driver -PortName $portName
		Write-Log "Printer '$($printer.Name)' added Port '$portName' - Driver : '$($printer.Driver)'."  -Level "SUCCESS"
		} catch {
			Write-Log "Failed to add printer using TCP/IP Port: $_" -Level "ERROR"
		}

}

Write-Log "Printers creation... Completed" -Level "SUCCESS"
# Record script end time and calculate execution duration
$scriptEndTime = Get-Date
$executionTime = $scriptEndTime - $scriptStartTime
Write-Log "Script execution time: $executionTime" -Level "INFO"