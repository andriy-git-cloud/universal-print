#Add printers from a CSV file filled with printer name and IP
#Printers will be added by using IPP protocol. If needed the printer will be renamed accordingly to CSV
#Log file will be generated into C:\UP
#Created by Andrii Zadorozhnyi (andrii.zadorozhnyi@wfp.org)
#version 1.3
#


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

# Function to check printer connectivity
function Test-PrinterConnection {
    param ([string]$printerIP)
    
    $pingResult = Test-Connection -ComputerName $printerIP -Count 4 -Quiet
    return $pingResult
}


$ImportFile = "C:\UP\PrintersIPP.csv"

# Import printers from CSV file
$printers = import-csv $ImportFile -Delimiter ";"

Write-Log "Printers creation... Started" -Level "SUCCESS"

foreach ($printer in $printers) {
    Write-Log "<--------"
    $startTime = Get-Date
	$startTime
    Write-Log "Adding the printer $($printer.Name)..."


    # Add printer and check success


	# Ping the printer before adding it
	if (Test-PrinterConnection -printerIP $printer.IP) {
		

	$addPrinterSuccess = $false
    try {
        Add-Printer -Name $printer.Name -IppURL http://$($printer.IP):631/ipp/print -ErrorAction Stop
        $addPrinterSuccess = $true
    } catch {
        Write-Log "Failed to add printer $($printer.Name), waiting for Event ID 300..." -Level "WARNING"
    
        $maxWaitTime = 90  # Maximum wait time in seconds (1.5 minutes)
        $waitInterval = 5    # Check every 5 seconds
        $elapsedTime = 0

        while ($elapsedTime -lt $maxWaitTime) {
            Start-Sleep -Seconds $waitInterval
            $elapsedTime += $waitInterval

            # Check for Event ID 300
            $event = Get-WinEvent -LogName "Microsoft-Windows-PrintService/Operational" |
                     Where-Object { $_.Id -eq 300 -and $_.TimeCreated -gt $startTime } |
                     Sort-Object TimeCreated -Descending |
                     Select-Object -First 1

            if ($event) {
                if ($event.Message -match "Printer\s+(.*?)\s+was created") {
                    $lastAddedPrinter = $matches[1]
                    Write-Log "Printer '$lastAddedPrinter' was registered after $elapsedTime seconds." -Level "SUCCESS"
                    $addPrinterSuccess = $true
                    break  # Exit loop early if the event appears
                }
            }

            Write-Log "Waiting... ($elapsedTime seconds elapsed)"
        }

        # If we reach the max wait time and no event was found
        if (-not $addPrinterSuccess) {
            Write-Log "No Event ID 300 found after $maxWaitTime seconds." -Level "ERROR"
        }
    }


    if ($addPrinterSuccess) {
        # Poll for Event ID 300 instead of using Start-Sleep
        $event = $null
        while (-not $event) {
            Start-Sleep -Seconds 1
            $event = Get-WinEvent -LogName "Microsoft-Windows-PrintService/Operational" |
                     Where-Object { $_.Id -eq 300 -and $_.TimeCreated -gt $startTime } |
                     Sort-Object TimeCreated -Descending |
                     Select-Object -First 1
        }

        if ($event) {
            # Extract the printer name from the event message
            if ($event.Message -match "Printer\s+(.*?)\s+was created") {
                $lastAddedPrinter = $matches[1]
                Write-Log "Last installed printer: '$lastAddedPrinter' at $($event.TimeCreated)" -Level "SUCCESS"
            } else {
                Write-Log "Could not extract printer name from the event message."
            }
        } else {
            Write-Log "No printer installation events found within timeout period." -Level "ERROR"
            continue  # Skip renaming step if no event was found
        }

        Write-Log "Printer '$lastAddedPrinter' added - Port: $($printer.IP)" -Level "SUCCESS"

        # Rename the printer only if necessary
        if ($lastAddedPrinter -ne $printer.Name) {
            Rename-Printer -Name $lastAddedPrinter -NewName $printer.Name
            Write-Log "Printer renamed from '$lastAddedPrinter' to '$($printer.Name)'" -Level "INFO"
        }

    }
	} else {
		Write-Log "Printer $($printer.Name) at $($printer.IP) is unreachable. Skipping installation." -Level "ERROR"
	}
    Write-Log "         ----->"
}

Write-Log "Printers creation... Completed" -Level "SUCCESS"
# Record script end time and calculate execution duration
$scriptEndTime = Get-Date
$executionTime = $scriptEndTime - $scriptStartTime
Write-Log "Script execution time: $executionTime" -Level "INFO"
