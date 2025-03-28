#Add printers from a CSV file filled with printer name and IP
#Printers will be added by using IPP protocol. If needed the printer will be renamed accordingly to CSV
#Created by Andrii Zadorozhnyi (andrii.zadorozhnyi@wfp.org)
#version 1.2
#
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

Write-Host "Printers creation... Started" -ForegroundColor Green

foreach ($printer in $printers) {
    Write-Host "<--------"
    $startTime = Get-Date
	$startTime
    Write-Host "Adding the printer $($printer.Name)..."


    # Add printer and check success


	# Ping the printer before adding it
	if (Test-PrinterConnection -printerIP $printer.IP) {
		

	$addPrinterSuccess = $false
    try {
        Add-Printer -Name $printer.Name -IppURL http://$($printer.IP):631/ipp/print -ErrorAction Stop
        $addPrinterSuccess = $true
    } catch {
        Write-Host "Failed to add printer $($printer.Name), waiting for Event ID 300..." -ForegroundColor Yellow
    
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
                    Write-Host "Printer '$lastAddedPrinter' was registered after $elapsedTime seconds." -ForegroundColor Green
                    $addPrinterSuccess = $true
                    break  # Exit loop early if the event appears
                }
            }

            Write-Host "Waiting... ($elapsedTime seconds elapsed)" -ForegroundColor DarkGray
        }

        # If we reach the max wait time and no event was found
        if (-not $addPrinterSuccess) {
            Write-Host "No Event ID 300 found after $maxWaitTime seconds." -ForegroundColor Red
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
                Write-Host "Last installed printer: '$lastAddedPrinter' at $($event.TimeCreated)" -ForegroundColor Green
            } else {
                Write-Host "Could not extract printer name from the event message."
            }
        } else {
            Write-Host "No printer installation events found within timeout period." -ForegroundColor Red
            continue  # Skip renaming step if no event was found
        }

        Write-Host "Printer '$lastAddedPrinter' added - Port: $($printer.IP)"

        # Rename the printer only if necessary
        if ($lastAddedPrinter -ne $printer.Name) {
            Rename-Printer -Name $lastAddedPrinter -NewName $printer.Name
            Write-Host "Printer renamed from '$lastAddedPrinter' to '$($printer.Name)'" -ForegroundColor Yellow
        }

    }
	} else {
		Write-Host "Printer $($printer.Name) at $($printer.IP) is unreachable. Skipping installation." -ForegroundColor Red
	}
    Write-Host "         ----->"
}

Write-Host "Printers creation... Completed" -ForegroundColor Green
# Record script end time and calculate execution duration
$scriptEndTime = Get-Date
$executionTime = $scriptEndTime - $scriptStartTime
Write-Host "Script execution time: $executionTime" -ForegroundColor Yellow
