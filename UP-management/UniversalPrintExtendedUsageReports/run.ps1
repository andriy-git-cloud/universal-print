Param (
    $Timer, $TriggerMetadata
    # Date should be in this format: 2020-09-01
    # Default is the first day of the previous month at 00:00:00 (Tenant time zone)
    #$StartDate = "",
    # Date should be in this format: 2020-09-30
    # Default is the last day of the previous month 23:59:59 (Tenant time zone)
    #$EndDate = "",
    # Set if only the Printer report should be generated
    #[switch]$PrinterOnly,
    # Set if only the User report should be generated
    #[switch]$UserOnly
)

# Description: This script connects to Microsoft Graph using the Universal Print API and retrieves extended usage reports for printers.

$tenantId     = $env:TenantId
$clientId     = $env:ClientId
$clientSecret = $env:ClientSecret
$StorageAccountName   = $env:STORAGE_ACCOUNT_NAME
#$ResourceGroupName    = $env:RESOURCE_GROUP_NAME
$ContainerName        = $env:STORAGE_CONTAINER_NAME  # e.g., "print-exports"

$secureSecret = ConvertTo-SecureString -String $clientSecret -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential($clientId, $secureSecret)

#############################
# IMPORT MODULES
#############################

Import-Module Microsoft.Graph.Reports -ErrorAction Stop
Import-Module Az.Accounts  -ErrorAction Stop
Import-Module Az.Storage   -ErrorAction Stop

# -----------------------------
# SET DATE RANGE (null-safe)
# -----------------------------
# Optional overrides via App Settings:
#   START_DATE = 2025-08-01
#   END_DATE   = 2025-08-31
$StartDateRaw = if ($env:START_DATE) { $env:START_DATE } else { $StartDate }
$EndDateRaw   = if ($env:END_DATE)   { $env:END_DATE }   else { $EndDate   }

function Get-PrevMonthStartIsoUtc { (Get-Date -Day 1).AddMonths(-1).ToString('yyyy-MM-ddT00:00:00Z') }
function Get-PrevMonthEndIsoUtc   { (Get-Date -Day 1).AddDays(-1).ToString('yyyy-MM-ddT23:59:59Z') }

# START
if ([string]::IsNullOrWhiteSpace($StartDateRaw)) {
    $StartDateIso = Get-PrevMonthStartIsoUtc
} else {
    try {
        $sd = Get-Date -Date $StartDateRaw -ErrorAction Stop
        $StartDateIso = $sd.ToString('yyyy-MM-ddT00:00:00Z')
    } catch {
        throw "Invalid START_DATE value '$StartDateRaw'. Expected e.g. 2025-08-01"
    }
}

# END
if ([string]::IsNullOrWhiteSpace($EndDateRaw)) {
    $EndDateIso = Get-PrevMonthEndIsoUtc
} else {
    try {
        $ed = Get-Date -Date $EndDateRaw -ErrorAction Stop
        $EndDateIso = $ed.ToString('yyyy-MM-ddT23:59:59Z')
    } catch {
        throw "Invalid END_DATE value '$EndDateRaw'. Expected e.g. 2025-08-31"
    }
}

Write-Host "Gathering reports between $StartDateIso and $EndDateIso."


# Label for filenames: the month we are reporting (previous month)
$monthLabel = (Get-Date -Day 1).AddMonths(-1).ToString('yyyy-MM')

# Temp file paths
$printerCsv = Join-Path $env:TMP ("printerReport_{0}.csv" -f $monthLabel)
$userCsv    = Join-Path $env:TMP ("userReport_{0}.csv" -f $monthLabel)

########################################
# SIGN IN & CONNECT TO MICROSOFT GRAPH
########################################

# These scopes are needed to get the list of users, list of printers, and to read the reporting data.
Connect-MgGraph -ClientSecretCredential $cred -TenantId $tenantId #-Scopes "Directory.AccessAsUser.All", "Printer.Read.All"
Connect-AzAccount -ServicePrincipal -Tenant $tenantId -Credential $cred | Out-Null


##########################
# GET PRINTER REPORT
##########################

if (!$UserOnly)
{
    Write-Progress -Activity "Gathering Printer usage..." -PercentComplete -1

# Get the printer usage report
    $printerReport = Get-MgReportMonthlyPrintUsageByPrinter -All -Filter "completedJobCount gt 0 and usageDate ge $StartDateIso and usageDate lt $EndDateIso"


## Join extended printer info with the printer usage report
    $reportWithPrinterNames = $printerReport | 
        Select-Object ( 
            @{Name = "UsageMonth"; Expression = {$_.Id.Substring(0,8)}}, 
            @{Name = "PrinterId"; Expression = {$_.PrinterId}}, 
            @{Name = "DisplayName"; Expression = {$_.PrinterName}}, 
            @{Name = "TotalJobs"; Expression = {$_.CompletedJobCount}},
            @{Name = "ColorJobs"; Expression = {$_.CompletedColorJobCount}},
            @{Name = "BlackAndWhiteJobs"; Expression = {$_.CompletedBlackAndWhiteJobCount}},
            @{Name = "ColorPages"; Expression = {$_.ColorPageCount}},
            @{Name = "BlackAndWhitePages"; Expression = {$_.BlackAndWhitePageCount}},
            @{Name = "TotalSheets"; Expression = {$_.MediaSheetCount}})

# Write the aggregated report CSV
    $reportWithPrinterNames | Export-Csv -Path $printerCsv -NoTypeInformation -Encoding utf8
}

##################
# GET USER REPORT
##################

if (!$PrinterOnly)
{
    Write-Progress -Activity "Gathering User usage..." -PercentComplete -1

# Get the user usage report
    $userReport = Get-MgReportMonthlyPrintUsageByUser -All -Filter "completedJobCount gt 0 and usageDate ge $StartDateIso and usageDate lt $EndDateIso"

    $reportWithUserInfo = $userReport | 
        Select-Object ( 
            @{Name = "UsageMonth"; Expression = {$_.Id.Substring(0,8)}}, 
            @{Name = "UserPrincipalName"; Expression = {$_.UserPrincipalName}}, 
            @{Name = "TotalJobs"; Expression = {$_.CompletedJobCount}},
            @{Name = "ColorJobs"; Expression = {$_.CompletedColorJobCount}},
            @{Name = "BlackAndWhiteJobs"; Expression = {$_.CompletedBlackAndWhiteJobCount}},
            @{Name = "ColorPages"; Expression = {$_.ColorPageCount}},
            @{Name = "BlackAndWhitePages"; Expression = {$_.BlackAndWhitePageCount}},
            @{Name = "TotalSheets"; Expression = {$_.MediaSheetCount}})
			        
    $reportWithUserInfo | Export-Csv -Path $userCsv -NoTypeInformation -Encoding utf8
}

# -----------------------------
# UPLOAD TO BLOB STORAGE
# -----------------------------

# Use the currently connected SP for AAD-auth to Blob
$ctx = New-AzStorageContext -StorageAccountName $storageAccountName -UseConnectedAccount


if (Test-Path $printerCsv) {
    $printerBlobPath = (Split-Path $printerCsv -Leaf)
    Set-AzStorageBlobContent -File $printerCsv -Container $ContainerName -Blob $printerBlobPath -Context $ctx -Force | Out-Null
    Write-Host "Uploaded $printerBlobPath"
}

if (Test-Path $userCsv) {
    $userBlobPath = (Split-Path $userCsv -Leaf) 
    Set-AzStorageBlobContent -File $userCsv -Container $ContainerName -Blob $userBlobPath -Context $ctx -Force | Out-Null
    Write-Host "Uploaded $userBlobPath"
}

Write-Host "Done."