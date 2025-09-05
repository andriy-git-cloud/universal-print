# Description: This script connects to Microsoft Graph using the Universal Print API and retrieves extended usage reports for printers.

$tenantId     = $env:TenantId
$clientId     = $env:ClientId
$clientSecret = $env:ClientSecret
$StorageAccountName   = $env:STORAGE_ACCOUNT_NAME
$ResourceGroupName    = $env:RESOURCE_GROUP_NAME
$ContainerName        = $env:STORAGE_CONTAINER_NAME  # e.g., "print-exports"

$secureSecret = ConvertTo-SecureString -String $clientSecret -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential($clientId, $secureSecret)

Param (
    # Date should be in this format: 2020-09-01
    # Default is the first day of the previous month at 00:00:00 (Tenant time zone)
    #$StartDate = "",
    # Date should be in this format: 2020-09-30
    # Default is the last day of the previous month 23:59:59 (Tenant time zone)
    $EndDate = "",
    # Set if only the Printer report should be generated
    [switch]$PrinterOnly,
    # Set if only the User report should be generated
    [switch]$UserOnly
)

#############################
# INSTALL & IMPORT MODULES
#############################

Import-Module Microsoft.Graph.Reports

#############################
# SET DATE RANGE
#############################
if ($StartDate -eq "") {
    $StartDate = (Get-Date -Day 1).AddMonths(-1).ToString("yyyy-MM-ddT00:00:00Z")
} else {
    $StartDate = ([DateTime]$StartDate).ToString("yyyy-MM-ddT00:00:00Z")
}

if ($EndDate -eq "") {
    $EndDate = (Get-Date -Day 1).AddDays(-1).ToString("yyyy-MM-ddT23:59:59Z")
} else {
    $EndDate = ([DateTime]$EndDate).ToString("yyyy-MM-ddT23:59:59Z")
}

Write-Output "Gathering reports between $StartDate and $EndDate."

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

##########################
# GET PRINTER REPORT
##########################

if (!$UserOnly)
{
    Write-Progress -Activity "Gathering Printer usage..." -PercentComplete -1

# Get the printer usage report
    $printerReport = Get-MgReportMonthlyPrintUsageByPrinter -All -Filter "completedJobCount gt 0 and usageDate ge $StartDate and usageDate lt $EndDate"

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
    $userReport = Get-MgReportMonthlyPrintUsageByUser -All -Filter "completedJobCount gt 0 and usageDate ge $StartDate and usageDate lt $EndDate"
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
$sa = Get-AzStorageAccount -Name $StorageAccountName -ResourceGroupName $ResourceGroupName
$ctx = $sa.Context

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