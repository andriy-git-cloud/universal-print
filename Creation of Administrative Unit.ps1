# Creation of Administrative Unit
Connect-MgGraph -Scopes "AdministrativeUnit.ReadWrite.All"

# Define the path to your CSV file
$csvPath = "C:\temp\AdministrativeUnits.csv"

# Import the CSV file into a variable
$csvData = Import-Csv -Path $csvPath

# Loop through each row in the CSV to create Administrative Units and Dynamic Memberships
foreach ($row in $csvData) {
    $AUName = $row.AUName
    $AUDescription = $row.AUDescription
    $deviceMembershipRule = $row.DeviceMembershipRule

    # Step 1: Create the Administrative Unit
    Write-Host "Creating Administrative Unit: $AUName"

    $params = @{
        displayName = $AUName
        description = $AUDescription
        membershipType = "Dynamic"
        membershipRule = $deviceMembershipRule
        membershipRuleProcessingState = "On"
    }

    New-MgDirectoryAdministrativeUnit -BodyParameter $params


    Write-Host "Finished creating AU and Dynamic Group for $AUName`n"
}

Write-Host "Script execution completed!"