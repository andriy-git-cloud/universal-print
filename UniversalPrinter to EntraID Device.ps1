$ErrorActionPreference = "Stop"
Connect-MgGraph -Scopes "Directory.AccessAsUser.All", "Printer.Read.All"

$tenantId = (Get-MgContext).TenantId
Write-Host "Starting processing of Universal Print printers in tenant $tenantId"

# This streams pages of printers and does not require them to all be loaded at once.
Get-MgPrintPrinter -All -ExpandProperty "connectors" | ForEach-Object -Process {
    $printer = $_

    Write-Host "Fetching Azure AD device for printer $($printer.DisplayName)"
    $device = Get-MgDevice -Filter "deviceId eq '$($printer.Id)'" -Top 1

    # The display name of the Azure AD device is set to the initial display name
    # of the printer. This sets extensionAttribute1 to the current name.
    $extensionAttribute1 = "$($printer.DisplayName)"

    # If the printer was registered with the Universal Print connector then the
    # display name of the connector will be present in extensionAttribute2.
    $extensionAttribute2 = "$($printer.Connectors[0].DisplayName)"

    # If the printer has a country or region set in its location properties it
    # will be set to extensionAttribute15. Other location properties can be used
    # as well.
    $extensionAttribute3 = "$($printer.Location.CountryOrRegion)"

    $existingExtensionAttributes = $device.AdditionalProperties.extensionAttributes
    if ($extensionAttribute1 -ne "$($existingExtensionAttributes.extensionAttribute1)" -or
        $extensionAttribute2 -ne "$($existingExtensionAttributes.extensionAttribute2)" -or
        $extensionAttribute3 -ne "$($existingExtensionAttributes.extensionAttribute3)")
    {
        Write-Host "Updating Azure AD device extension attributes for printer $($printer.DisplayName)"
        Update-MgDevice -DeviceId $device.Id -BodyParameter @{
            "extensionAttributes" = @{
                "extensionAttribute1" = $extensionAttribute1
                "extensionAttribute2" = $extensionAttribute2
                "extensionAttribute3" = $extensionAttribute3
            }
        }
    }
}