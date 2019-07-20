<#
.SYNOPSIS
    Script to obtain Queensland Registration Information.
.DESCRIPTION
    Simple script that will query the TMR database (via web session) and get details of current registration.
.PARAMETER Rego
    Registration/Licence Plate number.
.INPUTS
    Requires a valid QLD Registration/Licence Plate.
.OUTPUTS
    Returns an object containing the following properties Registration Number, VIN, Description, Purpose, Current Status, Expiry Date, Days until expiry.
.NOTES
  Version:        0.1.0
  Author:         Dominic P.
  Creation Date:  20/07/2019
  Purpose/Change: Initial script development.
  
.EXAMPLE
    Get-QldRegistrationDetails -Rego ABC123
#>

param(
    # Parameter help description
    [Parameter(Mandatory=$true,Position=0)]
    [string]
    $Rego
)

#----------------------------------------------------------[Declarations]----------------------------------------------------------

# Create output object and define list of properties.
$vehicleRegistration = New-Object PSObject
$properties = ("ResistrationNumber","VehicleIdentificationNumber","Description","PurposeOfUse","Status","ExpiryDate")

$initalUrl = "https://www.service.transport.qld.gov.au/checkrego/application/TermAndConditions.xhtml"
$searchUrl = "https://www.service.transport.qld.gov.au/checkrego/application/VehicleSearch.xhtml"

#-----------------------------------------------------------[Functions]------------------------------------------------------------


#-----------------------------------------------------------[Execution]------------------------------------------------------------


try {

    # Get session details from app.
    $response = Invoke-WebRequest -Uri $initalUrl -UseBasicParsing -SessionVariable "RegoSession"

    $viewstate = [System.Uri]::EscapeDataString("$($response.InputFields.Find("javax.faces.ViewState").value)")
    $clientWindow = [System.Uri]::EscapeDataString("$($response.InputFields.Find("javax.faces.ClientWindow").value)")

    # Generate request body for acceptance of T&Cs.
    $body = "tAndCForm%3AconfirmButton=&tAndCForm_SUBMIT=1&javax.faces.ViewState=$viewstate&javax.faces.ClientWindow=$clientWindow"

    # Accept T&Cs.
    $response = Invoke-WebRequest -Uri $initalUrl -UseBasicParsing -Body $body -Method "POST" -WebSession $RegoSession

    # Get new session details post T&C acceptance.
    $viewstate = [System.Uri]::EscapeDataString("$($response.InputFields.Find("javax.faces.ViewState").value)")
    $clientWindow = [System.Uri]::EscapeDataString("$($response.InputFields.Find("javax.faces.ClientWindow").value)")

    # Generate request body for rego search.
    $body = "vehicleSearchForm%3AplateNumber=$Rego&vehicleSearchForm%3AreferenceId=&vehicleSearchForm%3AconfirmButton=&vehicleSearchForm_SUBMIT=1&javax.faces.ViewState=$viewState&javax.faces.ClientWindow=$clientWindow"

    # Request registration details from app.
    $response = Invoke-WebRequest -Uri $searchUrl -UseBasicParsing -Body $body -Method "POST" -WebSession $RegoSession

} catch {
    Write-Host -ForegroundColor Red "$($_.Exception)"
    return
}

try { # Issues with NOT using basic parsing above causes freezing, so instantiating HTMLFile Com object for easier parsing (may not be availible on all PoSh versions).

    $htmlDoc = New-Object -ComObject "HTMLFile"
    $htmlDoc.IHTMLDocument2_write($response.Content)

    $dataElements = $htmlDoc.getElementsByTagName("DD")

    # Iterate through values and populate object data.
    0..$(($properties.Count)-1) | ForEach-Object {
            $vehicleRegistration | Add-Member -MemberType NoteProperty -Name $properties[$_] -Value $(if ($dataElements[$_]){$dataElements[$_].InnerText.Trim()} else {""})
    }

} catch{ # If unable to instantiate Com object (i.e. PowerShell Core/Linux) above use basic string parsing to capture registration data.

    # Begin processing response.
    $split = $response -split '<dl class="data">'

    # Grab both sets of <dl> tags - Set1: Registration Number, VIN and Description - Set2: Status, Purpose, Expiry.
    $set1 = $split[1]
    $set2 = $($split[2] -split '</dl>')[0]

    # Convert from HTML to basic string Key:Value pairs.
    $data01 = $($($($($set1 -replace '\s+',' ') -replace ' </dt> <dd>','=') -replace ' </dd> <dt>',"`n") -replace '<[^>]+>','')
    $data02 = $($($($($set2 -replace '\s+',' ') -replace " </dt> <dd>","=") -replace " </dd> <dt>","`n") -replace '<[^>]+>','')

    # Convert from String to Hashmap
    $respData = ConvertFrom-StringData "$data01`n$data02"

    # Description values used on the web page.
    $appDescriptionKeys = "Registration Number","Vehicle Identification Number (VIN)","Description","Purpose of use","Status","Expiry"

    # Iterate through values and populate object data.
    0..$(($properties.Count)-1) | ForEach-Object {
        $vehicleRegistration | Add-Member -MemberType NoteProperty -Name $properties[$_] -Value $(if ($respData."$($appDescriptionKeys[$_])"){$respData."$($appDescriptionKeys[$_])".Trim()} else {""})
    }
}

# Calculate Expiry Date - else populate member with null.
$vehicleRegistration | Add-Member -MemberType NoteProperty -Name DaysToExpiry -Value $(if ($vehicleRegistration.ExpiryDate) {$([datetime]::ParseExact($vehicleRegistration.ExpiryDate,"dd/MM/yyyy",$null)-$(Get-Date)).Days} else {""})

return $vehicleRegistration
