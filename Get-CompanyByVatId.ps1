<#
.SYNOPSIS
    Retrieves company information based on VAT ID from the PRH API and processes multiple VAT IDs from an input file.

.DESCRIPTION
    This script fetches company information from the PRH API using the provided VAT ID.
    It validates the VAT ID format, makes an API request, and processes the response to extract
    relevant company details such as name, visiting address, and postal address.
    The script can process multiple VAT IDs from an input file and optionally output the results to a CSV file.

.PARAMETER InputFile
    The path to the input file containing VAT IDs, one per line.

.PARAMETER OutputFile
    The path to the output CSV file where the company information will be saved. This parameter is optional.

.EXAMPLE
    PS> .\Get-CompanyByVatId.ps1 -InputFile 'vatids.txt'
    Retrieves the company information for the VAT IDs listed in 'vatids.txt' and outputs the results to the console.

.EXAMPLE
    PS> .\Get-CompanyByVatId.ps1 -InputFile 'vatids.txt' -OutputFile 'companies.csv'
    Retrieves the company information for the VAT IDs listed in 'vatids.txt' and saves the results to 'companies.csv'.

.NOTES
    Author: Your Name
    Date: YYYY-MM-DD
    Version: 1.0
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]
    $InputFile,
    [Parameter()]
    [string]
    $OutputFile
)

function Get-CompanyByVatId {
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $VatId
    )

    # Check that VAT ID is in correct format
    if ($VatId -notmatch '^\d{7}-\d$') {
        throw 'VAT ID is not in correct format'
    }

    # Get company information from PRH API
    try {
        $CompanyData = Invoke-WebRequest -Uri "https://avoindata.prh.fi/opendata-ytj-api/v3/companies?businessId=$VatId" -ErrorAction Stop
    } catch {
        throw "Failed to retrieve data from PRH API: $_"
    }

    # Convert JSON to object
    $Content = $CompanyData.Content | ConvertFrom-Json

    # Check if company was found
    if ($Content.totalResults -ne 1) {
        throw 'Company not found or multiple companies found'
    }

    $Company = $Content.companies[0]

    # Output company information
    $CompanyName = $Company.names | Where-Object { $_.type -eq 1 -and $null -eq $_.endDate }
    # Get company visiting address (type=1). Type 2 is postal address
    $CompanyVisitingAddress = $Company.addresses | Where-Object { $_.type -eq 1 }
    # Get post office information in Finnish
    $CompanyVisitingAddressCity = $CompanyVisitingAddress.postOffices | Where-Object { $_.languageCode -eq 1 }

    # Get company postal address (type=2). Type 1 is visiting address
    $CompanyPostalAddress = $Company.addresses | Where-Object { $_.type -eq 2 }
    # Get post office information in Finnish
    $CompanyPostalAddressCity = $CompanyPostalAddress.postOffices | Where-Object { $_.languageCode -eq 1 }

    # Create new powershell object
    $CompanyObject = New-Object PSObject
    $CompanyObject | Add-Member -MemberType NoteProperty -Name 'BusinessId' -Value $Company.businessId.value
    $CompanyObject | Add-Member -MemberType NoteProperty -Name 'Name' -Value $CompanyName.name

    $CompanyObject | Add-Member -MemberType NoteProperty -Name 'Visiting Street' -Value $CompanyVisitingAddress.street + ' ' + $CompanyVisitingAddress.streetNumber 
    $CompanyObject | Add-Member -MemberType NoteProperty -Name 'Visiting PostCode' -Value $CompanyVisitingAddress.postCode
    $CompanyObject | Add-Member -MemberType NoteProperty -Name 'Visiting City' -Value $CompanyVisitingAddressCity.name

    $CompanyObject | Add-Member -MemberType NoteProperty -Name 'Postal Street' -Value $CompanyPostalAddress.street + ' ' + $CompanyPostalAddress.streetNumber
    $CompanyObject | Add-Member -MemberType NoteProperty -Name 'Postal PostCode' -Value $CompanyPostalAddress.postCode
    $CompanyObject | Add-Member -MemberType NoteProperty -Name 'Postal City' -Value $CompanyPostalAddressCity.name

    return $CompanyObject
}

# Read input file
$VatIds = Get-Content $InputFile

# Create array for company objects
$Companies = @()

# Loop through VAT IDs
foreach ($VatId in $VatIds) {
    try {
        $Company = Get-CompanyByVatId -VatId $VatId
        $Companies += $Company
    } catch {
        Write-Error "Error processing VAT ID $VatId : $_"
    }
}

# Output to console
$Companies

# Output to file
if ($OutputFile) {
    $Companies | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
}
