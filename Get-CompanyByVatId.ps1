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
    PS> $CompanyInformation = .\Get-CompanyByVatId.ps1 -InputFile 'vatids.txt'
    Retrieves the company information for the VAT IDs listed in 'vatids.txt' and outputs the results to the console and stores it in $CompanyInformation.

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

    # Get company information from PRH API with retry logic for 429 status code
    $MaxRetries = 5
    $RetryDelay = 5
    $RetryCount = 0
    $Success = $false

    while (-not $Success -and $RetryCount -lt $MaxRetries) {
        try {
            $CompanyData = Invoke-WebRequest `
                -Uri "https://avoindata.prh.fi/opendata-ytj-api/v3/companies?businessId=$VatId" `
                -ErrorAction Stop 
            $Success = $true
        }
        catch {
            if ($_.Exception.Response.StatusCode -eq 429) {
                Write-Host "Received 429 Too Many Requests. Retrying in $RetryDelay seconds..."
                Start-Sleep -Seconds $RetryDelay
                $RetryCount++
            }
            else {
                throw "Failed to retrieve data from PRH API: $_"
            }
        }
    }

    if (-not $Success) {
        throw "Failed to retrieve data from PRH API after $MaxRetries attempts."
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

    
    $CompanyObject | Add-Member -MemberType NoteProperty -Name 'Visiting CO' -Value $CompanyVisitingAddress.co
    $CompanyObject | Add-Member -MemberType NoteProperty -Name 'Visiting Street' -Value $($CompanyVisitingAddress.street + ' ' + $CompanyVisitingAddress.buildingNumber + ' ' + $CompanyVisitingAddress.entrance + ' ' + $CompanyVisitingAddress.apartmentNumber)
    $CompanyObject | Add-Member -MemberType NoteProperty -Name 'Visiting PostCode' -Value $CompanyVisitingAddress.postCode
    $CompanyObject | Add-Member -MemberType NoteProperty -Name 'Visiting City' -Value $CompanyVisitingAddressCity.city
    
    $CompanyObject | Add-Member -MemberType NoteProperty -Name 'Postal CO' -Value $CompanyPostalAddress.co
    $CompanyObject | Add-Member -MemberType NoteProperty -Name 'Postal Postbox' -Value $(if ($null -ne $CompanyPostalAddress.postOfficeBox) { 'PL ' + $CompanyPostalAddress.postOfficeBox } else { $null })
    $CompanyObject | Add-Member -MemberType NoteProperty -Name 'Postal Street' -Value $($CompanyPostalAddress.street + ' ' + $CompanyPostalAddress.buildingNumber + ' ' + $CompanyPostalAddress.entrance + ' ' + $CompanyPostalAddress.apartmentNumber)
    $CompanyObject | Add-Member -MemberType NoteProperty -Name 'Postal PostCode' -Value $CompanyPostalAddress.postCode
    $CompanyObject | Add-Member -MemberType NoteProperty -Name 'Postal City' -Value $CompanyPostalAddressCity.city

    return $CompanyObject
}

# Read input file
$VatIds = Get-Content $InputFile

# Validate input file
if (-not $VatIds) {
    Write-Error 'Input file is empty or does not exist'
    exit
}

# Remove duplicates
$VatIds = $VatIds | Select-Object -Unique

# Trim whitespace
$VatIds = $VatIds | ForEach-Object { $_.Trim() }

# Count VAT IDs
$VatIdCount = $VatIds.Count
Write-Host "Processing $VatIdCount VAT IDs..."

# Create progress bar
$Progress = 0
$ProgressStep = 100 / $VatIdCount

# Create array for company objects
$Companies = @()

# Loop through VAT IDs
foreach ($VatId in $VatIds) {
    try {
        $Progress += $ProgressStep
        Write-Progress -Activity 'Processing VAT IDs' -Status "Processing VAT ID $VatId. $([math]::Round($Progress, 2))% complete." -PercentComplete $Progress
        $Company = Get-CompanyByVatId -VatId $VatId
        $Companies += $Company
    }
    catch {
        Write-Error "Error processing VAT ID $VatId : $_"
    }
}

# Output to console
$Companies

# Output to file
if ($OutputFile) {
    $Companies | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8 -Delimiter ';'
}
