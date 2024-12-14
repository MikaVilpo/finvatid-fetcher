# Get-CompanyByVatId.ps1

## Overview

`Get-CompanyByVatId.ps1` is a PowerShell script that retrieves company information based on VAT ID from the PRH API. It can process multiple VAT IDs from an input file and optionally output the results to a CSV file.

## Usage

### Parameters

- `InputFile` (Mandatory): The path to the input file containing VAT IDs, one per line.
- `OutputFile` (Optional): The path to the output CSV file where the company information will be saved.

### Examples

#### Example 1: Output to Console

```powershell
PS> .\Get-CompanyByVatId.ps1 -InputFile 'vatids.txt'
```

#### Example 2: Output to file

```powershell
PS> .\Get-CompanyByVatId.ps1 -InputFile 'vatids.txt' -OutputFile 'companies.csv'
```
