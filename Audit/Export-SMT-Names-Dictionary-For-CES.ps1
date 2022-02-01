<#
# Summary
  
  This script is used to gather info about the Senior Management Team for purposes of populating the Dictionary in Cisco Email Security for the Forged Email Detection function

    :: Process to change SMT_Names dictionary file in CES ::

    1) Generate new file using PowerShell
    2) Import new dictionary file (make sure to check "Match whole words" option) [use SMT_DATE for name]
    3) Adjust "Forged_SMT_Names" Incoming Content Filter to reference newly imported dictionary file
    4) Delete old dictionary file that is no longer being used by Content Filter
    5) Commit changes

    PowerShell scripts involved:
    Get-SMT-List-For-CES.ps1
    Export-SMT-Names-Dictionary-For-CES.ps1

#>

###
# Set base variables
###
$Today = (Get-Date).ToString('ddMMMyyyy')
$Directory = "C:\Temp\SMT_Names"
$File = "smt-ces-audit"
$ImportFile3Tier = "$Directory\$File-$Today.csv"

###
# Create function
###
function FormatForCESDictionary {
    process {
     ForEach-Object {$_ + ", 1"}
     }
}

###
# Define list of users and export to file - 3 tier
###
$UsersTier3 = Import-Csv $ImportFile3Tier | % {$_.FirstLast}

# Format list for CES Dictionary file
$CESDictionaryTier3 = $UsersTier3 | FormatForCESDictionary

# Export to file (no extension)
$CESDictionaryTier3 | Out-File "$Directory\SMT_Names_CES-$today" -Encoding utf8