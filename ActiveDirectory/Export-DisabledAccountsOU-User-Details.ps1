# Set date variable
$Date = (Get-Date).ToString('ddMMMyyyy')

# Set file variables
$OutputDir = "c:\temp\"
$OutputFile = "DisabledAccountsOU-Export-$Date.csv"

# Define DisabledAccounts OU
$DisabledAccountsOU = "OU=DisabledAccounts,OU=Disabled,DC=tennant,DC=tco,DC=corp"

# Get objects from DisabledAccounts OU
$DisabledAccountsOUUsers = Get-ADUser -Filter * -SearchBase $DisabledAccountsOU -Properties Mail,Description | Select Name,@{N='UserID';E={$_.SamAccountName}},Mail,Description | Sort Name

# Get count and display result
$Count = $DisabledAccountsOUUsers.count
Write-Output "Accounts in DisabledAccounts OU: $Count"

# Export to CSV
$DisabledAccountsOUUsers | Export-Csv $OutputDir$OutputFile -NoTypeInformation
Write-Output "File $OutputFile is saved to $OutputDir."