# Set date variable
$Date = (Get-Date).ToString('ddMMMyyyy')

# Define OU
$OU = "OU=Accounts,DC=tennant,DC=tco,DC=corp"

# Import AD module
Import-Module ActiveDirectory

# Perform lookup
Get-ADUser -Filter * -SearchBase $OU -Properties PasswordNeverExpires | ? {$_.PasswordNeverExpires -eq $true} | Select Name,SamAccountName,PasswordNeverExpires | Sort Name | Export-Csv "c:\temp\AccountsOU-PasswordNeverExpires_$Date.csv" -NoTypeInformation

# Write output
Write-Output "File is saved to c:\temp\AccountsOU-PasswordNeverExpires_$Date.csv"