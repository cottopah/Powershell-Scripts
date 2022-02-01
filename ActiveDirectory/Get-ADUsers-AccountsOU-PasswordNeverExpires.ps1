$Date = (Get-Date).ToString('ddMMMyyyy')
$OU = "OU=Accounts,DC=tennant,DC=tco,DC=corp"
Import-Module ActiveDirectory
Get-ADUser -Filter * -SearchBase $OU -Properties PasswordNeverExpires | ? {$_.PasswordNeverExpires -eq $true} | Select Name,SamAccountName,PasswordNeverExpires | Sort Name | Export-Csv "c:\temp\AccountsOU-PasswordNeverExpires_$Date.csv" -NoTypeInformation