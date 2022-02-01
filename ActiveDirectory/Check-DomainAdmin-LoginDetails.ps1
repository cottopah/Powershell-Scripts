$DomainAdmins = Get-ADGroupMember -Identity "Domain Admins" | Select Name,SamAccountName | Sort Name

$LoginCheck = ForEach ($User in $DomainAdmins) {Get-ADUser -Identity $User.SamAccountName -Properties LastLogonDate,PasswordLastSet | Select Name,SamAccountName,LastLogonDate,PasswordLastSet}

$LoginCheck | Export-Csv "c:\temp\DomainAdminsLoginCheck.csv" -NoTypeInformation