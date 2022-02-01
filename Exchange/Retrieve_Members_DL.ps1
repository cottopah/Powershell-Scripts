Import-Module ActiveDirectory

Get-ADGroupMember -Identity "DL Global Leadership Team" | Select Name, SamAccountName | Sort Name | Export-csv -Path "E:\SMTEmployees.csv" -NoTypeInformation


