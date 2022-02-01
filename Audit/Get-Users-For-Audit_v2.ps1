# Get date for file timestamp purposes
$date = (Get-Date).ToString('MMM-dd-yyyy_hhmmtt')

# Create directory for output files
mkdir -Force C:\temp\Audit\ActiveUsers > $null

# Import AD Module
Import-Module ActiveDirectory

# Creat functions
function GetADPwdLastSet ($User) {
    $pwdset = (Get-ADUser -Identity $User.SamAccountName -Properties PwdLastSet).PwdLastSet
    $pwdsetdate = [datetime]::FromFileTime($pwdset)
    $pwdlsetdatestring = $pwdsetdate.ToString("MM/dd/yyyy")
    If ($pwdlsetdatestring -eq "12/31/1600") {Write-Output "Never Changed By User"}
    Else {Write-Output $pwdlsetdatestring}
}

function GetADLastLogonTimestamp ($User) {
    $lastlogon = (Get-ADUser -Identity $User.SamAccountName -Properties LastLogonTimestamp).LastLogonTimestamp
    $lastlogondate = [datetime]::FromFileTime($lastlogon)
    $lastlogondatestring = $lastlogondate.ToString("MM/dd/yyyy")
    If ($lastlogondatestring -eq "12/31/1600") {Write-Output "Never Logged In"}
    Else {Write-Output $lastlogondatestring}
}

# Call out start time
$startdate = (Get-Date).ToString('MMM-dd-yyyy_hhmmtt')
Write-Output "Script started on $startdate."

# Filter out Resource Mailbox user objects from Exchange, various test accounts used while piloting solutions in a controlled manner, and other non-human accounts used for static functions
$activeusers = Get-ADUser -Filter * -Properties PasswordLastSet,LastLogonTimestamp,WhenChanged -SearchBase "OU=Accounts,DC=tennant,DC=tco,DC=corp" | 
? {
($_.CanonicalName -notlike "*Resource Mailbox*") -AND 
($_.CanonicalName -notlike "*Test*") -AND 
($_.CanonicalName -notlike "*RF-Units*") -AND
($_.CanonicalName -notlike "*Kiosk*")
}
$activeuserscount = ($activeusers).count
Write-Output "On $date, Active User count is $activeuserscount."
# Note: There are some user objects that haven't been moved to the most appropriate OU yet, so you'll find some entries in the results that correlate to resource accounts and not human user accounts

# Get all user objects in "DisabledAccounts" OU where all disabled accounts are temporarily moved to for a window before they are deleted from AD - also include note when password was change and when user last logged on
$disabledusers = Get-ADUser -Filter * -Properties PasswordLastSet,LastLogonTimestamp,WhenChanged -SearchBase "OU=DisabledAccounts,OU=Disabled,DC=tennant,DC=tco,DC=corp"
$disableduserscount = ($disabledusers).count
Write-Output "On $date, Disabled User count is $disableduserscount."
Write-Output "Disabled Users means user accounts in DisabledAccounts OU."
# Note: When user termination requests are submitted, the AD user objects gets the password reset and is moved to the "DisabledAccounts" OU where it remains for 30 days in case a recovery is needed (i.e. Legal Hold, etc.)

# Results for Active Users
$AuditActiveResult=@()
$i = 1
$ActiveUsers | ForEach-Object {
    $User = $_
    $AuditActiveResult += New-Object PSObject -Property ([ordered]@{
    'Name' = $User.Name
    'SamAccountName' = $User.SamAccountName
    'UserPrincipalName' = $User.UserPrincipalName
    'PwdLastSet' = GetADPwdLastSet $User
    'LastLogonTimestamp' = GetADLastLogonTimestamp $User
    'WhenChanged' = (Get-ADUser $User.SamAccountName -Properties WhenCreated | Select -ExpandProperty WhenCreated).ToString('MM/dd/yyyy')
    })
$i++
}
$AuditActiveResult | Sort Name | Export-Csv "c:\temp\audit\activeusers\Audit-ActiveUsers_$date-ItemCount-$activeuserscount.csv" -NoTypeInformation

# Results for Disabled Users
$AuditDisabledResult=@()
$i = 1
$DisabledUsers | ForEach-Object {
    $User = $_
    $AuditDisabledResult += New-Object PSObject -Property ([ordered]@{
    'Name' = $User.Name
    'SamAccountName' = $User.SamAccountName
    'UserPrincipalName' = $User.UserPrincipalName
    'PwdLastSet' = GetADPwdLastSet $User
    'LastLogonTimestamp' = GetADLastLogonTimestamp $User
    'WhenChanged' = (Get-ADUser $User.SamAccountName -Properties WhenCreated | Select -ExpandProperty WhenCreated).ToString('MM/dd/yyyy')
    })
$i++
}
$AuditDisabledResult | Sort Name | Export-Csv "c:\temp\audit\activeusers\Audit-DisabledUsers_NotDeletedYet_$date-ItemCount-$disableduserscount.csv" -NoTypeInformation

# Call out where files were saved to
Write-Output "Script is completed."
Write-Output "Results files are saved to C:\Temp\Audit\ActiveUsers\."
Write-Output "Active User file name is: Audit-ActiveUsers_$date-ItemCount-$activeuserscount.csv"
Write-Output "Disabled User file name is: Audit-DisabledUsers_NotDeletedYet_$date-ItemCount-$disableduserscount.csv"

# Call out end time
$enddate = (Get-Date).ToString('MMM-dd-yyyy_hhmmtt')
Write-Output "Script completed on $enddate."