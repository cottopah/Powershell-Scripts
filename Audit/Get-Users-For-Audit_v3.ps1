# Get date for file timestamp purposes
$Date = (Get-Date).ToString('MMM-dd-yyyy_hhmmtt')

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

function GetADUserLastLogonAudit ($UserID) {
    # Get all DCs
    $AllDCs = Get-ADDomainController -Filter * | Select -ExpandProperty Name | Sort Name

    # Pull value from each DC
    $LoginLookupAllDCs = ForEach ($DC in $AllDCs) { 
        # Get user login details
        $UserLoginDetails = Get-ADUser $UserID -Properties LastLogon,LastLogonDate -Server $DC | `
        Select Name,@{N='UserID';E={$_.SamAccountname}},@{N="LastLogon";Expression={[datetime]::FromFileTime($_.'LastLogon')}}

        # Show results
        $Result=@()
        $i = 1
        $UserLoginDetails | ForEach-Object {
            $Value = $_
            $Result += New-Object PSObject -Property ([ordered]@{
            'Name' = $Value.Name
            'UserID' = $Value.UserID
            'LastLogon' = $Value.LastLogon.ToString('MM/dd/yyyy')
            'AuthenticatingDC' = $DC
            })
        $i++
        }
        $Result | ? {$_.LastLogon -notlike "*1600*"}
    }

    # Find newest LastLogon
    $LoginLookupAllDCs | Sort-Object LastLogon -Descending | Select-Object -First 1
}

# Call out start time
$StartDate = (Get-Date).ToString('MMM-dd-yyyy_hhmmtt')
Write-Output "Script started on $startdate."

# Filter out Resource Mailbox user objects from Exchange, various test accounts used while piloting solutions in a controlled manner, and other non-human accounts used for static functions
$ActiveUsers = Get-ADUser -Filter * -Properties PwdLastSet,WhenChanged -SearchBase "OU=Accounts,DC=tennant,DC=tco,DC=corp" | 
? {
($_.CanonicalName -notlike "*Resource Mailbox*") -AND 
($_.CanonicalName -notlike "*Test*") -AND 
($_.CanonicalName -notlike "*RF-Units*") -AND
($_.CanonicalName -notlike "*Kiosk*")
}
$ActiveUsersCount = ($ActiveUsers).count
Write-Output "On $Date, Active User count is $ActiveUsersCount."
# Note: There are some user objects that haven't been moved to the most appropriate OU yet, so you'll find some entries in the results that correlate to resource accounts and not human user accounts

# Get all user objects in "DisabledAccounts" OU where all disabled accounts are temporarily moved to for a window before they are deleted from AD - also include note when password was change and when user last logged on
$DisabledUsers = Get-ADUser -Filter * -Properties PwdLastSet,WhenChanged -SearchBase "OU=DisabledAccounts,OU=Disabled,DC=tennant,DC=tco,DC=corp"
$DisabledUsersCount = ($DisabledUsers).count
Write-Output "On $date, Disabled User count is $DisabledUsersCount."
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
    'LastLogon' = GetADUserLastLogonAudit $User.SamAccountName | Select -ExpandProperty LastLogon
    'WhenChanged' = (Get-ADUser $User.SamAccountName -Properties WhenCreated | Select -ExpandProperty WhenCreated).ToString('MM/dd/yyyy')
    })
$i++
}
$AuditActiveResult | Sort Name | Export-Csv "c:\temp\audit\activeusers\Audit-ActiveUsers_$Date-ItemCount-$ActiveUsersCount.csv" -NoTypeInformation

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
    'LastLogon' = GetADUserLastLogonAudit $User.SamAccountName
    'WhenChanged' = (Get-ADUser $User.SamAccountName -Properties WhenCreated | Select -ExpandProperty WhenCreated).ToString('MM/dd/yyyy')
    })
$i++
}
$AuditDisabledResult | Sort Name | Export-Csv "c:\temp\audit\activeusers\Audit-DisabledUsers_NotDeletedYet_$Date-ItemCount-$DisabledUsersCount.csv" -NoTypeInformation

# Call out where files were saved to
Write-Output "Script is completed."
Write-Output "Results files are saved to C:\Temp\Audit\ActiveUsers\."
Write-Output "Active User file name is: Audit-ActiveUsers_$Date-ItemCount-$ActiveUsersCount.csv"
Write-Output "Disabled User file name is: Audit-DisabledUsers_NotDeletedYet_$Date-ItemCount-$DisabledUsersCount.csv"

# Call out end time
$EndDate = (Get-Date).ToString('MMM-dd-yyyy_hhmmtt')
Write-Output "Script completed on $EndDate."