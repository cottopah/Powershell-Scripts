# Create function to connect to Exchange 2016 on-premises
function Connect-Exchange2016OnPrem {
    # Define user variables
    $PSConnectADAESKeyFilePath = "\\dcposh101\encryptedcredentials$\psconnect-ad_aes-key.txt"
    $PSConnectADCredentialFilePath = "\\dcposh101\encryptedcredentials$\psconnect-ad_encrypted-pass.txt"

    # Create credential object
    $PSConnectADAdmin = “TENNANT\PSConnect-AD”
    $PSConnectADAESKeyFile = Get-Content $PSConnectADAESKeyFilePath
    $PSConnectADPassFile = Get-Content $PSConnectADCredentialFilePath
    $PSConnectADSecurePass = $PSConnectADPassFile | ConvertTo-SecureString –Key $PSConnectADAESKeyFile
    $PSConnectADCredObject = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $PSConnectADAdmin, $PSConnectADSecurePass

    # Get Exchange 2016 Servers
    $Exchange2016Servers = Get-ADComputer -Filter * | ? {($_.Name -like "D*EX16MBX*") -OR ($_.Name -like "D*EX16CAS*")} | Select -ExpandProperty DNSHostName

    # Get any one Exchange 2016 Server
    $Exchange2016Server = $Exchange2016Servers | Get-Random

    # Set connection session variables
    $PSConnectADSessionOption = New-PSSessionOption -SkipCNCheck -SkipCACheck 

    # Connect to Exchange remote PowerShell
    $OnPremExchangePSSession = New-PSSession -ConnectionUri "http://$Exchange2016Server/PowerShell" -ConfigurationName 'Microsoft.Exchange' -Credential $PSConnectADCredObject -SessionOption $PSConnectADSessionOption -Authentication Kerberos

    # Import PSSession to gain access to Exchange cmdlets
    Import-PSSession -Session $OnPremExchangePSSession -DisableNameChecking -AllowClobber | Out-Null
}

# Connect to Exchange
Connect-Exchange2016OnPrem

# Import AD module
Import-Module ActiveDirectory

# Define variables
$Date = (Get-Date).ToString('ddMMMyyyy')
$AccountsOU = "OU=Accounts,DC=tennant,DC=tco,DC=corp"

# Get all Users in Accounts OU
$AccountsOUUsers = Get-ADUser -Filter * -Properties Mail -SearchBase $AccountsOU | ? {$_.Mail -ne $null} | Select Name,@{N='UserID';E={$_.SamAccountName}},Mail | Sort Name

# Get all Users in Accounts OU that have a UserMailbox or RemoteUserMailbox
$AllMailboxUsers = ForEach ($User in $AccountsOUUsers) {Get-Recipient -Identity $User.UserID -ErrorAction SilentlyContinue | ? {($_.RecipientTypeDetails -like "UserMailbox") -or ($_.RecipientTypeDetails -like "RemoteUserMailbox")} | Select Name,@{N='UserID';E={$_.SamAccountName}},PrimarySmtpAddress,RecipientTypeDetails}

# Get count
$AllMailboxUsersCount = ($AllMailboxUsers).count
Write-Output "All Mailbox Users count: $AllMailboxUsersCount"

# Create additional functions for other details for report
function GetADWhenCreated ($UserID) {
    $usercreated = (Get-ADUser -Identity $UserID -Properties WhenCreated).WhenCreated
    $usercreated.ToString('MM/dd/yyyy')
}

# Create additional functions for other details for report
function GetADAccountEnabled ($UserID) {
    $userenabled = (Get-ADUser -Identity $UserID -Properties Enabled).Enabled
    $userenabled
}

function GetADAccountExpires ($UserID) {
    $accountexpire = (Get-ADUser -Identity $UserID -Properties AccountExpires).AccountExpires
    If ($accountexpire -eq "9223372036854775807") {Write-Output "Doesn't Expire"}
    ElseIf ($accountexpire -eq "0") {Write-Output "Doesn't Expire"}
    Else {
        $expiredate = [datetime]::FromFileTime($accountexpire)
        $expiredate.ToString("MM/dd/yyyy")}
}

function GetADPwdLastSet ($UserID) {
    $pwdset = (Get-ADUser -Identity $UserID -Properties PwdLastSet).PwdLastSet
    $pwdsetdate = [datetime]::FromFileTime($pwdset)
    $pwdlsetdatestring = $pwdsetdate.ToString("MM/dd/yyyy")
    If ($pwdlsetdatestring -eq "12/31/1600") {Write-Output "Never Changed By User"}
    Else {Write-Output $pwdlsetdatestring}
}

function GetADLastLogonTimestamp ($UserID) {
    $lastlogon = (Get-ADUser -Identity $UserID -Properties LastLogonTimestamp).LastLogonTimestamp
    $lastlogondate = [datetime]::FromFileTime($lastlogon)
    $lastlogondatestring = $lastlogondate.ToString("MM/dd/yyyy")
    If ($lastlogondatestring -eq "12/31/1600") {Write-Output "Never Logged In"}
    Else {Write-Output $lastlogondatestring}
}

# Perform lookup
$Result=@()
$i = 1
$AllMailboxUsers | ForEach-Object {
    $User = $_
    $Result += New-Object PSObject -Property ([ordered]@{
    'Name' = $User.Name
    'UserID' = $User.UserID
    'Mail' = $User.PrimarySmtpAddress
    'Type' = $User.RecipientTypeDetails
    'Enabled' = GetADAccountEnabled $User.UserID
    'WhenCreated' = GetADWhenCreated $User.UserID
    'AccountExpires' = GetADAccountExpires $user.UserID
    'PwdLastSet' = GetADPwdLastSet $User.UserID
    'LastLogonTimestamp' = GetADLastLogonTimestamp $User.UserID
    })
$i++
}

# View results
#$Result

# Export results
$Result | Export-Csv "c:\temp\MailboxAudit_User-to-Shared-Candidates_$Date.csv" -NoTypeInformation -Encoding UTF8
Write-Output "File is saved to c:\temp as MailboxAudit_User-to-Shared-Candidates_$Date.csv"