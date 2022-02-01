####################
# Connect to Exchange Online via PowerShell script using encrypted password file
####################

# Define user variables
$O365AESKeyFilePath = "\\dcposh100vm\encryptedcredentials$\psconnect-o365_aes-key.txt"
$O365CredentialFilePath = "\\dcposh100vm\encryptedcredentials$\psconnect-o365_encrypted-pass.txt"

# Define O365 credentials variables
$O365PSConnectAdmin = “PSConnect-O365@tennantco.onmicrosoft.com”
$O365PSConnectAESKeyFile = Get-Content $O365AESKeyFilePath
$O365PSConnectPassFile = Get-Content $O365CredentialFilePath
$O365PSConnectSecurePass = $O365PSConnectPassFile | ConvertTo-SecureString –Key $O365PSConnectAESKeyFile
$O365PSConnectCredObject = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $O365PSConnectAdmin, $O365PSConnectSecurePass

# Establish PSSession
$O365ConnectPSSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $O365PSConnectCredObject -Authentication Basic -AllowRedirection

# Import PSSession
Import-PSSession $O365ConnectPSSession

# Import Active Directory module
Import-Module ActiveDirectory

function Show-Menu
{
    param (
        [string]$Title = 'IT Toolbox - Check user password and mailbox'
    )
    Clear-Host
    Write-Host "==========$Title=========="

    Write-Host "1: Press '1' to define user you want to query"
    Write-Host "2: Press '2' to show user info"
    Write-Host "3: Press '3' to check when the user last changed their password"
    Write-Host "4: Press '4' to check if the user mailbox has forwarding or Inbox Rules"
    Write-Host "5: Press '5' to check if the user is on the Blocked Sender List"
    Write-Host "Q: Press 'Q' to quit"
}

do
{
    Show-Menu
    $selection = Read-Host "Please make a selection"
    switch ($selection)
    {
        '1' {
            # Define user and collect info
            $user = Read-Host "Enter alias of user that you want to check"
            Get-ADUser $user -Properties * | Select SamAccountName -ExpandProperty SamAccountName
        } '2' {
            # Get additional user info
            Get-ADUser $user -Properties * | Select Name,SamAccountName,UserPrincipalName,EmailAddress | FL
        } '3' {
            # Get user name and password last set
            Get-ADUser $user -Properties * | Select Name,PasswordLastSet | FL
        } '4' {
            # Get user UPN
            $userupn = Get-ADUser $user -Properties * | Select UserPrincipalName -ExpandProperty UserPrincipalName
            # Check to see if forwarding is set in the mailbox
            Get-Mailbox $userupn | Select Name,ForwardingAddress,ForwardingSmtpAddress | FL
            # Check to see if Inbox Rules exist in the mailbox
            Get-InboxRule -Mailbox $userupn | Select MailboxOwnerId,Name,Enabled,Description | FL
        } '5' {
            # Get user email address
            $useremail = Get-ADUser $user -Properties * | Select EmailAddress -ExpandProperty EmailAddress
            # Check Blocked Sender list
            Get-BlockedSenderAddress -SenderAddress $useremail | Out-GridView
        } 'q' {
            return
        }
    }
    pause
}
until ($selection -eq 'q')