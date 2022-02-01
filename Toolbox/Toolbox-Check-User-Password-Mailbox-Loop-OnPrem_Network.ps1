####################
# Connect to Exchange on-premises and check password change and Inbox Rules
####################

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
        } 'q' {
            return
        }
    }
    pause
}
until ($selection -eq 'q')