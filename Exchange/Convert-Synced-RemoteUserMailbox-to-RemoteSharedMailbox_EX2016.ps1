##########
#
# Changing a RemoteUserMailbox to a RemoteSharedMailbox
#
# Microsoft's recommended way to accomplish this is to move the mailbox back on-premises, convert it to a Shared Mailbox, then migrate back to Exchange Online. This is because if you convert the mailbox directly in Exchange Online
# some attributes that get changed don't write back to on-premises Active Directory. Sometimes it's not always feasible (or is not ideal) to perform this double migration. Because of this, here are the steps that can manually be taken
# perform the conversion.
#
# The attributes that come into play are:
#
# msExchRemoteRecipientType
# msExchRecipientTypeDetails
#
# Relevant values for these attributes are:
# 
# msExchRemoteRecipientType:
# 1 = ProvisionMailbox
# 4 = Migrated, UserMailbox
# 100 = Migrated, SharedMailbox
#
# msExchRecipientTypeDetails:
# 2147483648 = RemoteUserMailbox
# 34359738368 = RemoteSharedMailbox
#
# Note: The attribute msExchRecipientDisplayType with a value of -2147483642 means RemoteUserMailbox and although we're working with a RemoteSharedMailbox, the two attributes that come into play are:
# msExchRecipientDisplayType = -2147483642
# msExchRecipientTypeDetails = 34359738368
# msExchRemoteRecipientType = 100
#
#####

##########
# Create functions to connect to services
##########

#####
# Exchange on-premises
#####

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
    Import-PSSession -Session $OnPremExchangePSSession -DisableNameChecking -Prefix OnPrem -AllowClobber | Out-Null
}

#####
# Exchange Online
#####

function Connect-ExchangeOnline {
    # Define user variables
    $PSConnectO365AESKeyFilePath = "\\dcposh101\encryptedcredentials$\psconnect-o365_aes-key.txt"
    $PSConnectO365CredentialFilePath = "\\dcposh101\encryptedcredentials$\psconnect-o365_encrypted-pass.txt"

    # Define O365 credentials variables
    $PSConnectO365Admin = “psconnect-o365@tennantco.onmicrosoft.com”
    $PSConnectO365AESKeyFile = Get-Content $PSConnectO365AESKeyFilePath
    $PSConnectO365PassFile = Get-Content $PSConnectO365CredentialFilePath
    $PSConnectO365SecurePass = $PSConnectO365PassFile | ConvertTo-SecureString –Key $PSConnectO365AESKeyFile
    $PSConnectO365CredObject = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $PSConnectO365Admin, $PSConnectO365SecurePass

    # Establish PSSession
    $EXOConnectPSSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $PSConnectO365CredObject -Authentication Basic -AllowRedirection

    # Import PSSession
    Import-PSSession $EXOConnectPSSession -DisableNameChecking -Prefix O365 -AllowClobber | Out-Null
}

#####
# Import AD Module
#####

function Connect-OnPremAD {
    Import-Module ActiveDirectory
}

#####
# Connect to services
#####
Connect-ExchangeOnline
Connect-Exchange2016OnPrem
Connect-OnPremAD

##########
# Get info and convert to SharedMailbox
##########

#####
# Exchange Online
#####

# Set variables
$userid = Read-Host "Enter UserID"

# Change Type to Shared
Set-O365Mailbox $userid -Type Shared

#####
# OnPremAD
#####

# Get current values
Get-ADUser $userid -Properties msExchRemoteRecipientType,msExchRecipientTypeDetails | Select msExchRemoteRecipientType,msExchRecipientTypeDetails | FL

# Set updated values
Set-ADUser $userid -Replace @{msExchRemoteRecipientType="100"}
Set-ADUser $userid -Replace @{msExchRecipientTypeDetails="34359738368"}

# Verify values after changes
Get-ADUser $userid -Properties msExchRemoteRecipientType,msExchRecipientTypeDetails | Select msExchRemoteRecipientType,msExchRecipientTypeDetails | FL

# Note: The attribute 'LicenseReconciliationNeeded' needs to be False or the mailbox will be put into the grace period and will not function in 30 days. If set to True, assign license, run DirSync, remove license, run DirSync, then verify value.
Write-Output "The attribute 'LicenseReconciliationNeeded' needs to be False or the mailbox will be put into the grace period and will not function in 30 days. If set to True, assign license, run DirSync, remove license, run DirSync, then verify value."
Write-Output "Don't forget to remove the O365 license from the user object once the conversion is complete."