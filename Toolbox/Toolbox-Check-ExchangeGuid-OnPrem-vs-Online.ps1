####################
# 
# Use this script to check on-premises ExchangeGuid to match online ExchangeGuid so mailbox can be moved back on-premises 
#
####################

##########
# Create credential objects
##########

#####
# Exchange 2016 on-premises
#####

# Define user variables
$PSConnectADAESKeyFilePath = "C:\Automate\EncryptedCredentials\psconnect-ad_aes-key.txt"
$PSConnectADCredentialFilePath = "C:\Automate\EncryptedCredentials\psconnect-ad_encrypted-pass.txt"

# Define PSConnect-AD credentials variables
$PSConnectADAdmin = “TENNANT\PSConnect-AD”
$PSConnectADAESKeyFile = Get-Content $PSConnectADAESKeyFilePath
$PSConnectADPassFile = Get-Content $PSConnectADCredentialFilePath
$PSConnectADSecurePass = $PSConnectADPassFile | ConvertTo-SecureString –Key $PSConnectADAESKeyFile
$PSConnectADCredObject = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $PSConnectADAdmin, $PSConnectADSecurePass

#####
# Exchange Online
#####

# Define user variables
$O365AESKeyFilePath = "C:\Automate\EncryptedCredentials\psconnect-o365_aes-key.txt"
$O365CredentialFilePath = "C:\Automate\EncryptedCredentials\psconnect-o365_encrypted-pass.txt"

# Define O365 credentials variables
$O365PSConnectAdmin = “psconnect-o365@tennantco.onmicrosoft.com”
$O365PSConnectAESKeyFile = Get-Content $O365AESKeyFilePath
$O365PSConnectPassFile = Get-Content $O365CredentialFilePath
$O365PSConnectSecurePass = $O365PSConnectPassFile | ConvertTo-SecureString –Key $O365PSConnectAESKeyFile
$O365PSConnectCredObject = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $O365PSConnectAdmin, $O365PSConnectSecurePass

#####
# AzureADConnect
#####

# Define user variables
$AzureADConnectADAESKeyFilePath = "C:\Automate\EncryptedCredentials\azureadconnect-ad_aes-key.txt"
$AzureADConnectADCredentialFilePath = "C:\Automate\EncryptedCredentials\azureadconnect-ad_encrypted-pass.txt"

# Define AzureADConnectAD credentials variables
$AzureADConnectADAdmin = “TENNANT\AzureADConnect-AD”
$AzureADConnectADAESKeyFile = Get-Content $AzureADConnectADAESKeyFilePath
$AzureADConnectADPassFile = Get-Content $AzureADConnectADCredentialFilePath
$AzureADConnectADSecurePass = $AzureADConnectADPassFile | ConvertTo-SecureString –Key $AzureADConnectADAESKeyFile
$AzureADConnectADCredObject = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AzureADConnectADAdmin, $AzureADConnectADSecurePass

#####
# Quick reference:
#
# $PSConnectADCredObject = Exchange on-premises credential object
# $O365PSConnectCredObject = Exchange Online credential object
# $AzureADConnectADCredObject = AzureADConnect credential object
#
#####

##########
# Create functions to connect to Exchange via PowerShell
##########

#####
# Exchange on-premises
#####

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
# AzureADConnect
#####

function Run-AzureADConnect {    
    # Connect to DCMSSSO100 to run Azure AD Connect sync
    $AzureADConnectADPSSession = New-PSSession –ComputerName DCMSSSO100 -Credential $AzureADConnectADCredObject

    # Import ADSync module
    Invoke-Command -Session $AzureADConnectADPSSession -ScriptBlock {Import-Module "C:\Program Files\Microsoft Azure AD Sync\Bin\ADSync\ADSync.psd1"}

    # Run delta sync
    Invoke-Command -Session $AzureADConnectADPSSession -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}

    # End remote PSSession
    Remove-PSSession $AzureADConnectADPSSession
}

##########
# Run connection functions to establish Exchange PowerShell sessions
##########

Connect-Exchange2016OnPrem
Connect-ExchangeOnline

##########
# Create function to perform query and changes
##########

function Show-Menu
{
    param (
        [string]$Title = 'IT Toolbox - Check ExchangeGuid on-premises vs. online'
    )
    Clear-Host
    Write-Host "==========$Title=========="

    Write-Host "1: Press '1' to define user you want to query"
    Write-Host "2: Press '2' to check on-premises ExchangeGuid"
    Write-Host "3: Press '3' to check online ExchangeGuid"
    Write-Host "4: Press '4' to view both ExchangeGuid values"
    Write-Host "5: Press '5' to set on-premises ExchangeGuid to match online ExchangeGuid"
    Write-Host "6: Press '6' to run Azure AD Connect to sync ExchangeGuid values (if changes were made)"
    Write-Host "Q: Press 'Q' to quit"
}

do
{
    Show-Menu
    $selection = Read-Host "Please make a selection"
    switch ($selection)
    {
        '1' {
            # Define user alias
            $usermbx = Read-Host "Enter alias of user"
        } '2' {
            #####
            # Check Exchange on-premises RemoteMailbox ExchangeGuid
            #####
            # Check to see if ExchangeGuid is all zeros
            $ExchGuidOnPrem = Get-OnPremRemoteMailbox $usermbx | Select -ExpandProperty ExchangeGuid 
            Write-Host "ExchangeGuid in on-premises Exchange is: $ExchGuidOnPrem"
        } '3' {
            #####
            # Check Exchange Online Mailbox ExchangeGuid
            #####
            # Get ExchangeGuid from UserMailbox defined earlier
            $ExchGuidOnline = Get-O365Mailbox $usermbx | Select -ExpandProperty ExchangeGuid
            Write-Host "ExchangeGuid in Exchange Online is: $ExchGuidOnline"
        }  '4' {
            #####
            # View both Exchange Guids
            #####
            Write-Host "ExchangeGuid in on-premises Exchange is: $ExchGuidOnPrem"
            Write-Host "ExchangeGuid in Exchange Online is: $ExchGuidOnline"
            Write-Host "If ExchangeGuid in on-premises Exchange is all zeros, the mailbox needs to get ExchangeGuid from O365 stamped to it so it can be migrated back on-premises."
        } '5' {
            #####
            # Set ExchangeGuid to match to allow for migration back to on-premises
            #####
            # Set ExchangeGuid to match O365 ExchangeGuid
            Set-OnPremRemoteMailbox $usermbx -ExchangeGuid $ExchGuidOnline
        } '6' {
            #####
            # Run Azure AD Connect to push attribute changes to Office 365
            #####
            Run-AzureADConnect
        } 'q' {
            return
        }
    }
    pause
}
until ($selection -eq 'q')