# Create function to connect to Exchange 2016 on-premises
function Connect-Exchange2016OnPrem {
    # Define user variables
    $PSConnectADAESKeyFilePath = "\\dcposh101\encryptedcredentials$\psconnect-ad_aes-key.txt"
    $PSConnectADCredentialFilePath = "\\dcposh101\encryptedcredentials$\psconnect-ad_encrypted-pass.txt"

    # Create credential object
    $PSConnectADAdmin = "TENNANT\PSConnect-AD"
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

# Define variables
$UserID = Read-Host "Enter UserID of user to update email address"

# Present details about UserID before changes
Write-Output "Current settings for $UserID :"
Get-RemoteMailbox -Identity $UserID | Select Name,@{N='UserID';E={$_.SamAccountName}},@{N='EmailAddress';E={$_.PrimarySmtpAddress}} | FL

# Pause to provide chance to cancel
Write-Output "Please review the details above before continuing. Be sure you selected the right user."
pause

# Toggle Email Address Policy off
Write-Output "Setting EmailAddressPolicyEnabled to False"
Set-RemoteMailbox -Identity $UserID -EmailAddressPolicyEnabled $false

# Wait 30 seconds
Write-Output "Waiting 30 seconds before setting EmailAddressPolicyEnabled to True"
Start-Sleep -Seconds 30

# Toggle Email Address Policy on
Write-Output "Setting EmailAddressPolicyEnabled to True"
Set-RemoteMailbox -Identity $UserID -EmailAddressPolicyEnabled $true

# Present details about UserID after changes
Write-Output "Current settings for $UserID :"
Get-RemoteMailbox -Identity $UserID | Select Name,@{N='UserID';E={$_.SamAccountName}},@{N='EmailAddress';E={$_.PrimarySmtpAddress}} | FL