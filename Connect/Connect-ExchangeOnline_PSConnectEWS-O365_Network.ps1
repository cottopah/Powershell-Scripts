####################
# Connect to Exchange Online via PowerShell script using encrypted password file
####################

# Define user variables
$O365AESKeyFilePath = "\\DCPOSH101\EncryptedCredentials$\psconnectews-o365_aes-key.txt"
$O365CredentialFilePath = "\\DCPOSH101\EncryptedCredentials$\psconnectews-o365_encrypted-pass.txt"

# Define O365 credentials variables
$O365PSConnectAdmin = “psconnectews-o365@tennantco.com”
$O365PSConnectAESKeyFile = Get-Content $O365AESKeyFilePath
$O365PSConnectPassFile = Get-Content $O365CredentialFilePath
$O365PSConnectSecurePass = $O365PSConnectPassFile | ConvertTo-SecureString –Key $O365PSConnectAESKeyFile
$O365PSConnectCredObject = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $O365PSConnectAdmin, $O365PSConnectSecurePass

# Establish PSSession
$O365ConnectPSSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $O365PSConnectCredObject -Authentication Basic -AllowRedirection

# Import PSSession
Import-PSSession $O365ConnectPSSession -DisableNameChecking -AllowClobber