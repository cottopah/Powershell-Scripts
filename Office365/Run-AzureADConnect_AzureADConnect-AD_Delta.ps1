####################
# Run Azure AD Connect to push attribute changes to Office 365
####################

# Define user variables
$AzureADConnectADAESKeyFilePath = "C:\Automate\EncryptedCredentials\azureadconnect-ad_aes-key.txt"
$AzureADConnectADCredentialFilePath = "C:\Automate\EncryptedCredentials\azureadconnect-ad_encrypted-pass.txt"

# Define AzureADConnectAD credentials variables
$AzureADConnectADAdmin = “TENNANT\AzureADConnect-AD”
$AzureADConnectADAESKeyFile = Get-Content $AzureADConnectADAESKeyFilePath
$AzureADConnectADPassFile = Get-Content $AzureADConnectADCredentialFilePath
$AzureADConnectADSecurePass = $AzureADConnectADPassFile | ConvertTo-SecureString –Key $AzureADConnectADAESKeyFile
$AzureADConnectADCredObject = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AzureADConnectADAdmin, $AzureADConnectADSecurePass

# Connect to DCMSSSO100 to run Azure AD Connect sync
$AzureADConnectADPSSession = New-PSSession –ComputerName DCMSSSO100 -Credential $AzureADConnectADCredObject

# Import ADSync module
Invoke-Command -Session $AzureADConnectADPSSession -ScriptBlock {Import-Module "C:\Program Files\Microsoft Azure AD Sync\Bin\ADSync\ADSync.psd1"}

# Run delta sync
Invoke-Command -Session $AzureADConnectADPSSession -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}

# End remote PSSession
Remove-PSSession $AzureADConnectADPSSession