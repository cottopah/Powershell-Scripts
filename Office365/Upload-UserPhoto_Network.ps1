##########
#
# Per this MSFT KB: https://support.microsoft.com/en-us/help/3062745/user-photos-aren-t-synced-from-the-on-premises-environment-to-exchange
#
# "The thumbnailPhoto attribute is synced only one time between Azure AD and Exchange Online. Any later changes to the attribute from the on-premises environment are not synced to the Exchange Online mailbox."
#
# To make this process as easy as possible, save the updated photo as userid.jpg (ex. bme9.jpg) and make sure the file is 10K or less.
#
##########

####################
# Connect to Exchange Online via PowerShell using RPS as proxy method to use Set-UserPhoto
####################

# Define user variables
$O365AESKeyFilePath = "\\dcposh101\encryptedcredentials$\psconnect-o365_aes-key.txt"
$O365CredentialFilePath = "\\dcposh101\encryptedcredentials$\psconnect-o365_encrypted-pass.txt"

# Define O365 credentials variables
$O365PSConnectAdmin = “psconnect-o365@tennantco.onmicrosoft.com”
$O365PSConnectAESKeyFile = Get-Content $O365AESKeyFilePath
$O365PSConnectPassFile = Get-Content $O365CredentialFilePath
$O365PSConnectSecurePass = $O365PSConnectPassFile | ConvertTo-SecureString –Key $O365PSConnectAESKeyFile
$O365PSConnectCredObject = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $O365PSConnectAdmin, $O365PSConnectSecurePass

# Establish PSSession
$O365ConnectPSSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/?proxymethod=rps -Credential $O365PSConnectCredObject -Authentication Basic -AllowRedirection

# Import PSSession
Import-PSSession $O365ConnectPSSession -DisableNameChecking | Out-Null

##########
# Perform removal of existing picture and upload a new one
##########

Write-Host "Make sure the picture is saved to \\dcwebit100vm\c$\website\photos and is named userid.jpg (ex. bme9.jpg) and that the file is 10K or less."
Write-Host "`n"
# Define variables
Set-Location D:\Scripts\Office365
$photouserid = Read-Host "Enter UserID of user"
$photosourcefolder = "\\dcwebit100vm\c$\website\photos"
$photouserfile = Read-Host "Enter name of photo file (example: bme9.jpg)"
$photofullpath = join-path -path $photosourcefolder -ChildPath $photouserfile

# Remove existing photo
Write-Host "`n"
Write-Host "Removing existing photo"
Remove-UserPhoto -Identity $photouserid -Confirm:$false

# Upload new photo
Write-Host "`n"
Write-Host "Uploading new photo"
Set-UserPhoto -Identity $photouserid -PictureData ([System.IO.File]::ReadAllBytes("$photofullpath")) -Confirm:$false

# Indicate completion
Write-Host "`n"
Write-Host "Upload complete."