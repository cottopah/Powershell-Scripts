# Force TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Initial variables
$Date = (Get-Date).ToString('ddMMMyyyy')
mkdir "E:\MFAUsersAudit"
$OutputDir = "E:\MFAUsersAudit"
$OutputFile = "MFAUsersAudit-$Date.csv"

# Define groups
$PasswordlessADGroup = "Sec-Azure-AuthMethod-Passwordless" 
$ForceMFAADGroup = "Sec-Azure-ConditionalAccess-ForceMFA"
$PasswordlessADGroupDN = Get-ADGroup -Identity $PasswordlessADGroup | Select -ExpandProperty DistinguishedName
$ForceMFAADGroupDN = Get-ADGroup -Identity $ForceMFAADGroup | Select -ExpandProperty DistinguishedName

# Import AD module
Import-Module ActiveDirectory

# Function to connect to MSO
function ConnectMSO {
    # Define user variables
    $AzureADConnectO365AESKeyFilePath = "C:\Automate\EncryptedCredentials\azureadconnect-o365_aes-key.txt"
    $AzureADConnectO365CredentialFilePath = "C:\Automate\EncryptedCredentials\azureadconnect-o365_encrypted-pass.txt"

    # Define O365 credentials variables
    $AzureADConnectO365PSConnectAdmin = “azureadconnect-o365@tennantco.onmicrosoft.com”
    $AzureADConnectO365PSConnectAESKeyFile = Get-Content $AzureADConnectO365AESKeyFilePath
    $AzureADConnectO365PSConnectPassFile = Get-Content $AzureADConnectO365CredentialFilePath
    $AzureADConnectO365PSConnectSecurePass = $AzureADConnectO365PSConnectPassFile | ConvertTo-SecureString –Key $AzureADConnectO365PSConnectAESKeyFile
    $AzureADConnectO365PSConnectCredObject = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AzureADConnectO365PSConnectAdmin, $AzureADConnectO365PSConnectSecurePass

    # Establish PSSession
    $AzureADConnectO365ConnectPSSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $AzureADConnectO365PSConnectCredObject -Authentication Basic -AllowRedirection

    # Import PSSession
    Import-PSSession $AzureADConnectO365ConnectPSSession

    # Connect to MSOnline Service
    Connect-MsolService -Credential $AzureADConnectO365PSConnectCredObject
}

# Connect to MSO
ConnectMSO

# Create function to check AD group membership
Function CheckADGroupMembership ($UserUPN, $GroupDN) {
    $UserDetails = Get-ADUser -Filter * | ? {$_.UserPrincipalName -like $UserUPN}
    $MemberOf = Get-ADUser -Identity $UserDetails.SamAccountName -Properties MemberOf | Select -ExpandProperty MemberOf
    If ($MemberOf -contains $GroupDN) {Write-Output "Yes"}
    Else {Write-Output "No"}
}

# Create function to check for StrongAuthenticationMethods
Function CheckMSOMFADefaultType ($UserUPN) {
    (Get-MsolUser -UserPrincipalName $UserUPN).StrongAuthenticationMethods | ? {$_.IsDefault -eq $true} | Select -ExpandProperty MethodType
}

# Create function to get PwdLastSet
function GetADPwdLastSet ($UserUPN) {
    $pwdset = Get-ADUser -Filter * -Properties PwdLastSet| ? {$_.UserPrincipalName -like $UserUPN} | Select -ExpandProperty PwdLastSet
    $pwdsetdate = [datetime]::FromFileTime($pwdset)
    $pwdlsetdatestring = $pwdsetdate.ToString("MM/dd/yyyy")
    If ($pwdlsetdatestring -eq "12/31/1600") {Write-Output "Never Changed By User"}
    Else {Write-Output $pwdlsetdatestring}
}

# Create function to get LastLogonTimestamp
function GetADLastLogonTimestamp ($UserUPN) {
    $lastlogon = Get-ADUser -Filter * -Properties LastLogonTimestamp | ? {$_.UserPrincipalName -like $UserUPN} | Select -ExpandProperty LastLogonTimestamp
    $lastlogondate = [datetime]::FromFileTime($lastlogon)
    $lastlogondatestring = $lastlogondate.ToString("MM/dd/yyyy")
    If ($lastlogondatestring -eq "12/31/1600") {Write-Output "Never Logged In"}
    Else {Write-Output $lastlogondatestring}
}

# Find MSO users where MFA enabled (StrongAuthenticationMethods value is present)
$MFAMSOUsers = Get-MsolUser -All | ? {$_.StrongAuthenticationMethods -ne $null} | Select DisplayName,UserPrincipalName
#$MFAMSOUsers = Get-MsolUser -UserPrincipalName bme9@tennantco.com | Select DisplayName,UserPrincipalName

# Perform lookup
$Result=@()
$i = 1
$MFAMSOUsers | ForEach-Object {
    $User = $_
    $Result += New-Object PSObject -Property ([ordered]@{
    'Name' = $User.DisplayName
    'UPN' = $User.UserPrincipalName
    'MFADefaultType' = CheckMSOMFADefaultType $User.UserPrincipalName
    'PasswordlessADGroup' = CheckADGroupMembership $User.UserPrincipalName $PasswordlessADGroupDN
    'ForceMFAADGroup' = CheckADGroupMembership $User.UserPrincipalName $ForceMFAADGroupDN
    'PwdLastSet' = GetADPwdLastSet $User.UserPrincipalName
    'LastLogonTimestamp' = GetADLastLogonTimestamp $User.UserPrincipalName
    })
$i++
}

# Export results to file
$Result | Sort Name | Export-Csv "$OutputDir\$OutputFile" -NoTypeInformation -Encoding UTF8
Write-Output "File is saved to $OutputDir as $OutputFile."