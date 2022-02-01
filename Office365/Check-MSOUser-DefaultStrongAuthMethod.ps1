# Create function to connect to MSO using encrypted password files
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

# Force TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Connect to MSO
ConnectMSO

# Set base variables
$Date = (Get-Date).ToString('ddMMMyyyy')

# Get list of all MSO Users
$MSOUsers = Get-MsolUser -All | ? {$_.IsLicensed -eq $true} | Select DisplayName,UserPrincipalName,StrongPasswordRequired,Office,Department,Title,City,State,Country | Sort DisplayName

# Create function to check for IsDefault for StrongAuthenticationMethods
function CheckStrongAuthDefault ($UPN) {
    (Get-MsolUser -UserPrincipalName $UPN).StrongAuthenticationMethods | ? {$_.IsDefault -like $true} | Select -ExpandProperty MethodType
}

# Perform lookup
$Result=@()
$i = 1
$MSOUsers | ForEach-Object {
    $User = $_
    $Result += New-Object PSObject -Property ([ordered]@{
    'Name' = $User.DisplayName
    'UPN' = $User.UserPrincipalName
    #'StrongPasswordRequired' = $User.StrongPasswordRequired
    'DefaultStrongAuthMethod' = CheckStrongAuthDefault $User.UserPrincipalName
    'Title' = $User.Title
    'Office' = $User.Office
    'Department' = $User.Department
    'City' = $User.City
    'State' = $User.State
    'Country' = $User.Country
    })
$i++
}

# Export to CSV
$Result | Export-Csv "E:\temp\Check-MSOUser-DefaultStrongAuthMethod_$Date.csv" -NoTypeInformation -Encoding UTF8
Write-Output "File is saved to E:\temp as Check-MSOUser-DefaultStrongAuthMethod_$Date.csv"