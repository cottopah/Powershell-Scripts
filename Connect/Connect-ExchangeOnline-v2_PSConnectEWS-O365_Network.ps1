function Connect-ExchangeOnline-v2 {
    # Define user variables
    $PSConnectO365AESKeyFilePath = "\\dcposh101\encryptedcredentials$\psconnectews-o365_aes-key.txt"
    $PSConnectO365CredentialFilePath = "\\dcposh101\encryptedcredentials$\psconnectews-o365_encrypted-pass.txt"

    # Define O365 credentials variables
    $PSConnectO365Admin = “psconnectews-o365@tennantco.com”
    $PSConnectO365AESKeyFile = Get-Content $PSConnectO365AESKeyFilePath
    $PSConnectO365PassFile = Get-Content $PSConnectO365CredentialFilePath
    $PSConnectO365SecurePass = $PSConnectO365PassFile | ConvertTo-SecureString –Key $PSConnectO365AESKeyFile
    $PSConnectO365CredObject = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $PSConnectO365Admin, $PSConnectO365SecurePass

    # Initiate connection
    Connect-ExchangeOnline -Credential $PSConnectO365CredObject -Prefix O365
}

# Connect to Exchange Online
Connect-ExchangeOnline-v2