#####
#
# This script is to streamline creation of Mail Contacts
#
#####

# Import AD Module
Import-Module ActiveDirectory

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

# Define import file
$ContactsImportFile = Import-Csv "E:\MailContacts\Sherwin-MailContacts\MailContacts-Creation-Template.csv"

# Create Mail Contacts
ForEach ($Contact in $ContactsImportFile) {
    # Retrieve info from file
    $MailContactFirstName = $Contact.FirstName
    $MailContactLastName = $Contact.LastName
    $MailContactCompany = $Contact.Company
    $MailContactExternalEmailAddress = $Contact.ExternalEmailAddress
    
    # Construct attributes for Mail Contact object creation
    $MailContactDisplayName = "$MailContactFirstName $MailContactLastName "+"($MailContactCompany)"
    $MailContactName = "$MailContactFirstName $MailContactLastName "+"($MailContactCompany)"
    $MailContactAliasString = "$MailContactCompany"+"-"+"$MailContactFirstName"+"$MailContactLastName"
    $MailContactAlias = $MailContactAliasString.Replace(' ','')

    # Set static values
    $MailContactOrganizationalUnit = "OU=Tennant,OU=Mail Contacts,OU=Mail Contacts and Resource Mailboxes,DC=tennant,DC=tco,DC=corp"
    New-MailContact -FirstName $MailContactFirstName -LastName $MailContactLastName -DisplayName $MailContactDisplayName -Name $MailContactName -Alias $MailContactAlias -ExternalEmailAddress $MailContactExternalEmailAddress -OrganizationalUnit $MailContactOrganizationalUnit

    # Wait for Mail Contact to be created
    Write-Output "`n"
    Write-Output "Please wait 10 seconds while new Mail Contact info is replicated so additional info can be added."
    Start-Sleep -Seconds 10

    # Populate additional information
    Get-Contact -Identity $MailContactDisplayName | Set-Contact -Company $MailContactCompany
    Write-Output "`n"
    Write-Output "Any additional info attributes there were provided have been added to the Mail Contact."
}