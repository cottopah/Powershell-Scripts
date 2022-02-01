#####
#
# Use this script to create a Shared Mailbox and create an AD group that grants 'Full Access' and 'Send As' rights
#
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
    Import-PSSession -Session $OnPremExchangePSSession -DisableNameChecking -AllowClobber | Out-Null
}

# Connect to Exchange
Connect-Exchange2016OnPrem
Import-Module ActiveDirectory -WarningAction SilentlyContinue
# Clear screen
cls

###
# Get PDC Emulator to run AD commands against
###
$pdce = Get-ADDomain | Select-Object -ExpandProperty PDCEmulator
# Drop off .tennant.tco.corp to reveal just the hostname
$pdcehostname = $pdce -replace ".{17}$"

###
# Define brand/company variables - THIS IS THE ONLY SECTION THAT NEEDS CHANGES WHEN CLONING THE SCRIPT
###
$CompanyBrand = "IPC"
$CompanyBrandCountry = "GE"
$CompanyCountryCode = "DE"
$CompanyName = "Gansow"
$CompanyOffice = "Unna"
$ADUserUPNSuffix = "gansow.de"
$ADUserOU = "OU=Gansow,OU=IPC,OU=Resource Mailboxes,OU=Mail Contacts and Resource Mailboxes,DC=tennant,DC=tco,DC=corp"
$SharedMBXEmailDomain = "gansow.de"
$ADGroupOU = "OU=Security,OU=Gansow,OU=IPC,OU=User Account Groups,OU=Account Groups,DC=tennant,DC=tco,DC=corp"

###
# Prompt for input
###
$ADUserName = Read-Host "Enter desired email address prefix - i.e. invoicing"
$ADUserDisplayName = Read-Host "Enter desired Display Name"

###
# Define variables for AD User, AD Group, and Mailbox
###
$ADUserFullName = $CompanyBrand+"-"+$CompanyBrandCountry+"-"+$ADUserName
$ADUserSamName = $ADUserFullName.Substring(0, [Math]::Min($ADUserFullName.Length, 20))
$ADUserUPNPrefix = "$ADUserName"
$ADUserFullUPN = $ADUserUPNPrefix+"@"+$ADUserUPNSuffix
$SharedMBXAlias = "$ADUserSamName"
$SharedMBXEmailAddress = $ADUserName+"@"+$SharedMBXEmailDomain

###
# Create AD user for Shared Mailbox
###
New-ADUser -DisplayName "$ADUserDisplayName" -SamAccountName "$ADUserSamName" -Name "$ADUserDisplayName" -UserPrincipalName $ADUserFullUPN -Path $ADUserOU -Server $pdce -Enabled $false -Company "$CompanyName" -Office "$CompanyOffice" -Country "$CompanyCountryCode" -Description "AD User for Shared Mailbox $SharedMBXEmailAddress" | Out-Null

# Sleep for 30 seconds
Start-Sleep -Seconds 30

###
# Create mailbox
###
Enable-Mailbox -Identity $ADUserSamName -Shared -Database "EX16-DB-System" | Out-Null
# Write output
Write-Output "`n"
Write-Output "A mailbox has been created for ""$ADUserDisplayName""."

# Sleep for 30 seconds
Start-Sleep -Seconds 30

###
# Set PrimarySmtpAddress for Shared Mailbox
###
Set-Mailbox -Identity $ADUserSamName -EmailAddressPolicyEnabled $false -PrimarySmtpAddress $SharedMBXEmailAddress | Out-Null
# Write output
Write-Output "`n"
Write-Output "The primary SMTP address for ""$ADUserDisplayName"" has been set to $SharedMBXEmailAddress."
# Output update
Write-Output "`n"
Write-Output "Now that the AD user object and mailbox are created, we will create the Security Group that will be used to grant access rights to the Shared Mailbox."

###
# Define variables for Security Groups that will be used to grant access rights
###
$SecGroupDisplayName = "Sec-Exch-$CompanyBrand-$CompanyName-$ADUserName"
$SecGroupName = "Sec-Exch-$CompanyBrand-$CompanyName-$ADUserName"

###
# Create Security Group to grant "Full Access" and "Send As" rights
###
New-ADGroup -DisplayName "$SecGroupDisplayName" -Name $SecGroupName -GroupCategory Security -GroupScope Universal -Path "$ADGroupOU" -Description "Access to mailbox: $SharedMBXEmailAddress" | Out-Null
# Write output
Write-Output "`n"
Write-Output "The AD group ""$SecGroupDisplayName"" has been created."

# Pause for AD group creation so it can be used to set access rights to newly created Shared Mailbox
Write-Output "`n"
Write-Output "Waiting for AD group creation to be recognized so we can run additional commands against it."
Start-Sleep -Seconds 30

###
# Mail-enable the Security Group
###
Enable-DistributionGroup -Identity $SecGroupDisplayName | Out-Null
# Write output
Write-Output "`n"
Write-Output """$SecGroupDisplayName"" has been mail-enabled. It can now be used to set access rights on the Shared Mailbox."

# Sleep for 30 seconds
Start-Sleep -Seconds 30

###
# Set access rights for Shared Mailbox
###
# FullAccess
Add-MailboxPermission -Identity "$ADUserSamName" -User "$SecGroupDisplayName" -AccessRights FullAccess -InheritanceType All | Out-Null
# SendAs
Get-Mailbox $ADUserSamName | Add-ADPermission -User "$SecGroupDisplayName" -ExtendedRights "Send As" | Out-Null
# Write output
Write-Output "`n"
Write-Output "The group ""$SecGroupDisplayName"" has been granted 'Full Access' and 'Send As' rights to the Shared Mailbox."

###
# Post creation tasks
###
Write-Output "`n"
Write-Output "Remaining steps are to add users to the security group and determine if the mailbox should be on-premises or online."
# Remove PSSession
Get-PSSession | Remove-PSSession