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
# Define brand variables
$CompanyName = "Tennant"
$ADUserUPNSuffix = "tennantco.com"
$SharedMBXEmailDomain = "tennantco.com"

# Define variables for location of User object
$ADUserOUAPAC = "OU=APAC,OU=Tennant,OU=Resource Mailboxes,OU=Mail Contacts and Resource Mailboxes,DC=tennant,DC=tco,DC=corp"
$ADUserOUEMEA = "OU=EMEA,OU=Tennant,OU=Resource Mailboxes,OU=Mail Contacts and Resource Mailboxes,DC=tennant,DC=tco,DC=corp"
$ADUserOUNCSA = "OU=NCSA,OU=Tennant,OU=Resource Mailboxes,OU=Mail Contacts and Resource Mailboxes,DC=tennant,DC=tco,DC=corp"

# Define variables for location of Group object
$ADGroupOUAPAC = "OU=APAC,OU=Security,OU=Tennant,OU=User Account Groups,OU=Account Groups,DC=tennant,DC=tco,DC=corp"
$ADGroupOUEMEA = "OU=EMEA,OU=Security,OU=Tennant,OU=User Account Groups,OU=Account Groups,DC=tennant,DC=tco,DC=corp"
$ADGroupOUNCSA = "OU=NCSA,OU=Security,OU=Tennant,OU=User Account Groups,OU=Account Groups,DC=tennant,DC=tco,DC=corp"

###
# Prompt for input
###
$ADUserName = Read-Host "Enter desired email address prefix - i.e. invoicing"
$ADUserDisplayName = Read-Host "Enter desired Display Name - i.e. Tennant Invoicing"
$CompanyBrandCountry = Read-Host "Enter Country - i.e. US, Netherlands, etc."
$CompanyCountryCode = Read-Host "Enter 2 letter Country identifier - i.e. US, NL, DE, etc."
$CompanyOffice = Read-Host "Enter name of Office - i.e. Corporate Woods, Uden, etc."
$ADUserRegion = Read-Host "Enter region for Shared Mailbox - APAC, EMEA, or NCSA"

###
# Define variables for AD User, AD Group, and Mailbox
###
$ADUserFullName = $CompanyCountryCode+"-"+$ADUserName
$ADUserSamName = $ADUserFullName.Substring(0, [Math]::Min($ADUserFullName.Length, 20))
$ADUserUPNPrefix = "$ADUserName"
$ADUserFullUPN = $ADUserUPNPrefix+"@"+$ADUserUPNSuffix
$SharedMBXAlias = "$ADUserSamName"
$SharedMBXEmailAddress = $ADUserName+"@"+$SharedMBXEmailDomain

###
# Create user object and associated mailbox
###
# Create AD user for Shared Mailbox

# APAC
If ($ADUserRegion -eq "APAC") {
    New-ADUser -DisplayName "$ADUserDisplayName" -SamAccountName "$ADUserSamName" -Name "$ADUserDisplayName" -UserPrincipalName $ADUserFullUPN -Path $ADUserOUAPAC -Server $pdce -Enabled $false -Company "$CompanyName" -Office "$CompanyOffice" -Country "$CompanyCountryCode" -Description "AD User for Shared Mailbox $SharedMBXEmailAddress" | Out-Null
}
# EMEA
ElseIf ($ADUserRegion -eq "EMEA") {
    New-ADUser -DisplayName "$ADUserDisplayName" -SamAccountName "$ADUserSamName" -Name "$ADUserDisplayName" -UserPrincipalName $ADUserFullUPN -Path $ADUserOUEMEA -Server $pdce -Enabled $false -Company "$CompanyName" -Office "$CompanyOffice" -Country "$CompanyCountryCode" -Description "AD User for Shared Mailbox $SharedMBXEmailAddress" | Out-Null
}
# NCSA
ElseIf ($ADUserRegion -eq "NCSA") {
    New-ADUser -DisplayName "$ADUserDisplayName" -SamAccountName "$ADUserSamName" -Name "$ADUserDisplayName" -UserPrincipalName $ADUserFullUPN -Path $ADUserOUNCSA -Server $pdce -Enabled $false -Company "$CompanyName" -Office "$CompanyOffice" -Country "$CompanyCountryCode" -Description "AD User for Shared Mailbox $SharedMBXEmailAddress" | Out-Null
}

# Sleep for 30 seconds
Start-Sleep -Seconds 60

# Create mailbox
Enable-Mailbox $ADUserSamName -Shared | Out-Null
# Write output
Write-Output "A mailbox has been created for ""$ADUserDisplayName""."

# Sleep for 30 seconds
Start-Sleep -Seconds 60

# Set PrimarySmtpAddress for Shared Mailbox (this will also set 
Set-Mailbox -Identity $ADUserName  -PrimarySmtpAddress "$ADUserName@$ADUserUPNSuffix" -EmailAddressPolicyEnabled $false | Out-Null
# Write output
Write-Output "The primary SMTP address for ""$ADUserDisplayName"" has been set to $ADUserName@$ADUserUPNSuffix."

# Output update
Write-Output "Now that the AD user object and mailbox are created, we will create the Security Group that will be used to grant access rights to the Shared Mailbox."

###
# Create group object, mail-enable it, then use it to set permissions on the Shared Mailbox
###
# Define variables for Security Groups that will be used to grant access rights
###
$SecGroupDisplayName = "Sec-Exch-$ADUserRegion-$ADUserName"
$SecGroupName = "Sec-Exch-$ADUserRegion-$ADUserName"

###
# Create Security Group to grant "Full Access" and "Send As" rights
###

# APAC
If ($ADUserRegion -eq "APAC") {
    New-ADGroup -DisplayName "$SecGroupDisplayName" -Name $SecGroupName -GroupCategory Security -GroupScope Universal -Path "$ADGroupOUAPAC" -Server $pdce -Description "Access to mailbox: $SharedMBXEmailAddress" | Out-Null
}

# EMEA
ElseIf ($ADUserRegion -eq "EMEA") {
    New-ADGroup -DisplayName "$SecGroupDisplayName" -Name $SecGroupName -GroupCategory Security -GroupScope Universal -Path "$ADGroupOUEMEA" -Server $pdce -Description "Access to mailbox: $SharedMBXEmailAddress" | Out-Null
}

# NCSA
ElseIf ($ADUserRegion -eq "NCSA") {
    New-ADGroup -DisplayName "$SecGroupDisplayName" -Name $SecGroupName -GroupCategory Security -GroupScope Universal -Path "$ADGroupOUNCSA" -Server $pdce -Description "Access to mailbox: $SharedMBXEmailAddress" | Out-Null
}

# Write output
Write-Output "`n"
Write-Output "The AD group ""$SecGroupDisplayName"" has been created."

# Pause for AD group creation so it can be used to set access rights to newly created Shared Mailbox
Write-Output "`n"
Write-Output "Waiting for AD group creation to be recognized so we can run additional commands against it."
Start-Sleep -Seconds 60

###
# Mail-enable the Security Group
###
Enable-DistributionGroup -Identity $SecGroupDisplayName | Out-Null
# Write output
Write-Output "`n"
Write-Output """$SecGroupDisplayName"" has been mail-enabled. It can now be used to set access rights on the Shared Mailbox."

# Sleep for 30 seconds
Start-Sleep -Seconds 60

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