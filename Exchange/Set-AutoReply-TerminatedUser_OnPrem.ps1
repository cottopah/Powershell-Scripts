##########
#
# This script is used to set an auto reply for terminated users. It will auto reply to all messages sent to the user's mailbox and inform senders to contact the termianted user's manager.
#
##########

#####
# Prompt for UserID
#####

$userid = Read-Host "Enter UserID"
$companyname = Read-Host "Enter Company Name"
$custommessage = Read-Host "Enter any custom message to display"

#####
# Connect to Exchange on-premises
#####

# Define user variables
$PSConnectADAESKeyFilePath = "\\dcposh101\encryptedcredentials$\psconnect-ad_aes-key.txt"
$PSConnectADCredentialFilePath = "\\dcposh101\encryptedcredentials$\psconnect-ad_encrypted-pass.txt"

# Create credential object
$PSConnectADAdmin = “TENNANT\PSConnect-AD”
$PSConnectADAESKeyFile = Get-Content $PSConnectADAESKeyFilePath
$PSConnectADPassFile = Get-Content $PSConnectADCredentialFilePath
$PSConnectADSecurePass = $PSConnectADPassFile | ConvertTo-SecureString –Key $PSConnectADAESKeyFile
$PSConnectADCredObject = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $PSConnectADAdmin, $PSConnectADSecurePass

# Set connection session variables
$PSConnectADSessionOption = New-PSSessionOption -SkipCNCheck -SkipCACheck 

# Connect to Exchange remote PowerShell
$OnPremExchangePSSession = New-PSSession -ConnectionUri "https://ems-na.tennantco.com/PowerShell" -ConfigurationName 'Microsoft.Exchange' -Credential $PSConnectADCredObject -SessionOption $PSConnectADSessionOption -Authentication Basic

# Import PSSession to gain access to Exchange cmdlets
Import-PSSession -Session $OnPremExchangePSSession -DisableNameChecking -AllowClobber

#####
# Import additional modules
#####

# Import AD module
Import-Module ActiveDirectory

# Import EWS API module
Import-Module -Name "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"

#####
# Set user variables
#####

$user = Get-ADUser -Identity $userID
$useradvariables = Get-ADUser $user -Properties Manager,EmailAddress | Select Name,GivenName,Surname,Manager,UserPrincipalName,EmailAddress

# User variables
$userfirstname = $useradvariables.GivenName
$userlastname = $useradvariables.Surname
$userfullname = $useradvariables.Name
$userfirstlast = "$userfirstname"+" "+"$userlastname"

# User's manager variables
$usermanagerdn = $useradvariables.Manager
$usermanagername = Get-ADUser $usermanagerdn -Properties EmailAddress | Select Name,GivenName,Surname,EmailAddress
$usermanagerfirstname = $usermanagername.GivenName
$usermanagerlastname = $usermanagername.Surname
$usermanagerfirstlast = "$usermanagerfirstname"+" "+"$usermanagerlastname"
$usermanageremail = $usermanagername.EmailAddress

#####
# Grant Full Access to mailbox
#####

# Define the mailboxes
$usermbx = $useradvariables.UserPrincipalName
$usermbx2 = "PSConnectEWS-OnPrem@tennantco.com"

# Grant FullAccess permissions
Add-MailboxPermission -Identity $usermbx -User $usermbx2 -AccessRights FullAccess -InheritanceType All -Confirm:$false -AutoMapping $false

#####
# Set credential object
#####
$cred = Get-Credential -Message "Use PSConnectEWS-OnPrem credentials"

#####
# Create Inbox Rule for Terminated User Auto Reply
#####

# Create service variables
$exchService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
$exchService.Credentials = New-Object System.Net.NetworkCredential -ArgumentList $cred.UserName, $cred.Password
$exchService.URL = New-Object Uri("https://ems-na.tennantco.com/EWS/Exchange.asmx")
$exchService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$useradvariables.EmailAddress)
$exchService.TraceEnabled = $True
    
# Set template variables
$templateEmail = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage($exchService)
$templateEmail.ItemClass = "IPM.Note.Rules.ReplyTemplate.Microsoft"
$templateEmail.IsAssociated = $true
$templateEmail.Subject = "$userfirstlast is no longer with $CompanyName"
$htmlBodyString = @"
    Hello,<BR>
    <BR>
    $userfirstlast is no longer with $CompanyName.<BR>
    <BR>
    Please contact $usermanagerfirstlast at $usermanageremail with any questions.<BR>
    <BR>
    $custommessage
"@
$templateEmail.Body = New-Object Microsoft.Exchange.WebServices.Data.MessageBody($htmlBodyString)
$PidTagReplyTemplateId = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x65C2, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
$templateEmail.SetExtendedProperty($PidTagReplyTemplateId, [System.Guid]::NewGuid().ToByteArray())
$templateEmail.Save([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)

# Create Inbox Rule
$inboxRule = New-Object Microsoft.Exchange.WebServices.Data.Rule
$inboxRule.DisplayName = "Termination Auto Reply"
$inboxRule.IsEnabled = $true
$inboxRule.Conditions.SentToOrCcMe = $true
$inboxRule.Exceptions.FromAddresses.Add("itsupport@tennantco.com")
$inboxRule.Exceptions.IsAutomaticReply = $true
$inboxRule.Actions.ServerReplyWithMessage = $templateEmail.Id
$createRule = New-Object Microsoft.Exchange.WebServices.Data.CreateRuleOperation[] 1
$createRule[0] = $inboxRule
$exchService.UpdateInboxRules($createRule,$true)

#####
# Cleanup tasks
#####

# Remove FullAccess permissions
Remove-MailboxPermission -Identity $usermbx -User $usermbx2 -AccessRights FullAccess -InheritanceType All -Confirm:$false

# Remove PSSessions
Get-PSSession | Remove-PSSession