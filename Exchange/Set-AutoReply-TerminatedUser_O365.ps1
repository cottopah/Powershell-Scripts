##########
#
# This script is used to set an auto reply for terminated users. It will auto reply to all messages sent to the user's mailbox and inform senders to contact the terminated user's manager.
#
##########

#####
# Prompt for UserID
#####

$userid = Read-Host "Enter UserID"
$companyname = Read-Host "Enter Company Name"
$custommessage = Read-Host "Enter any custom message to display"

#####
# Connect to Exchange Online via PowerShell script using encrypted password file
#####

# Define user variables
$O365AESKeyFilePath = "C:\Automate\EncryptedCredentials\psconnectews-o365_aes-key.txt"
$O365CredentialFilePath = "C:\Automate\EncryptedCredentials\psconnectews-o365_encrypted-pass.txt"

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
$useradvariables = Get-ADUser $user -Properties EmailAddress | Select Name,GivenName,Surname,UserPrincipalName,EmailAddress

# User variables
$userfirstname = $useradvariables.GivenName
$userlastname = $useradvariables.Surname
$userfullname = $useradvariables.Name
$userfirstlast = "$userfirstname"+" "+"$userlastname"

#####
# Grant Full Access to mailbox
#####

# Define the mailboxes
$usermbx = $useradvariables.UserPrincipalName
$usermbx2 = "PSConnectEWS-O365@tennantco.com"

# Grant FullAccess permissions
Add-MailboxPermission -Identity $usermbx -User $usermbx2 -AccessRights FullAccess -InheritanceType All -Confirm:$false -AutoMapping $false

#####
# Set credential object
#####
$cred = Get-Credential -Message "Use PSConnectEWS-O365 credentials"

#####
# Create Inbox Rule for Terminated User Auto Reply
#####

# Create service variables
$exchService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
$exchService.Credentials = New-Object System.Net.NetworkCredential -ArgumentList $cred.UserName, $cred.Password
$exchService.URL = New-Object Uri("https://outlook.office365.com/EWS/Exchange.asmx")
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