# Configure Sent Items behavior for Shared Mailboxes
$sharedmbx = Read-Host "Enter name of Shared Mailbox"
Set-Mailbox -Identity $sharedmbx -MessageCopyForSentAsEnabled $true