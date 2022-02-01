# Prompt for variables
$userid = Read-Host "Enter UserID of user to create Remote Mailbox"

# Enable Remote Mailbox
Enable-RemoteMailbox $userid -RemoteRoutingAddress $userid@tennantco.mail.onmicrosoft.com