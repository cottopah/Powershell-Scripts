# Define variables
$userremotembx = Read-Host "Enter UserID of mailbox to work with"

# Disable mailbox
Write-Output "Disabling Remote Mailbox for $userremotembx"
Disable-RemoteMailbox $userremotembx

# Enable mailbox
Write-Output "Enabling Remote Mailbox for $userremotembx"
Enable-RemoteMailbox $userremotembx -RemoteRoutingAddress $userremotembx@tennantco.mail.onmicrosoft.com

# Finish
Write-Output "Completed."