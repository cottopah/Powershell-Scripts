# Define input file
$Users = Import-Csv "E:\MailContacts\MailContacts-Creation-Template.csv"

# Confirm count
$Count = $Users.Count
Write-Output "Count is: $Count"

# Check values before making changes
ForEach ($User in $Users) {Get-Mailbox -Identity $User.TennantEmail | Select Name,PrimarySmtpAddress,ForwardingAddress}

# Set forwarding to addresses
ForEach ($User in $Users) {Set-Mailbox -Identity $User.TennantEmail -ForwardingAddress $User.ExternalEmail}

# Confirm forwarding
ForEach ($User in $Users) {Get-Mailbox -Identity $User.TennantEmail | Select Name,PrimarySmtpAddress,ForwardingAddress}