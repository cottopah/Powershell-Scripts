# Initial info
Write-Output "This script is used to connect to Security & Compliance Center when you are setup for MFA"

# Set variables
$UserUPN = Read-Host "Enter your UPN"

# Import the module, requires that you are administrator and are able to run the script
Import-Module ExchangeOnlineManagement

# Connect specifying username, if you already have authenticated to another moduel, you actually do not have to authenticate
Connect-IPPSSession -UserPrincipalName $UserUPN

# This will make sure when you need to reauthenticate after 1 hour that it uses existing token and you don't have to write password and stuff
$global:UserPrincipalName="$UserUPN"