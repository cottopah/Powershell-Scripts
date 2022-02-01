##########
#
# Use this script to search for and delete messages in Exchange Server - Part 1: Create and Start Search
#
# Primary commands used:
#
# New-ComplianceSearch - used to create the search based on desired criteria
#
# Start-ComplianceSearch - used to run the search
#
# Articles used for reference: 
#
# "Search for and delete messages in Exchange Server": https://docs.microsoft.com/en-us/exchange/policy-and-compliance/ediscovery/delete-messages?view=exchserver-2019
# "New-ComplianceSearch": https://docs.microsoft.com/en-us/powershell/module/exchange/new-compliancesearch?view=exchange-ps
# "Keyword Query Language (KQL) syntax reference": https://docs.microsoft.com/en-us/sharepoint/dev/general-development/keyword-query-language-kql-syntax-reference
# "Start-ComplianceSearch": https://docs.microsoft.com/en-us/powershell/module/exchange/start-compliancesearch?view=exchange-ps
#
##########

###
# Define variables
###

$Name = Read-Host "Enter name you want for title of search (do not use a space in this name)"
$Sender = Read-Host "Enter message sender address"
$Subject = Read-Host "Enter message subject"
$Received = Read-Host "Enter date message was received (use MM/DD/YYYY format)"

###
# Create search
###

New-ComplianceSearch -Name "$Name" -ContentMatchQuery "(From:$Sender) AND (Received:$Received) AND (Subject:\$Subject)" -ExchangeLocation All -Confirm:$false -Force | Out-Null

###
# Perform search
###

Start-ComplianceSearch -Identity "$Name"

###
# Check status
##

Write-Output "Waiting 30 seconds then checking status of search. Please wait."
Start-Sleep -s 30
Get-ComplianceSearch -Identity "$Name" | Select Name,Status,Items,ContentMatchQuery | FL
Write-Output "Proceed to Part2 script if status is 'Completed'. If it is not, continue to use ""Get-ComplianceSearch -Identity $Name | Select Name,Status,Items"" until you see it as 'Completed'. Then proceed with second script to perform delete action."