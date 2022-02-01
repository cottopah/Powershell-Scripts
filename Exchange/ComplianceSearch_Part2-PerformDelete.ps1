##########
#
# Use this script to search for and delete messages in Exchange Server - Part 2: Perform Delete
#
# Primary commands used:
#
# New-ComplianceSearchAction - used to execute an action based on the results of the search
#
# Note: When using "-Purge -PurgeType SoftDelete" the message goes to the Deletion folder in Recoverable Deleted Items folder.
#
# Articles used for reference: 
#
# "Search for and delete messages in Exchange Server": https://docs.microsoft.com/en-us/exchange/policy-and-compliance/ediscovery/delete-messages?view=exchserver-2019
# "New-ComplianceSearchAction": https://docs.microsoft.com/en-us/powershell/module/exchange/new-compliancesearchaction?view=exchange-ps
#
##########

###
# Create functions
###

function pspause ($message)
{
    # Check if running PowerShell ISE
    if ($psISE)
    {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show("$message")
    }
    else
    {
        Write-Host "$message" -ForegroundColor Yellow
        $x = $host.ui.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}

###
# Define variables
###

$Name = Read-Host "Enter name of Compliance Search"
$NamePurge = "$Name"+"_Purge"

###
# Confirm status before taking action
###
Write-Output "Checking status of search."
Write-Output "`n"
Get-ComplianceSearch -Identity "$Name" | Select Name,Status,Items,ContentMatchQuery | FL
# Pause for chance to cancel or continue
pspause "Do not continue unless status is 'Completed'. If 'Completed' press Enter to continue, otherwise close this window and try again later."

###
# Perform delete action
###
# Perform delete
New-ComplianceSearchAction -SearchName "$Name" -Purge -PurgeType SoftDelete -Confirm:$false | Out-Null
# Pause for action to be performed
Write-Output "Please wait 30 seconds then the status of the action will be displayed."
Start-Sleep -Seconds 30
# Check status
Get-ComplianceSearchAction -Identity "$NamePurge" | Select SearchName,Action,Status | FL
Write-Output "If status is not 'Completed' use ""Get-ComplianceSearchAction -Identity $NamePurge"" to monitor status."