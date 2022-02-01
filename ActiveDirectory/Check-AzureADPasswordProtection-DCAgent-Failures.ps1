<#
    :Enabled Mode:
    Event: Fail (due to customer password policy)
    Password change EventID: 10016, 30002
    Password set EventID: 10017, 30003
    
    Event: Fail (due to Microsoft password policy)
    Password change EventID: 10016, 30004
    Password set EventID: 10017, 30005
    
    Event: Fail (due to combined Microsoft and customer password policies)
    Password change EventID: 10016, 30026
    Password set EventID: 10017, 30027
    
    Event: Fail (due to user name)
    Password change EventID: 10016, 30021
    Password set EventID: 10017, 30022

    :Audit Mode:
    Event: Audit-only Pass (would have failed customer password policy)
    Password change EventID: 10024, 30008
    Password set EventID: 10025, 30007

    Event: Audit-only Pass (would have failed Microsoft password policy)
    Password change EventID: 10024, 30010
    Password set EventID: 10025, 30009

    Event: Audit-only Pass (would have failed combined Microsoft and customer password policies)
    Password change EventID: 10024, 30028
    Password set EventID: 10025, 30029

    Event: Audit-only Pass (would have failed due to user name)
    Password change EventID: 10016, 30024
    Password set EventID: 10017, 30023
#>

# Define base variables
$Date = (Get-Date).ToString('ddMMMyyyy')

# Get all DCs in Forest
$AllDCHostnames = (Get-ADForest).Domains | % {Get-ADDomainController -Filter * -Server $_ } | Select -ExpandProperty Name

# Create function to check AD Password Protection Agent DCAgent logs remotely
function CheckDCAgentLogs ($DC) {
    Invoke-Command -ComputerName $DC -ScriptBlock {
        # Define EventLog
        $DCAgentLogs = "Microsoft-AzureADPasswordProtection-DCAgent/Admin"

        # Define possible EventID numbers for failures
        $EnabledEventIDs = "10016","10017","30004","30005","30021","30022","30026","30027"
        $AuditOnlyEventIDs = "10024","10025","10016","10017","30007","30008","30009","30010","30023","30024","30028","30029"

        # Perform lookup (Enabled Mode)
        ForEach ($EventID in $EnabledEventIDs) {Get-WinEvent $DCAgentLogs -ErrorAction SilentlyContinue| ? {$_.Id -like "$EventID"} | Select TimeCreated,Id,Message | FL}

        # Perform lookup (Audit Mode)
        ForEach ($EventID in $AuditOnlyEventIDs) {Get-WinEvent $DCAgentLogs -ErrorAction SilentlyContinue| ? {$_.Id -like "$EventID"} | Select TimeCreated,Id,Message | FL}  
    }
}

# Perform lookup
$FailureLookup = ForEach ($DC in $AllDCHostnames) {Write-Output "Checking $DC"; CheckDCAgentLogs $DC}

# Export to file
$FailureLookup | Out-File "c:\temp\AzureADPasswordProtection-DCAgent-Failures_$Date.txt"
Write-Output "File saved to c:\temp as AzureADPasswordProtection-DCAgent-Failures_$Date.txt"