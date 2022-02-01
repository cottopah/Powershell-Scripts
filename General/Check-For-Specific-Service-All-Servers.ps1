# Set base variables
$ServiceToLookFor = "CiscoAMP"

# Create variable for all active Windows Servers in AD (within 30 days)
$Days = 30
$DaysInactive = (Get-Date).AddDays(-($Days))
$ActiveServers = Get-ADComputer -LDAPFilter "(&(objectcategory=computer)(OperatingSystem=*Windows*Server*))" -Properties LastLogonDate,OperatingSystem | ? {$_.LastLogonDate -gt $DaysInactive} | Select Name,OperatingSystem,LastLogonDate | Sort Name

# Create try/catch function to check for presence of service
function CheckForService ($Server) {
try {
    $CheckService = $ServiceToLookFor = "CiscoAMP*"; Get-Service -Name $ServiceToLookFor -ComputerName $Server -ErrorAction Stop | Out-Null
    If ($CheckService -ne $null) {Write-Output "Present"}
    }
catch [Microsoft.PowerShell.Commands.ServiceCommandException] {
    Write-Output "Not present"
    }
}

# Check against list of servers
$Result=@()
$i = 1
$ActiveServers | ForEach-Object {
    $Server = $_
    $Result += New-Object PSObject -Property ([ordered]@{
    'Server' = $Server.Name
    'Results' = CheckForService $Server.Name
    })
$i++
}
$Result | Sort Server

# Export to CSV
$Result | Sort Server | Export-Csv "c:\temp\ServiceAudit-$ServiceToLookFor.csv" -NoTypeInformation