# Set base variables
$ServiceToLookFor = "ManageEngine - Password Sync Agent"

# Create variable for all Domain Controllers
$DCs = Get-ADDomainController -Filter * | Select Name

# Create try/catch function to check for presence of service
function CheckForService ($Server) {
try {
    $CheckService = $ServiceToLookFor = "ManageEngine - Password Sync Agent"; Get-Service -Name $ServiceToLookFor -ComputerName $Server -ErrorAction Stop | Out-Null
    If ($CheckService -ne $null) {Write-Output "Present"}
    }
catch [Microsoft.PowerShell.Commands.ServiceCommandException] {
    Write-Output "Not present"
    }
}

function CheckServiceStatus ($Server) {
try {
    $CheckServiceStatus = $ServiceToLookFor = "ManageEngine - Password Sync Agent"; Get-Service -Name $ServiceToLookFor -ComputerName $Server -ErrorAction Stop | Out-Null
    If ($CheckServiceStatus -ne $null) {$Status = Get-Service -Name $ServiceToLookFor -ComputerName $Server | Select -ExpandProperty Status; Write-Output "$Status"}
    }
catch [Microsoft.PowerShell.Commands.ServiceCommandException] {
    Write-Output "Error"
    }
}

# Check against list of servers
$Result=@()
$i = 1
$DCs | ForEach-Object {
    $Server = $_
    $Result += New-Object PSObject -Property ([ordered]@{
    'Server' = $Server.Name
    "$ServiceToLookFor" = CheckForService $Server.Name
    'Status' = CheckServiceStatus $Server.Name
    })
$i++
}
$Result | Sort Server

# Export to CSV
$Result | Sort Server | Export-Csv "c:\temp\ServiceAudit-$ServiceToLookFor.csv" -NoTypeInformation