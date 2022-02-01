$userName = Read-Host "Enter UserID of user to check lockout source"

# Import AD module
Import-Module ActiveDirectory

$DomainControllers = Get-ADDomainController -Filter *
$PDCEmulator = ($DomainControllers | Where-Object {$_.OperationMasterRoles -contains "PDCEmulator"})

foreach($pdc in $PDCEmulator){
        $pdcName = $pdc.HostName
        Write-Host "Checking PDCEmulator: $pdcName" 
        Get-WinEvent -ComputerName $pdcName -FilterHashtable @{LogName='Security';Id=4740;StartTime=(Get-Date).AddDays(-1)} | Where-Object {$_.Properties[0].Value -like "*$userName*"} | Select-Object -Property TimeCreated, @{Label='UserName';Expression={$_.Properties[0].Value}},@{Label='ClientName';Expression={$_.Properties[1].Value}}
        }