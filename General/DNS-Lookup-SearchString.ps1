# Define variables
$SearchString = Read-Host "Enter search string"

# Define static variables
$DNSServer = Get-ADDomainController -Discover | Select -ExpandProperty Name
$Zone = "tco.corp"

# Perform lookup
$SearchResults = Get-DnsServerResourceRecord -ComputerName $DNSServer -ZoneName $Zone | ? {$_.HostName -like $SearchString} | Select -ExpandProperty HostName
$HostNames = $SearchResults -replace ".{8}$"

# Get HostName and IPAddress of each result
$Lookup = ForEach ($HostName in $HostNames) {Resolve-DnsName -Name $HostName | Select Name,IPAddress}

# Show output
$Result=@()
$i = 1
$Lookup | ForEach-Object {
    $Line = $_
    $Result += New-Object PSObject -Property ([ordered]@{
    'Name' = $Line.Name -replace ".{17}$"
    'IPAddress' = $Line.IPAddress
    })
$i++
}
$Result