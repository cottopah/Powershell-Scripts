# Get random DC to run DNS lookups against
$RandomDC = Get-ADDomainController | Select -ExpandProperty Name

# Define variables
$UmbrellaSearchString = "*umb*dns*"

# Perform lookup
$UmbrellaServers = (Get-DnsServerResourceRecord -ZoneName tco.corp -ComputerName $RandomDC | ? {$_.Hostname -like "$UmbrellaSearchString"} | Select -ExpandProperty Hostname) -replace ".{8}$" | Sort

# Create function to get IP and format into table
function GetIP ($Server) {
    Resolve-DnsName -Name $Server | ? {$_.Type -like "A"} | Select -ExpandProperty IPAddress
}

# Generate output
$Result=@()
$i = 1
$UmbrellaServers | ForEach-Object {
    $Server = $_
    $Result += New-Object PSObject -Property ([ordered]@{
    'Name' = $Server
    'IPAddress' = GetIP $Server
    })
$i++
}
$Result