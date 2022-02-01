# Set time range
$90DaysInactive = 90
$90days = (Get-Date).AddDays(-($90DaysInactive))

# Get all server objects with LastLogonDate less than 90 days
$ActiveComputers90Days = Get-ADComputer -LDAPFilter "(&(objectcategory=computer)(OperatingSystem=*Windows*))" -Properties LastLogonDate,OperatingSystem,OperatingSystemVersion | ? {$_.LastLogonDate -gt $90days} | Select Name,OperatingSystem,OperatingSystemVersion,LastLogonDate | Sort OperatingSystem

# Export results
$ActiveComputers90Days | Export-Csv "d:\temp\active-computers-90days.csv" -NoTypeInformation

# Close
$Shell = New-Object -ComObject "WScript.Shell"
$Button = $Shell.Popup("File is saved to D:\Temp. Click OK to close this window.", 0, "Hello", 0)