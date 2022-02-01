# Reference file
$confrooms = Import-Csv "c:\temp\HQ-ConfRooms.csv"
# Columns should be SamAccountName and DisplayName

# Get count
$roomcount = ($confrooms).count
Write-Output "Total: $roomcount"

# View staged output
ForEach ($room in $confrooms) {
    $Name = $room.SamAccountName
    $DisplayName = $room.DisplayName
    Write-Output "$Name will be used for $DisplayName"}

# Create Room Mailboxes
ForEach ($room in $confrooms) {
    $Name = $room.SamAccountName
    $DisplayName = $room.DisplayName
    New-Mailbox -Name "$Name" -DisplayName "$DisplayName" -Room}

# Set working hours to "Central Standard Time"
ForEach ($room in $confrooms) {
    $Name = $room.SamAccountName
    Set-MailboxCalendarConfiguration -Identity "$Name" -WorkingHoursTimeZone "Central Standard Time"
}

# Set working hours to Start and End
ForEach ($room in $confrooms) {
    $Name = $room.SamAccountName
    Set-MailboxCalendarConfiguration -Identity "$Name" -WorkingHoursStartTime "06:00:00" -WorkingHoursEndTime "18:00:00"
}

# Set to automatically accept meetings
ForEach ($room in $confrooms) {
    $Name = $room.SamAccountName
    Set-CalendarProcessing -Identity "$Name" -AutomateProcessing AutoAccept
}