# Create functions
function GetADGroupMembersDetails ($Group) {
    $MembersCheck = Get-ADGroupMember $Group | Select-Object Name,SamAccountName
    If ($MembersCheck -eq $null) {Write-Output "No Members"}
    Else {
        # Perform lookup
        $Result=@()
        $i = 1
        $Group | ForEach-Object {
            $Group = $_
            $Members = Get-ADGroupMember $Group | Select Name,SamAccountName
            ForEach ($Member in $Members) {
                # Add any additional attributes to the $UserDetails line
                $UserDetails = Get-ADUser -Identity $Member.SamAccountName -Properties Language,Country
                $UserID = $UserDetails.SamAccountName
                $UserName = $UserDetails.Name
                $UserLanguage = $UserDetails.Language
                $UserCountry = $UserDetails.Country
                $PSObject = New-Object PSObject
                $PSObject | Add-Member -MemberType NoteProperty -Name "Group" -Value $Group
                $PSObject | Add-Member -MemberType NoteProperty -Name "UserID" -Value $UserID
                $PSObject | Add-Member -MemberType NoteProperty -Name "Name" -Value $UserName
                $PSObject | Add-Member -MemberType NoteProperty -Name "Language" -Value $UserLanguage
                $PSObject | Add-Member -MemberType NoteProperty -Name "Country" -Value $UserCountry
                # Add any attributes added to $UserDetails here using the Add-Member format shown above to add to the resulting output
               $Result += $PSObject
            }
        $i++
        }
        $Result
    }
}