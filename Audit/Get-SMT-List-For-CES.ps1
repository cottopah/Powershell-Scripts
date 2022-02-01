<#
# Summary
  
  This script is used to gather info about the Senior Management Team for purposes of populating the Dictionary in Cisco Email Security for the Forged Email Detection function

    :: Process to change SMT_Names dictionary file in CES ::

    1) Generate new file using PowerShell (UTF-8 encoding)
    2) Import new dictionary file (make sure to check "Match whole words" option) [use SMT_DATE for name]
    3) Adjust "Forged_SMT_Names" Incoming Content Filter to reference newly imported dictionary file
    4) Delete old dictionary file that is no longer being used by Content Filter
    5) Commit changes

    PowerShell scripts involved:
    Get-SMT-List-For-CES.ps1
    Export-SMT-Names-Dictionary-For-CES.ps1

#>

###
# Set base variables
###
mkdir "C:\Temp\SMT_Names" -Force | Out-Null
$Today = (Get-Date).ToString('ddMMMyyyy')
$Directory = "C:\Temp\SMT_Names"

###
# Create functions
###

# Get list of Direct Reports (by default, this provides DistinguishedName for each user, but we want Name)
function GetDirectReports ($UserID) {
    $DirectReports = Get-ADUser $UserID -Properties directReports | Select -ExpandProperty directReports
    ForEach ($Person in $DirectReports) {Get-ADUser $Person | ? {$_.Enabled -eq $true} | Select -ExpandProperty Name }
}

# Get additional user details from SamAccountName/UserID
function GetUserDetails ($UserID) {
    Get-ADUser $UserID -Properties Company,Department,Title,Mail | Select Name,SamAccountName,Company,Department,Title,Mail
}

# Get user details from name (instead of SamAccountName/UserID)
function GetUserDetailsFromName ($Name) {
    Get-ADUser -Filter * -Properties Company,Department,Title,Mail | ? {$_.Name -like "$Name"} | Select Name,SamAccountName,Company,Department,Title,Mail
}

# Get manager info from SamAccountName/UserID
function GetUserManager ($UserID) {
    $UserManager = (Get-ADUser -Identity $UserID -Properties Manager).Manager
    If ($UserManager -eq $null)
        {Write-Output "Empty"}
    Else {
            $UserManagerName = Get-ADUser -Identity $UserManager
            $UserManagerFirstName = $UserManagerName.GivenName
            $UserManagerLastName = $UserManagerName.Surname
            $UserManagerLastFirst = "$UserManagerLastName"+", "+"$UserManagerFirstName"
            Write-Output "$UserManagerLastFirst"
        }   
}

# Retrieve name format of "First Last"
function GetNameFirstLast ($UserID) {
    $User = Get-ADUser $UserID -Properties GivenName,Surname
    $UserFirstName = $User.GivenName
    $UserLastName = $User.Surname
    $UserFirstLast = "$UserFirstName"+" "+"$UserLastName"
    Write-Output "$UserFirstLast"
}

###
# Define Group1 - top level SMT
###

$Group1Users = Get-ADUser -Filter * -Properties Department | ? {($_.Department -like "President & CEO - 68280") -and ($_.Enabled -eq $true)} | Select Name,SamAccountName
$Group1UsersDetails = ForEach ($User in $Group1Users) {GetUserDetails $User.SamAccountName}
$Group1DirectReports = ForEach ($User in $Group1UsersDetails) {GetDirectReports $User.SamAccountName}

###
# Define Group2 - second level SMT
###

$Group2Users = $Group1DirectReports | % {GetUserDetailsFromName $_ | Select Name,SamAccountName}
$Group2UsersDetails = ForEach ($User in $Group2Users) {GetUserDetails $User.SamAccountName}
$Group2DirectReports = ForEach ($User in $Group2UsersDetails) {GetDirectReports $User.SamAccountName}

###
# Define Group3 - third level SMT
###

$Group3Users = $Group2DirectReports | % {GetUserDetailsFromName $_ | Select Name,SamAccountName}
$Group3UsersDetails = ForEach ($User in $Group3Users) {GetUserDetails $User.SamAccountName}

###
# Join results - 3 tier
###

$AllUsersDetails3Tier = @($Group1UsersDetails+$Group2UsersDetails+$Group3UsersDetails) | Sort Name -Unique

# Compile complete list and export to CSV
$Result3Tier=@()
$i = 1
$AllUsersDetails3Tier | ForEach-Object {
    $User = $_
    $Result3Tier += New-Object PSObject -Property ([ordered]@{
    'Name' = $User.Name
    'SamAccountName' = $User.SamAccountName
    'Company' = $User.Company
    'Department' = $User.Department
    'Title' = $User.Title
    'Mail' = $User.Mail
    'Manager' = GetUserManager $User.SamAccountName
    'FirstLast' = GetNameFirstLast $User.SamAccountName
    })
$i++
}
$Result3Tier | Export-Csv "$Directory\smt-ces-audit-$Today.csv" -NoTypeInformation -Encoding UTF8

###
# Final steps
###

Write-Output "Results files are saved to: $Directory"