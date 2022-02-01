# Import AD Module
Import-Module ActiveDirectory

# Define variables
$ImportFileName = "e:\temp\FL_names_emails_v2.csv"
$ResultFile = "e:\temp\FL_names_emails_details.csv"

# Import CSV content and format as Last, First
$UserList = Import-Csv $ImportFileName | Select-Object @{N='FullName';E={$_.LastName + ", " + $_.FirstName}}

# Define variable with LastFirst as variable
$UserLastFirst = $UserList | Select-Object -ExpandProperty FullName

# Create function to get user details
function GetUserDetails ($FullName) {
    Get-ADUser -Filter {Name -like $FullName} -Properties Mail | Select-Object Name,Mail
}

# Perform lookup
$UserDetails = ForEach ($User in $UserLastFirst) {GetUserDetails $User}

# Export to CSV
$UserDetails | Export-Csv $ResultFile -NoTypeInformation