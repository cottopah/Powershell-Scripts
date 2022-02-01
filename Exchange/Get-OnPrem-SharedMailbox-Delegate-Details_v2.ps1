# Create function to connect to Exchange 2016 on-premises
function Connect-Exchange2016OnPrem {
    # Define user variables
    $PSConnectADAESKeyFilePath = "\\dcposh101\encryptedcredentials$\psconnect-ad_aes-key.txt"
    $PSConnectADCredentialFilePath = "\\dcposh101\encryptedcredentials$\psconnect-ad_encrypted-pass.txt"

    # Create credential object
    $PSConnectADAdmin = “TENNANT\PSConnect-AD”
    $PSConnectADAESKeyFile = Get-Content $PSConnectADAESKeyFilePath
    $PSConnectADPassFile = Get-Content $PSConnectADCredentialFilePath
    $PSConnectADSecurePass = $PSConnectADPassFile | ConvertTo-SecureString –Key $PSConnectADAESKeyFile
    $PSConnectADCredObject = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $PSConnectADAdmin, $PSConnectADSecurePass

    # Get Exchange 2016 Servers
    $Exchange2016Servers = Get-ADComputer -Filter * | ? {($_.Name -like "D*EX16MBX*") -OR ($_.Name -like "D*EX16CAS*")} | Select -ExpandProperty DNSHostName

    # Get any one Exchange 2016 Server
    $Exchange2016Server = $Exchange2016Servers | Get-Random

    # Set connection session variables
    $PSConnectADSessionOption = New-PSSessionOption -SkipCNCheck -SkipCACheck 

    # Connect to Exchange remote PowerShell
    $OnPremExchangePSSession = New-PSSession -ConnectionUri "http://$Exchange2016Server/PowerShell" -ConfigurationName 'Microsoft.Exchange' -Credential $PSConnectADCredObject -SessionOption $PSConnectADSessionOption -Authentication Kerberos

    # Import PSSession to gain access to Exchange cmdlets
    Import-PSSession -Session $OnPremExchangePSSession -DisableNameChecking -AllowClobber | Out-Null
}

# Connect to Exchange
Connect-Exchange2016OnPrem

# Import AD module
Import-Module ActiveDirectory

# Set initial variables
$StartTime = (Get-Date).ToString('ddMMMyyyy_hhmmtt')
Write-Output "Script started: $StartTime"

# Set directory and file variables
$Date = (Get-Date).ToString('ddMMMMyyyy')
mkdir "c:\temp\OnPremSharedMBXAudit" -Force | Out-Null
$WorkingDir = "c:\temp\OnPremSharedMBXAudit"
$ResultFile = "OnPremSharedMBXAudit-$Date.csv"

# Create functions to lookup O365 license groups
function GetO365LicenseGroupTryCatch ($UserID) {    
    # Define O365 license groups
    $Group1 = "Sec-O365-Licensing-Base"
    $Group2 = "Sec-O365-Licensing-M365F3"
    $Group3 = "Sec-O365-Licensing-Base-NoEXO"
    $Group4 = "Sec-O365-Licensing-M365F3-NoEXO"

    # Get DistinguishedName of Groups
    $Group1DN = (Get-ADGroup $Group1).DistinguishedName
    $Group2DN = (Get-ADGroup $Group2).DistinguishedName
    $Group3DN = (Get-ADGroup $Group3).DistinguishedName
    $Group4DN = (Get-ADGroup $Group4).DistinguishedName

    # Try/Catch for checking for valid AD user then check license group
    try{
        $CheckUserID = Get-ADUser -Identity $UserID -ErrorAction Stop | Select-Object -ExpandProperty SamAccountName
        }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]{
        Write-Output "Invalid UserID"
        $CheckUserID = "invalid"
        }
    if ($CheckUserID -ne "invalid") {
        $MemberOf = Get-ADUser -Identity $CheckUserID -Properties MemberOf | Select-Object -ExpandProperty MemberOf
        try{
            $GroupOutput = ""
            If ($MemberOf -contains $Group1DN) {$GroupOutput = -join($GroupOutput+";"+$Group1)}
            If ($MemberOf -contains $Group2DN) {$GroupOutput = -join($GroupOutput+";"+$Group2)}
            If ($MemberOf -contains $Group3DN) {$GroupOutput = -join($GroupOutput+";"+$Group3)}
            If ($MemberOf -contains $Group4DN) {$GroupOutput = -join($GroupOutput+";"+$Group4)}
            If (
                ($MemberOf -notcontains $Group1DN) -and
                ($MemberOf -notcontains $Group2DN) -and
                ($MemberOf -notcontains $Group3DN) -and
                ($MemberOf -notcontains $Group4DN)
                ) {$GroupOutput = "No License Group"}

        Write-Output $GroupOutput
        }
        catch{
            Write-Output "Error checking for MemberOf for $CheckUserID"
        }
    }
}

function GetO365EMSLicenseTryCatch ($UserID) {    
    <#
    This function is used to check to see if a UserID is valid then checks MemberOf for presence of the specified group
    #>
    # Define O365 license group
    $Group5 = "Sec-O365-Licensing-EMS"
    # Get DistinguishedName of Group
    $Group5DN = (Get-ADGroup $Group5).DistinguishedName

    # Try/Catch for checking for valid AD user then check license
    try{
        # Check to see if the UserID defined is valid or not
        $CheckUserID = Get-ADUser -Identity $UserID -ErrorAction Stop | Select-Object -ExpandProperty SamAccountName
        }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]{
        # Catch for error if user doesn't exist
        Write-Output "Invalid UserID"
        $CheckUserID = "invalid"
        }
    if ($CheckUserID -ne "invalid") {
        # Get expanded value of MemberOf to check for presence of group using 'contains'
        $MemberOf = Get-ADUser -Identity $CheckUserID -Properties MemberOf | Select-Object -ExpandProperty MemberOf
        try{
            If ($MemberOf -contains $Group5DN) {Write-Output "$Group5"}
            Else  {Write-Output "No EMS License"}
        }
        catch{
            Write-Output "Error checking for MemberOf for $CheckUserID"
        }
    }
}

# Create remaining functions
function GetMailboxDelegateGroup ($Alias) {
    $GroupACL = Get-Mailbox $Alias | Get-MailboxPermission | ? {($_.User -like "TENNANT\*") -and ($_.IsInherited -like "False")} | Select-Object -ExpandProperty User
    $GroupName = $GroupACL | Where-Object {($_ -like "*Sec*") -or ($_ -like "*TC*")}
    If (($GroupName -like "TENNANT\Sec*") -or ($GroupName -like "TENNANT\-TC*")) {$GroupName.Substring(8)}
    Else {Write-Output "No Group"}
}

function GetADGroupMembers ($Group) {
    If ($Group -like "No Group") {$Members = "No Members"}
    If ($Group -notlike "No Group") {$Members = Get-ADGroupMember $Group | Select -ExpandProperty Name}
    If ($null -ne $Members) {Write-Output $Members}
    If ($null -eq $Members) {Write-Output "No Members"}
}

function GetMailboxRegion ($Alias) {
    $Region = Get-Mailbox $Alias | Select-Object -ExpandProperty DistinguishedName
    If ($Region -like "*NCSA*") {Write-Output "NCSA"}
    If ($Region -like "*EMEA*") {Write-Output "EMEA"}
    If ($Region -like "*APAC*") {Write-Output "APAC"}
    If (($Region -notlike "*NCSA*") -and ($Region -notlike "*EMEA*") -and ($Region -notlike "*APAC*")) {Write-Output "Other"}
}

function GetGroupMailEnabledTryCatch ($Group) {    
    # Try/Catch for checking for valid AD group
    try{
        # Check to see if the group is valid or not
        $CheckGroup = Get-Recipient -Identity $Group -ErrorAction Stop
        }
    catch [System.Management.Automation.RemoteException]{
        # Catch for error if group is not mail-enabled
        Write-Output "Group Not Mail-Enabled"
        $CheckGroup = "invalid"
        }
    if ($CheckGroup -ne "invalid") {
        # Get RecipientType of Mail-Enabled Group
        $GroupType = Get-Recipient $Group -ErrorAction SilentlyContinue | Select-Object -ExpandProperty RecipientType
        try{
            If ($GroupType -like "MailUniversalSecurityGroup") {Write-Output "Mail-Enabled Security"}
            If ($GroupType -like "MailUniversalDistributionGroup") {Write-Output "Distribution"}
        }
        catch{
            Write-Output "Error checking group for $Group"
        }
    }
}

function GetGroupCategory ($Group) {
    If ($Group -like "No Group") {Write-Output "No Group"}
    If ($Group -notlike "No Group") {Get-ADGroup $Group | Select-Object -ExpandProperty GroupCategory}
}

function GetGroupScope ($Group) {
    If ($Group -like "No Group") {Write-Output "No Group"}
    If ($Group -notlike "No Group") {Get-ADGroup $Group | Select-Object -ExpandProperty GroupScope}
}

function GetUserIDTryCatch ($DisplayName) {    
    # Try/Catch for checking for valid AD user
    try{
        # Check to see if the AD user is valid or not
        $CheckDisplayName = @(Get-ADUser -Filter {Name -like $DisplayName} -ErrorAction Stop)
        }
    catch{
        # Catch for error if user doesn't exist
        Write-Output "$DisplayName is invalid"
        }
    if ($CheckDisplayName.Length -ne 0) {$CheckDisplayName | Select-Object -ExpandProperty SamAccountName}
    else {Write-Output "Invalid UserID"}
}

function GetMailboxLocationTryCatch ($UserID) {    
    # Try/Catch for checking for valid AD user then check group membership
    try{
        # Check to see if the mailbox is valid or not
        $CheckMBX = Get-Recipient -Identity $UserID -ErrorAction Stop
        }
    catch [System.Management.Automation.RemoteException]{
        # Catch for error if user doesn't exist
        Write-Output "Invalid UserID"
        $CheckMBX = "invalid"
        }
    if ($CheckMBX -ne "invalid") {
        # Get expanded value of MemberOf to check for presence of group using 'contains'
        $MBXType = Get-Recipient -Identity $UserID | Select-Object -ExpandProperty RecipientTypeDetails
        try{
            If ($MBXType -like "UserMailbox") {Write-Output "OnPrem"}
            If ($MBXType -like "RemoteUserMailbox") {Write-Output "O365"}
        }
        catch{
            Write-Output "Error checking mailbox for $UserID"
        }
    }
}

function GetUserCountryTryCatch ($UserID) {    
    # Try/Catch for checking for valid AD user
    try{
        # Check to see if the AD user is valid or not
        $CheckUserID = @(Get-ADUser -Filter {SamAccountName -like $UserID} -Properties Country -ErrorAction Stop)
        }
    catch{
        # Catch for error if user doesn't exist
        Write-Output "$UserID is invalid"
        }
    if ($CheckUserID.Length -ne 0) {$CheckUserID | Select-Object -ExpandProperty Country}
    else {Write-Output "Invalid UserID"}
}

function GetUserLanguageTryCatch ($UserID) {    
    # Try/Catch for checking for valid AD user
    try{
        # Check to see if the AD user is valid or not
        $CheckUserID = @(Get-ADUser -Filter {SamAccountName -like $UserID} -Properties Language -ErrorAction Stop)
        }
    catch{
        # Catch for error if user doesn't exist
        Write-Output "$UserID is invalid"
        }
    if ($CheckUserID.Length -ne 0) {$CheckUserID | Select-Object -ExpandProperty Language}
    else {Write-Output "Invalid UserID"}
}


# Get all Shared Mailboxes
$SharedMailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object {$_.RecipientTypeDetails -like "SharedMailbox"} | Select-Object Name,Alias,SamAccountName

# Perform lookup
$Result=@()
$i = 1
$SharedMailboxes | ForEach-Object {
    $MBX = $_
    $Group = GetMailboxDelegateGroup $MBX.Alias
    $Members = GetADGroupMembers $Group
    $Region  = GetMailboxRegion $MBX.Alias
    $MailEnabled = GetGroupMailEnabledTryCatch $Group
    $GroupCategory = GetGroupCategory $Group
    $GroupScope = GetGroupScope $Group
    ForEach ($Member in $Members) {
       $UserID = GetUserIDTryCatch $Member
       $O365LicenseGroup = GetO365LicenseGroupTryCatch $UserID
       $EMSLicense = GetO365EMSLicenseTryCatch $UserID
       $MBXLocation = GetMailboxLocationTryCatch $UserID
       $UserCountry = GetUserCountryTryCatch $UserID
       $UserLanguage = GetUserLanguageTryCatch $UserID
       $PSObject = New-Object PSObject
       $PSObject | Add-Member -MemberType NoteProperty -Name "SharedMailbox" -Value $MBX.Name
       $PSObject | Add-Member -MemberType NoteProperty -Name "Region" -Value $Region
       $PSObject | Add-Member -MemberType NoteProperty -Name "AccessGroup" -Value $Group
       $PSObject | Add-Member -MemberType NoteProperty -Name "Mail-Enabled" -Value $MailEnabled
       $PSObject | Add-Member -MemberType NoteProperty -Name "GroupCategory" -Value $GroupCategory
       $PSObject | Add-Member -MemberType NoteProperty -Name "GroupScope" -Value $GroupScope
       $PSObject | Add-Member -MemberType NoteProperty -Name "GroupMember" -Value $Member
       $PSObject | Add-Member -MemberType NoteProperty -Name "MemberUserID" -Value $UserID
       $PSObject | Add-Member -MemberType NoteProperty -Name "MemberO365License" -Value $O365LicenseGroup
       $PSObject | Add-Member -MemberType NoteProperty -Name "MemberEMSLicense" -Value $EMSLicense
       $PSObject | Add-Member -MemberType NoteProperty -Name "MemberMBXLocation" -Value $MBXLocation
       $PSObject | Add-Member -MemberType NoteProperty -Name "MemberCountry" -Value $UserCountry
       $PSObject | Add-Member -MemberType NoteProperty -Name "MemberLanguage" -Value $UserLanguage
       $Result += $PSObject
    }
$i++
}

# Export to CSV
$Result | Sort-Object SharedMailbox | Export-Csv "$WorkingDir\$ResultFile" -NoTypeInformation -Encoding UTF8

# Set end variables
$StopTime = (Get-Date).ToString('ddMMMyyyy_hhmmtt')
Write-Output "Script completed: $StopTime"