####################
#
# This script is used to gather info about all stale accounts in the domain to identify if they can be properly purged
#
####################

####################
# Connect to on-premises Exchange PowerShell
####################

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

# Import AD Module and clear screen of initial output
Import-Module ActiveDirectory
cls

# Prerequisites
mkdir "c:\temp\Stale-Account-Audit" -Force

# Set start time variables
$datetoday = (Get-Date).ToString("MM-dd-yyyy")
$starttime = (Get-Date).ToString('MM-dd-yyyy_hhmm')

Write-Output "Script is running."
Write-Output "Please wait..."
Write-Output "Script started at $starttime"

# Define O365 license groups
$Group1 = "Sec-O365-Licensing-Base"
$Group2 = "Sec-O365-Licensing-M365F3"
$Group3 = "Sec-O365-Licensing-Base-NoEXO"
$Group4 = "Sec-O365-Licensing-M365F3-NoEXO"
$Group5 = "Sec-O365-Licensing-EMS"

# Get DistinguishedName of Groups
$Group1DN = (Get-ADGroup $Group1).DistinguishedName
$Group2DN = (Get-ADGroup $Group2).DistinguishedName
$Group3DN = (Get-ADGroup $Group3).DistinguishedName
$Group4DN = (Get-ADGroup $Group4).DistinguishedName
$Group5DN = (Get-ADGroup $Group5).DistinguishedName

# Create functions to lookup O365 license group
function GetO365LicenseGroup ($User){
    If ((Get-ADUser -Identity $User.SamAccountName -Properties MemberOf | Select -ExpandProperty MemberOf) -contains $Group1DN) {Write-Output $Group1}
    ElseIf ((Get-ADUser -Identity $User.SamAccountName -Properties MemberOf | Select -ExpandProperty MemberOf) -contains $Group2DN) {Write-Output $Group2}
    ElseIf ((Get-ADUser -Identity $User.SamAccountName -Properties MemberOf | Select -ExpandProperty MemberOf) -contains $Group3DN) {Write-Output $Group3}
    ElseIf ((Get-ADUser -Identity $User.SamAccountName -Properties MemberOf | Select -ExpandProperty MemberOf) -contains $Group4DN) {Write-Output $Group4}
    Else {Write-Output "No License Group"}
}

function GetO365EMSLicense ($User) {
    If ((Get-ADUser -Identity $User.SamACcountName -Properties MemberOf | Select -ExpandProperty MemberOf) -contains $Group5DN) {Write-Output $Group5}
    Else {Write-Output "No EMS License"}
}

# Create functions to grab additional AD attributes
function GetADExtensionAttribute7 ($User) {
    Get-ADUser -Identity $User.SamAccountName -Properties ExtensionAttribute7 | Select -ExpandProperty ExtensionAttribute7
}

function GetADExtensionAttribute8 ($User) {
    Get-ADUser -Identity $User.SamAccountName -Properties ExtensionAttribute8 | Select -ExpandProperty ExtensionAttribute8
}

function GetADExtensionAttribute9 ($User) {
    Get-ADUser -Identity $User.SamAccountName -Properties ExtensionAttribute9 | Select -ExpandProperty ExtensionAttribute9
}

function GetADExtensionAttribute11 ($User) {
    Get-ADUser -Identity $User.SamAccountName -Properties ExtensionAttribute11 | Select -ExpandProperty ExtensionAttribute11
}

function GetADExtensionAttribute12 ($User) {
    Get-ADUser -Identity $User.SamAccountName -Properties ExtensionAttribute12 | Select -ExpandProperty ExtensionAttribute12
}

function GetADExtensionAttribute13 ($User) {
    Get-ADUser -Identity $User.SamAccountName -Properties ExtensionAttribute13 | Select -ExpandProperty ExtensionAttribute13
}

function GetADExtensionAttribute14 ($User) {
    Get-ADUser -Identity $User.SamAccountName -Properties ExtensionAttribute14 | Select -ExpandProperty ExtensionAttribute14
}

function GetADExtensionAttribute15 ($User) {
    Get-ADUser -Identity $User.SamAccountName -Properties ExtensionAttribute15 | Select -ExpandProperty ExtensionAttribute15
}

function GetADDivision ($User) {
    Get-ADUser -Identity $User.SamAccountName -Properties Division | Select -ExpandProperty Division
}

function GetADEmployeeID ($User) {
    Get-ADUser -Identity $User.SamAccountName -Properties EmployeeID | Select -ExpandProperty EmployeeID
}

function GetADWhenCreated ($User) {
    $usercreated = (Get-ADUser -Identity $User.SamAccountName -Properties WhenCreated).WhenCreated
    $usercreated.ToString('MM/dd/yyyy')
}

function GetADAccountExpires ($User) {
    $accountexpire = (Get-ADUser -Identity $User.SamAccountName -Properties AccountExpires).AccountExpires
    If ($accountexpire -eq "9223372036854775807") {Write-Output "Doesn't Expire"}
    ElseIf ($accountexpire -eq "0") {Write-Output "Doesn't Expire"}
    Else {
        $expiredate = [datetime]::FromFileTime($accountexpire)
        $expiredate.ToString("MM/dd/yyyy")}
}

function GetADPwdLastSet ($User) {
    $pwdset = (Get-ADUser -Identity $User.SamAccountName -Properties PwdLastSet).PwdLastSet
    $pwdsetdate = [datetime]::FromFileTime($pwdset)
    $pwdlsetdatestring = $pwdsetdate.ToString("MM/dd/yyyy")
    If ($pwdlsetdatestring -eq "12/31/1600") {Write-Output "Never Changed By User"}
    Else {Write-Output $pwdlsetdatestring}
}

function GetADLastLogonTimestamp ($User) {
    $lastlogon = (Get-ADUser -Identity $User.SamAccountName -Properties LastLogonTimestamp).LastLogonTimestamp
    $lastlogondate = [datetime]::FromFileTime($lastlogon)
    $lastlogondatestring = $lastlogondate.ToString("MM/dd/yyyy")
    If ($lastlogondatestring -eq "12/31/1600") {Write-Output "Never Logged In"}
    Else {Write-Output $lastlogondatestring}
}

function GetADTitle ($User) {
    Get-ADUser -Identity $User.SamAccountName -Properties Title | Select -ExpandProperty Title
}

function GetADPhysicalDeliveryOfficeName ($User) {
    $physicaldeliveryofficename = (Get-ADUser -Identity $User.SamAccountName -Properties PhysicalDeliveryOfficeName).PhysicalDeliveryOfficeName
    $physicaldeliveryofficename
}

function GetADDepartment ($User) {
    $department = (Get-ADUser -Identity $User.SamAccountName -Properties Department).Department
    $department
}

function GetADCompany ($User) {
    $company = (Get-ADUser -Identity $User.SamAccountName -Properties Company).Company
    $company
}

function GetADManager ($User) {
    $usermanager = (Get-ADUser -Identity $User.SamAccountName -Properties Manager).Manager
    If ($usermanager -eq $null)
        {Write-Output "Empty"}
    Else {
            $usermanagername = Get-ADUser -Identity $usermanager
            $usermanagerfirstname = $usermanagername.GivenName
            $usermanagerlastname = $usermanagername.Surname
            $usermanagerlastfirst = "$usermanagerlastname"+", "+"$usermanagerfirstname"
            Write-Output "$usermanagerlastfirst"
        }   
}

function GetADManagerOffice ($User) {
    $usermanager = (Get-ADUser -Identity $User.SamAccountName -Properties Manager).Manager
    If ($usermanager -eq $null)
        {Write-Output "Empty"}
    Else {
            $usermanagername = Get-ADUser -Identity $usermanager -Properties Office
            $usermanageroffice = ($usermanagername).Office
            Write-Output "$usermanageroffice"
        }   
}

function GetADUserOU ($User) {
    $userou = (Get-ADUser -Identity $User.SamAccountName -Properties DistinguishedName).DistinguishedName
    If ($userou -like "*OU=Provisioning,OU=Accounts,DC=tennant,DC=tco,DC=corp") {Write-Output "Provisioning"}
    ElseIf ($userou -like "*OU=DisabledAccounts,OU=Disabled,DC=tennant,DC=tco,DC=corp") {Write-Output "DisabledAccounts"}
    ElseIf ($userou -like "*APAC*") {Write-Output "Tennant - APAC"}
    ElseIf ($userou -like "*EMEA*") {Write-Output "Tennant - EMEA"}
    ElseIf ($userou -like "*NCSA*") {Write-Output "Tennant - NCSA"}
    ElseIf ($userou -like "*Office 365*") {Write-Output "Tennant - O365"}
    ElseIf ($userou -like "*IPC*") {Write-Output "IPC"}
    ElseIf ($userou -like "*GaoMei*") {Write-Output "GaoMei"}
    Else {Write-Output "Other"}
}

function GetADEnabled ($User) {
    $enabled = (Get-ADUser -Identity $User.SamAccountName).Enabled
    $enabled
}

function GetADLanguage ($User) {
    Get-ADUser -Identity $User.SamAccountName -Properties Language | Select -ExpandProperty Language
}

function GetMailboxType ($User) {    
    # Try/Catch for checking for valid AD user then check mail-enabled status
    try{
        # Check to see if the UserID defined is a valid mail-enabled object or not
        $CheckMailboxUser = Get-Recipient -Identity $User.SamAccountName -ErrorAction Stop | Select-Object -ExpandProperty SamAccountName
        }
    catch [System.Management.Automation.RemoteException] {
        # Catch for error if user doesn't exist
        Write-Output "Not Mail-Enabled"
        $CheckMailboxUser = "invalid"
        }
    if ($CheckMailboxUser -ne "invalid") {
        # Get RecipientTypeDetails
        $RecipientType = Get-Recipient -Identity $CheckMailboxUser | Select-Object -ExpandProperty RecipientTypeDetails
        try{
            If ($RecipientType -like "UserMailbox") {Write-Output "UserMailbox"}
            If ($RecipientType -like "RemoteUserMailbox") {Write-Output "RemoteUserMailbox"}
            If ($RecipientType -like "SharedMailbox") {Write-Output "SharedMailbox"}
            If ($RecipientType -like "RemoteSharedMailbox") {Write-Output "RemoteSharedMailbox"}
            If ($RecipientType -like "RoomMailbox") {Write-Output "RoomMailbox"}
            If ($RecipientType -like "RemoteRoomMailbox") {Write-Output "RemoteRoomMailbox"}
        }
        catch{
            Write-Output "Error"
        }
    }
}

###
# Get stale user accounts
###

# Requirements
$90Days = (Get-Date).AddDays(-90).ToString('MM/dd/yyyy hh:mm:ss tt')

# Get stale accounts
$StaleUserAccounts = Get-ADUser -Filter * -Properties * | Select Name,SamAccountName,UserPrincipalName,Mail,MemberOf,ExtensionAttribute7,ExtensionAttribute8,ExtensionAttribute9,ExtensionAttribute11,ExtensionAttribute12,ExtensionAttribute13,ExtensionAttribute14,ExtensionAttribute15,Division,EmployeeID,WhenCreated,AccountExpires,PwdLastSet,Manager,Title,PhysicalDeliveryOfficeName,Department,Company,DistinguishedName,Enabled,Language,@{N="lastLogonTimestamp";E={[datetime]::FromFileTime($_.lastLogonTimestamp)}} | ? {$_.LastLogonTimestamp -lt $90Days}
$StaleUserAccountsCount = $StaleUserAccounts.Count

# Display counts
Write-Output "Total stale accounts with no login activity in 90 days: $StaleUserAccountsCount"

###
# Run report
###

$Result=@()
$i = 1
$StaleUserAccounts | ForEach-Object {
    $User = $_
    $Result += New-Object PSObject -Property ([ordered]@{
    'Name' = $User.Name
    'UserID' = $User.SamAccountName
    'UserPrincipalName' = $User.UserPrincipalName
    'EmailAddress' = $User.Mail
    'LastLogonTimestamp' = GetADLastLogonTimestamp $User
    'UserOU' = GetADUserOU $User
    'Enabled' = GetADEnabled $User
    'MailboxType' = GetMailboxType $User
    'LicensingGroup' = GetO365LicenseGroup $User
    'EMSLicense' = GetO365EMSLicense $User
    'EmployeeID' = GetADEmployeeID $User
    'WorkerType' = GetADExtensionAttribute11 $User
    'employee_status' = GetADExtensionAttribute7 $User
    'Division' = GetADDivision $User
    'Title' = GetADTitle $User
    'Office' = GetADPhysicalDeliveryOfficeName $User
    'Department' = GetADDepartment $User
    'Company' = GetADCompany $User
    'EE_is_MGR' = GetADExtensionAttribute8 $User
    'Candidate_ID' = GetADExtensionAttribute9 $User
    'business_function_name' = GetADExtensionAttribute12 $User
    'p_sub_area_name' = GetADExtensionAttribute13 $User
    'geo_location' = GetADExtensionAttribute14 $User
    'p_sub_area' = GetADExtensionAttribute15 $User
    'WhenCreated' = GetADWhenCreated $User
    'AccountExpires' = GetADAccountExpires $User
    'PwdLastSet' = GetADPwdLastSet $User
    'Manager' = GetADManager $User
    'ManagerOffice' = GetADManagerOffice $User
    'Language' = GetADLanguage $User
    })
$i++
}
$Result | Out-Null

###
# Export results
###

# Export to CSV
$Result | Sort Name | Export-Csv "C:\Temp\Stale-Account-Audit\stale-account-audit_details-$datetoday.csv" -NoTypeInformation -Encoding UTF8

# Post-script output
$endtime = (Get-Date).ToString('MM-dd-yyyy_hhmm')
Write-Output "Script has completed."
Write-Output "Script completed at $endtime"

#####
# Email results
#####

# RecipientAddress1
$Splat1 = @{
        SmtpServer  = 'mail.tennant.tco.corp'
        Body        = "Stale Account Audit Report for $datetoday"
        BodyAsHtml  = $false
        To          = 'brian.marcotte@tennantco.com','christian.ottopah@tennantco.com'
        From        = 'StaleAccountAuditReport@tennantco.com'
        Subject     = "Stale Account Audit Report for $datetoday"
        Attachments = (Get-ChildItem -Path "C:\Temp\Stale-Account-Audit\" -Recurse -Include *.csv).FullName
    }
Send-MailMessage @Splat1

#####
# Cleanup files
#####

Remove-Item -Path (Get-ChildItem -Path "C:\Temp\Stale-Account-Audit\" -Recurse -Include *.csv).FullName