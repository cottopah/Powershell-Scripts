### Define variables

# Prompt for File Server info
$ServerName = Read-Host "Enter Server Name - i.e. DCFILE02"
$ShareName = Read-Host "Enter Share Name - i.e. tdrive$\Shipping"

# Create UNC path
$ShareUNC = "\\$ServerName\$ShareName\"

# Get variable of who is running
$whoami = whoami
$account = $whoami.substring(8)

### Import NTFSSecurity module

# Define paths
$PSModuleSourcePath = "\\DCPOSH101\PSModules$\NTFSSecurity\"
$PSModuleDestinationPath = "C:\Users\$account\Documents\WindowsPowerShell\Modules\NTFSSecurity\"

# Create function to copy PSModule files
function CopyPSModuleFiles {
    if (-not (Test-Path $PSModuleDestinationPath))
    {
        Copy-Item $PSModuleSourcePath $PSModuleDestinationPath -Recurse
    }
}

# Execute function to copy PSModule files
CopyPSModuleFiles

# Import module
Import-Module NTFSSecurity

### Check network share ACL

# Create function to check ACL
function GetShareACL ($ShareUNC) {
    Get-NTFSAccess -Path $ShareUNC | ? {
        ($_.Account -notlike "*owner*") -AND
        ($_.Account -notlike "*system*") -AND
        ($_.Account -notlike "*administrator*") -AND
        ($_.Account -notlike "*installer*") -AND
        ($_.Account -notlike "*AUTHORITY*")
        } | Select FullName,Account,AccessRights,AccessControlType
  }

# Check share ACL
GetShareACL $ShareUNC | FL