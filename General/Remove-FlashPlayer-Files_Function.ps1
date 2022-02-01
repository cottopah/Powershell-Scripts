# Define remote computer
$Server = Read-Host "Enter Server Name"

# Create function
function RemoteRemoveFlashFiles ($Server) {
    Invoke-Command -ComputerName $Server -ScriptBlock {
        # Create variables for Flash folders
        $FlashFolder = "C:\Windows\system32\Macromed\Flash"
        $FlashFolderX64 = "C:\Windows\SysWOW64\Macromed\Flash"
    
        # Create function to force stop iexplore
        function ForceStopRunningProcesses {
        param(
            [parameter(Mandatory=$true)] $processName,
                                         $timeout = 5
        )
        [System.Diagnostics.Process[]]$processList = Get-Process $processName -ErrorAction SilentlyContinue

        ForEach ($Process in $processList) {
            # Try gracefully first
            $Process.CloseMainWindow() | Out-Null
        }

        # Check the 'HasExited' property for each process
        for ($i = 0 ; $i -le $timeout; $i++) {
            $AllHaveExited = $True
            $processList | ForEach-Object {
                If (-NOT $_.HasExited) {
                    $AllHaveExited = $False
                }                    
            }
            If ($AllHaveExited -eq $true){
                Return
            }
            Start-Sleep 1
        }
        # If graceful close has failed, loop through 'Stop-Process'
        $processList | ForEach-Object {
            If (Get-Process -ID $_.ID -ErrorAction SilentlyContinue) {
                Stop-Process -Id $_.ID -Force
            }
        }
    }
        
        # Create function to take ownership of folder
        function Take-Ownership-Folder {
        param(
        [String]$Folder
        )
        takeown.exe /A /F $Folder
        $CurrentACL = Get-Acl $Folder
        Write-Host ...Adding NT Authority\SYSTEM to $Folder -Fore Yellow
        $SystemACLPermission = "NT AUTHORITY\SYSTEM","FullControl","ContainerInherit,ObjectInherit","None","Allow"
        $SystemAccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $SystemACLPermission
        $CurrentACL.AddAccessRule($SystemAccessRule)
        Write-Host ...Adding Domain Admins to $Folder -Fore Yellow
        $AdminACLPermission = "TENNANT\Domain Admins","FullControl","ContainerInherit,ObjectInherit","None","Allow"
        $AdminAccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $AdminACLPermission
        $CurrentACL.AddAccessRule($AdminAccessRule)
        Set-Acl -Path $Folder -AclObject $CurrentACL
    }

        # Create function to take ownership of files
        function Take-Ownership-File {
        param(
        [String]$File
        )
        takeown.exe /A /F $File
        $CurrentACL = Get-Acl $File
        Write-Host ...Adding NT Authority\SYSTEM to $File -Fore Yellow
        $SystemACLPermission = "NT AUTHORITY\SYSTEM","FullControl","Allow"
        $SystemAccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $SystemACLPermission
        $CurrentACL.AddAccessRule($SystemAccessRule)
        Write-Host ...Adding Domain Admins to $File -Fore Yellow
        $AdminACLPermission = "TENNANT\Domain Admins","FullControl","Allow"
        $AdminAccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $AdminACLPermission
        $CurrentACL.AddAccessRule($AdminAccessRule)
        Set-Acl -Path $File -AclObject $CurrentACL
    }

        # Create function to unregister OCX
        function UnregisterOCXFile ($Folder) {
        $File = Get-ChildItem -Path $Folder | ? {$_.Name -like "*.ocx"} | Select -ExpandProperty FullName
        regsvr32.exe /u /s $File
    }

        # Create function to remove folder and files
        function RemoveFiles ($Folder) {
        Get-ChildItem -Path $Folder | Remove-Item -Force -Confirm:$false
    }
    
        # Force stop iexplore
        ForceStopRunningProcesses iexplore

        # Take ownership of folders
        (Take-Ownership-Folder $FlashFolder),
        (Take-Ownership-Folder $FlashFolderX64),

        # Take ownership of files
        ((Get-ChildItem -Path $FlashFolder -Recurse).FullName | % {Take-Ownership-File $_ }),
        ((Get-ChildItem -Path $FlashFolderX64 -Recurse).FullName | % {Take-Ownership-File $_ }),

        # Unregister OCX
        (UnregisterOCXFile $FlashFolder),
        (UnregisterOCXFile $FlashFolderX64),

        # Remove files
        (RemoveFiles $FlashFolder),
        (RemoveFiles $FlashFolderX64)
    } -ArgumentList $FlashFolder, $FlashFolderX64
}

# Execute
RemoteRemoveFlashFiles $Server