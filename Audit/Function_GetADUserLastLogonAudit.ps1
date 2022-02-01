function GetADUserLastLogonAudit ($UserID) {
    # Get all DCs
    $AllDCs = Get-ADDomainController -Filter * | Select -ExpandProperty Name | Sort Name

    # Pull value from each DC
    $LoginLookupAllDCs = ForEach ($DC in $AllDCs) { 
        # Get user login details
        $UserLoginDetails = Get-ADUser $UserID -Properties LastLogon,LastLogonDate -Server $DC | `
        Select Name,@{N='UserID';E={$_.SamAccountname}},@{N="LastLogon";Expression={[datetime]::FromFileTime($_.'LastLogon')}}

        # Show results
        $Result=@()
        $i = 1
        $UserLoginDetails | ForEach-Object {
            $Value = $_
            $Result += New-Object PSObject -Property ([ordered]@{
            'Name' = $Value.Name
            'UserID' = $Value.UserID
            'LastLogon' = $Value.LastLogon.ToString('MM/dd/yyyy')
            'AuthenticatingDC' = $DC
            })
        $i++
        }
        $Result | ? {$_.LastLogon -notlike "*1600*"}
    }

    # Find newest LastLogon
    $LoginLookupAllDCs | Sort-Object LastLogon -Descending | Select-Object -First 1
}