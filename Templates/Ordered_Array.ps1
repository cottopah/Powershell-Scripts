$Result=@()
$i = 1
$VARIABLELIST | ForEach-Object {
    $VARIABLE = $_
    $Result += New-Object PSObject -Property ([ordered]@{
    'Field1' = $VARIABLE.attribute
    'Field2' = $VARIABLE.attribute
    })
$i++
}
$Result