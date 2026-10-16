Start-Sleep 3
if ((Get-DesktopList).Count -lt 2) { New-Desktop }
$desktop2 = Get-Desktop -Index 1 
