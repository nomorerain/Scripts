$vpnname = "MSFTVPN"
$vpn = Get-VpnConnection | Where-Object {$_.Name -eq $vpnname}
if ($vpn.ConnectionStatus -eq "Disconnected")
{
    $cmd = $env:WINDIR + "\System32\rasdial.exe"
    $expression = "$cmd "+"$vpnname"
    Invoke-Expression -Command $expression 
}
