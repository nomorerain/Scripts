# =====================================================================================
#       Customization
# =====================================================================================

# Import AutoItX module if not loaded
if (!(Get-Module AutoItx))
{
    Import-Module "C:\MyTools\AutoIt-v3\AutoItX\AutoItX.psd1"
}

# 2 = Match any substring in the title
Set-AU3Option -Option "WinTitleMatchMode" -Value "2" > $null
Set-AU3Option -Option "SendKeyDelay" -Value 35 > $null

$VS = 'Public - Microsoft Visual Studio'

function VS {
    Send-AU3Key "{HOME}Or{TAB}ID{TAB 2}"  > $null
    #Start-Sleep -Milliseconds 10
    #Send-AU3Key "^v"  > $null
    Send-AU3Key $ticket
    Send-AU3Key "{TAB}" > $null
}

$tickets = Get-Content $env:USERPROFILE\downloads\list.txt
$numOfTickets = (Get-Content $env:USERPROFILE\downloads\list.txt).length

    Show-AU3WinActivate -Title $VS > $null
    $counter = 0
    foreach ($ticket in $tickets)
    {
        $counter++
        Set-Clipboard $ticket
        Write-Host ("#{0:d3} / {1:d3}" -f $numOfTickets, $counter)
        VS
    }
