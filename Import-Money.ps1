$DebugPreference = "Continue"
#$DebugPreference = "SilentlyContinue"

# Import AutoItX module if not loaded
if (!(Get-Module AutoItx)) { Import-Module "C:\MyTools\AutoIt-v3\AutoItX\AutoItX.psd1" }

# 2 = Match any substring in the title
Set-AU3Option -Option "WinTitleMatchMode" -Value "2" > $null
Set-AU3Option -Option "SendKeyDelay" -Value 100

$downloadPath = Join-Path -Path $env:USERPROFILE -ChildPath "downloads"
$moneyApp = "- Microsoft Money"

# Check if Money is running
if ( ((get-process | Where-Object name -eq "msmoney" | Measure-Object).Count) -eq 0 ) {
    Write-Host -ForegroundColor Red "Money is not running"
}

function ImportTransaction {
    param (
        [string] $tfile
    )
    Show-AU3WinActivate -Title $moneyApp | Out-Null
    Send-AU3Key -Key "!fid" # > File > Import... > Downloaded Statement...
    
    Wait-AU3Win -Title "Import a File" -Timeout 15
    Show-AU3WinActivate -Title "Import a File" | Out-Null
    Send-AU3Key -Key "!n"   # > File name:
    Write-Debug $($tfile)
    Set-Clipboard -Value $tfile
    Send-AU3Key "^v"
    Send-AU3Key -Key "{ENTER}" # Enter to OK button

    Wait-AU3Win -Title "Import a file" -Timeout 15
    Send-AU3Key -Key "{ENTER}" # Enter to OK button
}

# Collect transaction files
$transactions = Get-ChildItem -Path $downloadPath *.qfx

Write-Debug "$($transactions.count) transaction file(s) found."

foreach ($item in $transactions) {
    ImportTransaction $item.fullname
}