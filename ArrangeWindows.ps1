#$DebugPreference = "Continue"
$DebugPreference = "SilentlyContinue"

# Import AutoItX module if not loaded
if (!(Get-Module AutoItx))
{
    Import-Module "C:\MyTools\AutoIt-v3\AutoItX\AutoItX.psd1"
}

Write-Host ""

# 2 = Match any substring in the title
Set-AU3Option -Option "WinTitleMatchMode" -Value "2" > $null

function Resize-Window
{
    param($title, $text, $x, $y, $w, $h)

    if (Assert-AU3WinExists -Title $title)
    {
        Write-Host "`"$title`""
        Write-Host "$title found"
        $CurrentPosition = Get-AU3WinPos -Title $title
        Write-Host "Current Position: $CurrentPosition"
        Write-Host "Target Position: $x $y $w $h"
        #Get-AU3WinPos -Title $title
        Move-AU3Win -Title $title -X $x -y $y -Width $w -Height $h | Out-Null
        Write-Host "After Position: $CurrentPosition"
        if ( ($CurrentPosition.x -ne $x))
        {
            Move-AU3Win -Title $title -Text $text -X $x -y $y -Width $w -Height $h | Out-Null
            Write-Host "After2 Position: $CurrentPosition"
        }
        Show-AU3WinActivate -Title $title | Out-Null
        Write-Host ""
    }
}

# Location can be MUK or RED
# Location is determined by an argument which is quicker than an IP range based logic
# $IP = (Get-NetIPAddress|Where-Object InterfaceAlias -eq "Ethernet").IPv4Address
if ($null -ne $args[0])
{
    Write-Debug "in if"
    switch ( $args[0].ToString().ToLower() )
    {
        "r" { $Location = "RED" } # Redmond
        "b" { $Location = "RED" } # Bellevue
        "m" { $Location = "MUK" } # Mukilteo
        "mo" { $Location = "MOB" } # Mukilteo Order Book
        default { $Location = "MUK" }
    }
}
else {
    Write-Debug "in else"
    $Location = "MUK" # When there is no argument for location
}

if ($Location -eq "MUK")
{
    Write-Host "Work location is MUK Vertical."

    # OneNote
    Resize-Window '- OneNote' '' 2300 0 1408 1208

    # Teams
    Resize-Window '| Microsoft Teams' ''  1910 1048 1210 500

    # Chrome
    Resize-Window '- Google Chrome' '' 230 0 1570 1167

    # Edge
    Resize-Window '- Microsoft​ Edge' '' 400 0 1525 1167

    # Excel
    Resize-Window '- Excel' '' 1920 0 1400 1200

    # Outlook
    Resize-Window '- Outlook' '' 0 0 1700 1160

    # Ultra Edit
    Resize-Window ' - UltraEdit 64-bit' '' 300 0 1300 1167

    # Visual Studio
    Resize-Window 'Public - Microsoft Visual Studio' '' 1910 -352 1220 1500

    # Kakao
    Resize-Window 'KakaoTalk' 'OnlineMainView' 2050 150 400 600

}

if ($Location -eq "MUKHO")
{
    Write-Host "Work location is MUK Horizontal."

    # OneNote
    Resize-Window '- OneNote' '' 2300 0 1408 1208

    # Teams
    Resize-Window '| Microsoft Teams' ''  1920 0 1200 1200

    # Chrome
    Resize-Window '- Google Chrome' '' 230 0 1570 1167

    # Edge
    Resize-Window '- Microsoft​ Edge' '' 400 0 1530 1167

    # Excel
    Resize-Window '- Excel' '' 1920 0 1400 1200

    # Outlook
    Resize-Window '- Outlook' '' 0 0 1700 1160

    # Ultra Edit
    Resize-Window ' - UltraEdit 64-bit' '' 300 0 1300 1167

    # Visual Studio
    Resize-Window 'Public - Microsoft Visual Studio' '' 1921 0 1500 1200

    # Kakao
    Resize-Window 'KakaoTalk' 'OnlineMainView' 2050 150 400 600

}

if ($Location -eq "RED")
{
    Write-Host "Work location is RED."

    # Outlook
    Resize-Window '- Outlook' -8 -8 1936 1056

    #Resize-Window [title] [x] [y] [width] [height]
    Resize-Window ' - UltraEdit 64-bit' 0 0 1200 1086

    # Teams
    Resize-Window "| Microsoft Teams" 100 0 1227 1040

    # Chrome
    Resize-Window 'Google Chrome' 300 0 1550 1047

    # OneNote
    Resize-Window '- OneNote' 450 0 1450 1047

    # Excel
    Resize-Window '- Excel' 1920 0 1400 1080

    # Visual Studio
    Resize-Window ' Microsoft Visual Studio' 2600 0 1240 1080
}

# Mukilteo Order Book arrangement
if ($Location -eq "MOB")
{
    Write-Host "Work location is MUK Order Book."

    # Teams
    Resize-Window "| Microsoft Teams" 0 0 1200 1160

    # Chrome
    Resize-Window 'Google Chrome' 150 0 1340 1167

    # Edge
    Resize-Window ' - Microsoft​ Edge' 150 0 1340 1167

    # Excel
    Resize-Window '- Excel' 1920 0 1400 1200

    # Visual Studio Code
    Resize-Window '- Visual Studio Code' 2200 0 1450 1200

    # Ultra Edit
    Resize-Window '- UltraEdit 64-bit' 0 0 780 1167

    # Visual Studio
    Resize-Window ' Microsoft Visual Studio' 780 0 1140 1160
}
