$searchWord="
A Briefer History of Time
" -replace " ", "+" -replace "`r", "" -replace "`n", ""

$sleeptime = 500

$format = "audiobook-mp3"
$format = "ebook-kindle"

# Lee County
# Start-Process "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" "https://lcls.overdrive.com/search?query=$searchWord&format=$format"
# Start-Sleep -Milliseconds $sleeptime

# Multnomah County
Start-Process "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" "https://multcolib.overdrive.com/search?query=$searchWord&format=$format"
Start-Sleep -Milliseconds $sleeptime

# San Antonio
Start-Process "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" "https://sanantonio.overdrive.com/search?query=$searchWord&format=$format"
Start-Sleep -Milliseconds $sleeptime

# Pierce County
Start-Process "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" "https://piercecounty.overdrive.com/search?query=$searchWord&format=$format"
Start-Sleep -Milliseconds $sleeptime

# Kitsap
Start-Process "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" "https://krl.overdrive.com/search?query=$searchWord&format=$format"
Start-Sleep -Milliseconds $sleeptime

# Sno-isle
Start-Process "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" "https://sno-isle.overdrive.com/search?query=$searchWord&format=$format"
Start-Sleep -Milliseconds $sleeptime

<#
# --- Loans ---
$sleeptime = 600
Start-Process "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" "https://sno-isle.overdrive.com/account/loans"
Start-Sleep -Milliseconds $sleeptime
Start-Process "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" "https://piercecounty.overdrive.com/account/loans"
Start-Sleep -Milliseconds $sleeptime
Start-Process "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" "https://krl.overdrive.com/account/loans"
Start-Sleep -Milliseconds $sleeptime
Start-Process "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" "https://sanantonio.overdrive.com/account/loans"
Start-Sleep -Milliseconds $sleeptime
Start-Process "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" "https://multcolib.overdrive.com/account/loans"
Start-Sleep -Milliseconds $sleeptime
Start-Process "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" "https://lcls.overdrive.com/account/loans"

# --- Holds ---
$sleeptime = 700
Start-Process "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" "https://sno-isle.overdrive.com/account/holds"
Start-Sleep -Milliseconds $sleeptime
Start-Process "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" "https://piercecounty.overdrive.com/account/holds"
Start-Sleep -Milliseconds $sleeptime
Start-Process "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" "https://krl.overdrive.com/account/holds"
Start-Sleep -Milliseconds $sleeptime
Start-Process "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" "https://sanantonio.overdrive.com/account/holds"
Start-Sleep -Milliseconds $sleeptime
Start-Process "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" "https://multcolib.overdrive.com/account/holds"
Start-Sleep -Milliseconds $sleeptime
Start-Process "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" "https://lcls.overdrive.com/account/holds"

Get-ChildItem -Recurse -include *.azw3r, *.azw3f, *.Phl, EndAction*.asc, *.opf, LanguageLayer.*.kll| remove-item
#>
