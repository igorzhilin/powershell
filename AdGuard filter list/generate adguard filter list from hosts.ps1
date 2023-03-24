<#
This script helps generate adblock filter lists from .hosts files and upload the list as a text snippet to web. 
From the web, it can be loaded into the ad blocker, such as AdGuard.
As an example, I create an AdGuard filter list that blocks Huawei domains. This is useful to block ads in Huawei's Gspace app on Android.

Simply speaking, a filter list is a list of URL patterns for ad blocking software, such as
- uBlock Origin
- AdGuard

The syntax for the filters is described here https://adguard.com/kb/general/ad-filtering/create-own-filters/
    and in a more visual form here: https://adblockplus.org/filter-cheatsheet#blocking

In this script, I don't use a complex AdGuard filtering syntax, I just include the domains as-is 

The script works as follows:
- download .hosts file(s) (e.g. from github)
- extract the domains from those .hosts files
- make a filter list following AdGuard syntax 
- post the list to web as a text snippet (currently I use http://sprunge.us because it does the job)

When the list is posted to web, add its URL to AdGuard:
Settings > Content blocking > Filters > Custom filters > + New custom filter > Add the URL
#>

# provide a list of urls, from which to parse the domain names - these should be raw .hosts files in .hosts syntax
$urls = @"
https://codeberg.org/Cyanic76/Hosts/raw/branch/pages/corporations/hicloud.txt
"@ -Split([Environment]::NewLine) | Where-Object {$_}

<#
           _                  _   
          | |                | |  
  _____  _| |_ _ __ __ _  ___| |_ 
 / _ \ \/ / __| '__/ _` |/ __| __|
|  __/>  <| |_| | | (_| | (__| |_ 
 \___/_/\_\\__|_|  \__,_|\___|\__|

Get the list of domains in a two-pass approach: 
- filter the lines that match the desired structure (i.e. not comments or empty)
- out of those, extract the required pieces by regex replace 
 #>

$newLine = [char]10 # splits the lines in the .hosts file
$filterRegex = '^[^\#\!\s].*' # filter the lines that are not # or ! or empty

# $extractRegex = '^.*?([^\.\s]+\.[^\.\s]+)\s*$' # more violent filtering: extract last second-level-domain.first-level-domain in the line - can filter too much / create false positives!
$extractRegex = '^[^\s]+\s+([^\s]+)\s*$' # extract complete address after IP 

$collectedData = @()

# download all .hosts files and extract the domains
ForEach($url in $urls) {
    Write-Host "Downloading URL: " -NoNewline
    Write-Host $url -ForegroundColor Magenta
    
    $data = Invoke-WebRequest -Uri $url
    
    Write-Host "Downloaded bytes: " -NoNewline
    Write-Host $data.Content.Length -ForegroundColor Magenta
    
    # filtering and extracting in one go
    Write-Host "Filtering and extracting domains..." -ForegroundColor Gray
    $result = $data.Content -Split($newLine) | Where-Object {$_ -match $filterRegex} | ForEach-Object {[regex]::replace($_, $extractRegex, '$1')} 
    $collectedData += $result
}

# deduplicate the complete list of domains
$collectedData = $collectedData | Select-Object -Unique

Write-Host "Extracted unique domains: " -NoNewline
Write-Host $collectedData.Length -ForegroundColor Magenta

Write-Host "Please review this list because it can contain false positives and filter too much!" -ForegroundColor Yellow

$collectedData | Sort-Object | Out-String | Write-Host -ForegroundColor Green

<#
                                 _       
                                | |      
  __ _  ___ _ __   ___ _ __ __ _| |_ ___ 
 / _` |/ _ \ '_ \ / _ \ '__/ _` | __/ _ \
| (_| |  __/ | | |  __/ | | (_| | ||  __/
 \__, |\___|_| |_|\___|_|  \__,_|\__\___|
  __/ |                                  
 |___/

generates AdGuard filters following the format string
#>

Write-Host "Generating filter list..."
$generatedData = @()

$formatString = '||{0}^'

$generatedData = $collectedData | ForEach-Object {$formatString -f $_}

# $generatedData | Sort-Object | Out-String | Write-Host -ForegroundColor Green

<#
             _                 _             _                  _     _                       _     
            | |               | |           (_)                | |   | |                     | |    
 _   _ _ __ | | ___   __ _  __| |  ___ _ __  _ _ __  _ __   ___| |_  | |_ ___   __      _____| |__  
| | | | '_ \| |/ _ \ / _` |/ _` | / __| '_ \| | '_ \| '_ \ / _ \ __| | __/ _ \  \ \ /\ / / _ \ '_ \ 
| |_| | |_) | | (_) | (_| | (_| | \__ \ | | | | |_) | |_) |  __/ |_  | || (_) |  \ V  V /  __/ |_) |
 \__,_| .__/|_|\___/ \__,_|\__,_| |___/_| |_|_| .__/| .__/ \___|\__|  \__\___/    \_/\_/ \___|_.__/ 
      | |                                     | |   | |                                             
      |_|                                     |_|   |_|
#>

$generatedDataString = $generatedData -join [Environment]::NewLine
Write-Host "Posting the list to web: " -NoNewline

$pastedTextUrl = Invoke-RestMethod -Body @{"sprunge" = $generatedDataString} -Method Post -Uri 'http://sprunge.us'
Write-Host $pastedTextUrl -ForegroundColor Green