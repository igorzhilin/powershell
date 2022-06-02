cls

$downloadRoot = 'C:\temp' # or $env:TEMP
$subdir = 'sonar' # to this subfolder
$calendarFileName = 'Sonar 2022 Calendar.csv'


$downloadPath = Join-Path $downloadRoot $subdir

$stopwatch = New-Object System.Diagnostics.Stopwatch
$stopwatch.Start()

######################################################################################
 ######   ######## ########          ##     ## ########  ##        ######  
##    ##  ##          ##             ##     ## ##     ## ##       ##    ## 
##        ##          ##             ##     ## ##     ## ##       ##       
##   #### ######      ##             ##     ## ########  ##        ######  
##    ##  ##          ##             ##     ## ##   ##   ##             ## 
##    ##  ##          ##             ##     ## ##    ##  ##       ##    ## 
 ######   ########    ##              #######  ##     ## ########  ######  

"Getting URLs of Sonar events"

$pageURLwithURLs = 'https://sonar.es/en/2022/schedules' # take URLs from this page
$baseURL = 'https://sonar.es' # this will be base URL to prefix the URLs extracted from the page
$URLregexToExtract = '.*/en/2022/artists/.*' # we are interested only in artists URLs

$pageDownloaded = Invoke-WebRequest -Uri $pageUrlwithURLs -UseBasicParsing # download the page

$urls = @()
ForEach($link in $pageDownloaded.Links) {
    if ([regex]::IsMatch($link.href, $URLregexToExtract)) {
        $urls += "{0}{1}" -F $baseURL, $link.href
    }
}

"Got {0} URLs" -F $urls.Count

######################################################################################
########   #######  ##      ## ##    ## ##        #######     ###    ########        ######## ##     ## 
##     ## ##     ## ##  ##  ## ###   ## ##       ##     ##   ## ##   ##     ##       ##       ###   ### 
##     ## ##     ## ##  ##  ## ####  ## ##       ##     ##  ##   ##  ##     ##       ##       #### #### 
##     ## ##     ## ##  ##  ## ## ## ## ##       ##     ## ##     ## ##     ##       ######   ## ### ## 
##     ## ##     ## ##  ##  ## ##  #### ##       ##     ## ######### ##     ##       ##       ##     ## 
##     ## ##     ## ##  ##  ## ##   ### ##       ##     ## ##     ## ##     ##       ##       ##     ## 
########   #######   ###  ###  ##    ## ########  #######  ##     ## ########        ######## ##     ## 

# create location to download the webpages
New-Item $downloadPath -ItemType Directory -ErrorAction SilentlyContinue | Out-Null

"Downloading files to {0}" -F $downloadPath

# This function will run in parallel
$saveURLToFile = {
    $url = $args[0]
    $filePath = $args[1]

    $startTime = Get-Date
    $mywatch = [System.Diagnostics.Stopwatch]::StartNew()

    $request = Invoke-WebRequest -Uri $url -UseBasicParsing

    $fileSize = $request.Content.Length
    $request.Content | Set-Content $filePath -Force

    $endTime = Get-Date
    $duration = $mywatch.Elapsed
    
    $out = 1 | Select-Object `
        @{N='URL';E={$url}}, `
        @{N='FilePath';E={$filePath}}, `
        @{N='Size';E={$fileSize}}, `
        @{N='StartTime';E={$startTime}}, `
        @{N='EndTime';E={$endTime}}, `
        @{N='Duration';E={$duration}}
    
    $out
}

# do the setup of parallel downloading - how many jobs should run in parallel
$nrLogicalProcessors = (Get-CimInstance –ClassName Win32_Processor | Select-Object NumberOfLogicalProcessors).NumberOfLogicalProcessors # does not make sense to run more parallel jobs than are logical processors
$maxJobsRunning = $nrLogicalProcessors 

# clean up existing jobs, just in case
Get-Job | Remove-Job 

ForEach($url in $urls) {
    $fileExt = '.html'
    $saveFileName = ($url -replace '[^\w]+','-') + $fileExt
    $saveFilePath = Join-Path $downloadPath $saveFileName
    
    $running = @(Get-Job | Where-Object { $_.State -eq 'Running' })
    
    # if max nr jobs reached, wait until any is finished and more slots open
    if ($running.Count -eq $maxJobsRunning) {
        $running | Wait-Job -Any | Out-Null
    }
    
    Start-Job -ScriptBlock $saveURLToFile -ArgumentList $url, $saveFilePath | Out-Null
}

#Wait for all jobs to finish.
While ($(Get-Job -State Running).count -gt 0){
    start-sleep 1
}

$resultsDownload = @()

#Get information from each job.
foreach($job in Get-Job){
    $resultsDownload += Receive-Job -Id ($job.Id)
}
#Remove all jobs created.
Get-Job | Remove-Job

"Elapsed {0}" -F $stopwatch.Elapsed.ToString()

######################################################################################
########     ###    ########   ######  ########        ######   #######  ##    ##    ###    ########  
##     ##   ## ##   ##     ## ##    ## ##             ##    ## ##     ## ###   ##   ## ##   ##     ## 
##     ##  ##   ##  ##     ## ##       ##             ##       ##     ## ####  ##  ##   ##  ##     ## 
########  ##     ## ########   ######  ######          ######  ##     ## ## ## ## ##     ## ########  
##        ######### ##   ##         ## ##                   ## ##     ## ##  #### ######### ##   ##   
##        ##     ## ##    ##  ##    ## ##             ##    ## ##     ## ##   ### ##     ## ##    ##  
##        ##     ## ##     ##  ######  ########        ######   #######  ##    ## ##     ## ##     ## 

"Parsing webpages to get the event details"

Function ParseSonarHTML {
    # This function opens a Sonar artist html file and parses it. Parsing is of course specific to the structure of Sonar webpages
    [CmdletBinding()]
    param(
        $filePath,
        $url
    )

    # hardcoded year and month because they are known
    $year = 2022
    $month = 6

    $singleFileHTML = Get-Content -path $filePath -raw 
    $HTML = New-Object -Com "HTMLFile" # make a HTML object out of the file, this allows to select DOM elements
    $HTML.IHTMLDocument2_write($singleFileHTML)

    $headers = $HTML.getElementsByTagName('h2')
    $artist = $headers[$headers.Length - 1].innerText 

    $eventDetails = ($HTML.body.getElementsByClassName('artist-spectacles') | Select-Object innerText).innerText

    $eventDateData = ($eventDetails -split [Environment]::NewLine)[0]
    $eventDateData = [regex]::Replace($eventDateData, '\s+$', '')

    $pattern = '.*?([\d]+) (.*) - (.*)'

    [int]$eventDate = [regex]::Replace($eventDateData, $pattern, '$1')

    # If event starts at 0X:XX, then it is actually a night event on the next date. 
    # There are no day events that start before noon, so it's safe to use the above logic.

    $eventStartDay = $eventDate
    $eventStartTime = [regex]::Replace($eventDateData, $pattern, '$2')
    If ($eventStartTime[0] -eq '0') {
        $eventStartDay++
    }

    $hour = [int]($eventStartTime -split ':')[0]
    $minute = [int]($eventStartTime -split ':')[1]

    $startDateTime = Get-Date -Year $year -Month $month -Day $eventStartDay -Hour $hour -Minute $minute -Second 0 -Millisecond 0

    $eventEndDay = $eventDate
    $eventEndTime = [regex]::Replace($eventDateData, $pattern, '$3')
    If ($eventEndTime[0] -eq '0') {
        $eventEndDay++
    }

    $hour = [int]($eventEndTime -split ':')[0]
    $minute = [int]($eventEndTime -split ':')[1]

    $endDateTime = Get-Date -Year $year -Month $month -Day $eventEndDay -Hour $hour -Minute $minute -Second 0 -Millisecond 0

    $eventLocData = ($eventDetails -split [Environment]::NewLine)[1]
    $eventNight = ($eventLocData -split ' - ')[0]
    $eventVenue = (($eventLocData -split ' - ')[1] -split ' by ')[0]

    $eventVenue = [regex]::Replace($eventVenue, '\s+$','')
    $eventNight = [regex]::Replace($eventNight, '\s+$','')

    1 | Select-Object `
        @{N='Artist';E={$artist}}, 
        @{N='DayNight';E={$eventNight}}, 
        @{N='Venue';E={$eventVenue}},
        @{N='Start';E={$startDateTime}},
        @{N='End';E={$endDateTime}},
        @{N='URL';E={$url}}
}

$resultsParsing = @()

ForEach($result in $resultsDownload) {
    $filePath = $result.FilePath
    $url = $result.URL

    $resultsParsing += ParseSonarHTML -filePath $filePath -url $url 
}

$resultsParsing = $resultsParsing | Sort-Object Start, Venue

$resultsParsing | Format-Table -AutoSize #-Wrap

"Elapsed {0}" -F $stopwatch.Elapsed.ToString()

######################################################################################
 ######    ######     ###    ##             ######## #### ##       ######## 
##    ##  ##    ##   ## ##   ##             ##        ##  ##       ##       
##        ##        ##   ##  ##             ##        ##  ##       ##       
##   #### ##       ##     ## ##             ######    ##  ##       ######   
##    ##  ##       ######### ##             ##        ##  ##       ##       
##    ##  ##    ## ##     ## ##             ##        ##  ##       ##       
 ######    ######  ##     ## ########       ##       #### ######## ######## 

<# 
Google cal expects Outlook CSV format with these columns:

Subject
Start Date
Start Time
End Date
End Time
Description
Location
#>

# $resultsParsing

$calFileName = "Sonar calendar export.csv"
$calFilePath = Join-Path $downloadPath $calFileName

"Creating outlook csv file {0}" -F $calFilePath

$outlookResults = @()

ForEach($result in $resultsParsing) {
    $Subject     = '{0} - {1}' -F $result.Venue, $result.Artist
    $StartDate   = $result.Start.ToString('dd.MM.yyyy')
    $StartTime   = $result.Start.ToString('HH:mm')
    $EndDate     = $result.End.ToString('dd.MM.yyyy')
    $EndTime     = $result.End.ToString('HH:mm')
    $Description = $result.URL
    $Location    = $result.Venue

    $outlookResults += 1 | Select-Object `
        @{N="Subject";E={$Subject}},
        @{N="Start Date";E={$StartDate}},
        @{N="Start Time";E={$StartTime}},
        @{N="End Date";E={$EndDate}},
        @{N="End Time";E={$EndTime}},
        @{N="Description";E={$Description}},
        @{N="Location";E={$Location}}
}

$outlookResults | Format-Table

$calendarFilePath = Join-Path $downloadPath $calendarFileName
$outlookResults | ConvertTo-Csv -NoTypeInformation | Set-Content $calendarFilePath

"Wrote calendar CSV to import to GCAL: {0}" -F $calendarFilePath

"Elapsed total {0}" -F $stopwatch.Elapsed.ToString()