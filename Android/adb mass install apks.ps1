<#
    Using this script, you can batch install Android apps from the PC.
    - Specify the path to the directory where adb.exe is located on the PC
    - Specify the path to the dir with apks on the PC
    - Execute
#>

cls
$apkDir = 'c:\Temp\apk\__\' # install apks from here

# set adb location - it should be the dir where adb.exe is located
$adbRoot = "$($env:UserProfile)\totalcmd\programs\adb"

# add adb location to the session's $PATH - this will allow to use adb straight from PS
$paths = $($env:Path -split ';') | Where-Object {$_}
If($adbRoot -notin $paths) {
    $paths += $adbRoot; 
    $env:Path = ($paths -join ';')
}

# process
Write-Host "Installing apks from: " -NoNewline -ForegroundColor Gray
Write-Host $apkDir -ForegroundColor Magenta

# get all apks in the dir
$items = Get-ChildItem -Path $apkDir | Where-Object Extension -EQ '.apk'

Write-Host "Found apks: " -NoNewline -ForegroundColor Gray
Write-Host $items.Count -ForegroundColor Magenta

# install each
$timerTotal = [Diagnostics.Stopwatch]::StartNew()
ForEach($item in $items) {
    $filePath = $item.FullName
    Write-Host "Installing: " -NoNewline -ForegroundColor Gray
    Write-Host $filePath -ForegroundColor Magenta -NoNewline
    Write-Host (" ({0}) " -f $item.Length.ToString('#,##0')) -ForegroundColor Yellow
    
    $timer = [Diagnostics.Stopwatch]::StartNew()
    adb install "$filePath"
    
    Write-Host "Elapsed: " -ForegroundColor Gray -NoNewline
    Write-Host $timer.Elapsed.ToString() -ForegroundColor Green
}

Write-Host ("Elapsed total: {0}" -F $timerTotal.Elapsed.ToString()) -ForegroundColor Cyan