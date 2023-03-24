<#
    Using this script, you can mass copy (push) files and directories to Android device from PC.
    - Specify the path to the directory where adb.exe is located on the PC
    - Specify a list of paths to be pushed to Android
    - Specify target location on Android
    - Execute
#>

cls

# set adb location - it should be the dir where adb.exe is located
$adbRoot = "$($env:UserProfile)\totalcmd\programs\adb"

# add adb location to the session's $PATH - this will allow to use adb straight from PS
$paths = $($env:Path -split ';') | Where-Object {$_}
If($adbRoot -notin $paths) {
    $paths += $adbRoot; 
    $env:Path = ($paths -join ';')
}

# process
# list of paths to push to Android
$sourceDirs = @"
c:\Temp\apk1\
c:\Temp\apk2\
"@ -Split([Environment]::NewLine) | Where-Object {$_}

# dir where to push them on Android
$targetOnAndroid = '/storage/emulated/0'

ForEach($sourceDir in $sourceDirs)
{
    Write-Host $sourceDir -ForegroundColor Magenta -NoNewline
    Write-Host " -> " -ForegroundColor Green -NoNewline
    Write-Host $targetOnAndroid -ForegroundColor Cyan 
    
    $timer = [Diagnostics.Stopwatch]::StartNew()

    # remove trailing backslash - otherwise adb errors
    $sourceDir = [regex]::Replace($sourceDir, '[\\\s]+$', '')

    $cmd = 'adb push --sync -Z "{0}" "{1}"' -F $sourceDir, $targetOnAndroid

    Write-Host $cmd -ForegroundColor Green

    <#
        push parameters
            --sync - do not copy if file is present on the device
            -Z - no compression
    #>
    adb push --sync -Z "$sourceDir" "$targetOnAndroid"

    Write-Host "Elapsed: " -ForegroundColor Gray -NoNewline
    Write-Host $timer.Elapsed.ToString() -ForegroundColor Green
}
