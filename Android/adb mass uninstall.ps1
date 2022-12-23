<#
    Using this script, you can mass uninstall Android apps.
    - Specify the path to the directory where adb.exe is located on the PC
    - Specify a list of package names to uninstall
        - to get the names of packages currently on the device, run "adb shell pm list packages"
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
# paste names of packages here - each of them will be uninstalled if it exists on the device
# in this example, these are apps on a Huawei tablet
$packageNames = @"
com.myscript.calculator.huawei              
com.myscript.nebo.huawei                    
com.huawei.magazine                         
com.huawei.notepad                          
com.huawei.pcassistant                                        
com.huawei.photos                           
com.huawei.soundrecorder                    
com.huawei.tips                             
com.huawei.videoeditor                      
com.kikaoem.hw.qisiemoji.inputmethod        
com.huawei.hiai
com.huawei.associateassistant
com.huawei.gameassistant
com.huawei.hiassistantoversea
com.huawei.hwread.dz
com.huawei.browser
com.huawei.browserhomepage
com.huawei.HwMultiScreenShot
com.huawei.game.kitserver
com.huawei.gamebox
com.huawei.himovie.overseas
com.huawei.hwsearch
com.huawei.mycenter
com.huawei.systemmanager
com.huawei.welinknow
com.hicloud.android.clone
com.huawei.android.thememanager
com.huawei.android.tips
com.huawei.phoneservice
com.swiftkey.swiftkeyconfigurator
com.touchtype.swiftkey
cn.wps.moffice_i18n
com.huawei.music
com.huawei.android.findmyphone
com.huawei.contacts
com.huawei.contacts.sync
com.huawei.contactscamcard
com.huawei.hidisk
com.huawei.appmarket
com.huawei.stylus.floatmenu
"@ -Split [System.Environment]::NewLine | Where-Object {$_} | ForEach-Object {[regex]::Replace($_, '^\s*|\s*$', '')}

# get existing apps
$apps = adb shell pm list packages | Sort-Object | Where-Object {$_}

# strip to keep only the package name
$apps = $apps | ForEach-Object {[regex]::Replace($_, '^.*\:(.*)\s*$', '$1')}

# prepare a list out of apps existing on the device that are in the list above
$appsFound = $apps | Where-Object {$_ -In $packageNames}

ForEach($app in $appsFound) {
    Write-Host "Uninstalling " -ForegroundColor Gray -NoNewline
    Write-Host $app -ForegroundColor Magenta

    <#
        adb shell pm uninstall command line parameters
            --user 0 - not sure but it's needed for the apps to disappear :)
    #>

    adb shell pm uninstall --user 0 $app
}