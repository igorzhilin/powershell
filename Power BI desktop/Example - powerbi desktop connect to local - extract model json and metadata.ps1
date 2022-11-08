<#
    This script connects to PBI data model and allows to access the metadata, i.e. tables, columns, relationships.

    Limitations: currently works with a single pbi file only

    Process:
        - close all other pbi files in PBI desktop
        - open pbix file in PBI desktop, wait for it to load
        - run the script, and it will write Metadata xlsx and Model.json with the original name as prefix
    
    How it works:
        - find SSAS process msmdsrv.exe and its network port
        - get SSAS's parent process command line - it should be PBIDesktop.exe with the path to pbix file
        - connect to the SSAS instance - then you can use object Microsoft.AnalysisServices.Tabular

    Dependencies:
        sqlserver 22.0.20
        joinmodule
        importexcel
        psHelpers.ps1
        psPbiDesktop.ps1
#>
cls


. "$PSScriptRoot\psPbiDesktop.ps1"
. "$PSScriptRoot\psHelpers.ps1"

<#
                                                            _                
                                                           | |               
 _   _ ___  ___ _ __   _ __   __ _ _ __ __ _ _ __ ___   ___| |_ ___ _ __ ___ 
| | | / __|/ _ \ '__| | '_ \ / _` | '__/ _` | '_ ` _ \ / _ \ __/ _ \ '__/ __|
| |_| \__ \  __/ |    | |_) | (_| | | | (_| | | | | | |  __/ ||  __/ |  \__ \
 \__,_|___/\___|_|    | .__/ \__,_|_|  \__,_|_| |_| |_|\___|\__\___|_|  |___/
                      | |                                                    
                      |_|
#>

# You can get the pbi process with file name fulfilling the criteria
$pbixFilePathRegex = '.*o2c.*'

Write-Host "Looking for PBI desktop that has the file with this path pattern: " -ForegroundColor Gray -NoNewline
Write-Host "$pbixFilePathRegex" -ForegroundColor Yellow 

$localPBIDesktopInstance = Get-LocalPBIDesktopInstances -pbixFilePathRegex $pbixFilePathRegex

If(-not $localPBIDesktopInstance) {
    Write-Host "Did not find running PBI desktop that has the file with this path pattern open: " -ForegroundColor Gray -NoNewline
    Write-Host "$pbixFilePathRegex" -ForegroundColor Yellow 

    return
}

<#
                                 _   
                                | |  
  ___ ___  _ __  _ __   ___  ___| |_ 
 / __/ _ \| '_ \| '_ \ / _ \/ __| __|
| (_| (_) | | | | | | |  __/ (__| |_ 
 \___\___/|_| |_|_| |_|\___|\___|\__|
#>

$port = $localPBIDesktopInstance.LocalPort
$SSASTabularHostPort = "localhost:$port"
$pbixFilePath = $localPBIDesktopInstance.PbixFilePath

Write-Host "Found process for opened pbix: $pbixFilePath" -ForegroundColor Gray
Write-Host "Local SSAS Tabular running at host/port: $SSASTabularHostPort" -ForegroundColor Gray

$db = Get-SSASTabularDBs -SSASTabularHostPort $SSASTabularHostPort

If(-not $db) {
    Write-Host "Failed to connect to local SSAS behind PBI desktop: " -ForegroundColor Gray -NoNewline
    Write-Host "$SSASTabularHostPort" -ForegroundColor Yellow 

    return
}

<#
           _ _           _                    _            _       _        
          | | |         | |                  | |          | |     | |       
  ___ ___ | | | ___  ___| |_   _ __ ___   ___| |_ __ _  __| | __ _| |_ __ _ 
 / __/ _ \| | |/ _ \/ __| __| | '_ ` _ \ / _ \ __/ _` |/ _` |/ _` | __/ _` |
| (_| (_) | | |  __/ (__| |_  | | | | | |  __/ || (_| | (_| | (_| | || (_| |
 \___\___/|_|_|\___|\___|\__| |_| |_| |_|\___|\__\__,_|\__,_|\__,_|\__\__,_|
#>

$dbMetadata = Get-SSASDBMetadata -SSASTabularDb $db
$relationships = Get-SSASTabularModelRelationships -SSASTabularDb $db
$tablesFields = Get-SSASTabularTablesFields -SSASTabularDb $db
$measures = Get-SSASTabularMeasures -SSASTabularDb $db
$tableScripts = Get-SSASDBTableScripts -SSASTabularDb $db

$modelJSON = Get-SSASTabularModelJSON -SSASTabularDb $db

<#
                            _   
                           | |  
  _____  ___ __   ___  _ __| |_ 
 / _ \ \/ / '_ \ / _ \| '__| __|
|  __/>  <| |_) | (_) | |  | |_ 
 \___/_/\_\ .__/ \___/|_|   \__|
          | |                   
          |_|
#>

# Export
# collect pbix file info to save stuff later
$pbixFileItem = Get-Item $pbixFilePath
$pbixFileName = $pbixFileItem.Name
$pbixFileBaseName = $pbixFileItem.BaseName
$pbixFileDir = $pbixFileItem.DirectoryName

$pbixFileMetadata = $pbixFileItem | Select-Object Name, DirectoryName, Length, CreationTime, LastWriteTime, LastAccessTime

$timestampString = (Get-Date).ToString('yyyy-MM-dd HHmmss')

# Export metadata in Excel
$excelFilePath = "$pbixFileDir\$pbixFileBaseName - Metadata $timestampString.xlsx"

Write-Host "Writing excel with metadata: $excelFilePath"

$objectToSheetMapping = @(
    [PSCustomObject]@{SheetName = "TablesFields"; Data = $tablesFields},
    [PSCustomObject]@{SheetName = "TableScripts"; Data = $tableScripts},
    [PSCustomObject]@{SheetName = "Measures"; Data = $measures},
    [PSCustomObject]@{SheetName = "Relationships"; Data = $relationships},
    [PSCustomObject]@{SheetName = "DBMetadata"; Data = $dbMetadata},
    [PSCustomObject]@{SheetName = "PBIXFileInfo"; Data = $pbixFileMetadata}
)

Remove-Item $excelFilePath -ErrorAction SilentlyContinue

Save-ObjectsToExcelSheets -excelFilePath $excelFilePath -objectToSheetMapping $objectToSheetMapping

# Export model JSON
$modelJSONFilePath = "$pbixFileDir\$pbixFileBaseName - Model $timestampString.json"
Remove-Item $modelJSONFilePath -ErrorAction SilentlyContinue

Write-Host "Writing model.json: $modelJSONFilePath"
$modelJSON | Set-Content -Path $modelJSONFilePath