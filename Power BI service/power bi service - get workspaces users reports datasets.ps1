<#
    First steps in automation of PBI service (e.g. Power BI Premium)
    As a start, we connect to PBI service and get the token. Using this token, further REST calls are made.
    Make an overview of all workspaces, reports, datasets, and orphanated datasets (report was deleted, but underlying dataset was not)

    The result is saved into an Excel.

    Dependencies:
        JoinModule
        MicrosoftPowerBIMgmt
        ImportExcel
        psHelpers.ps1 (available on github, e.g. https://github.com/igorzhilin/powershell/)
#>

cls
Import-Module MicrosoftPowerBIMgmt
Import-Module JoinModule -DisableNameChecking

. "$PSScriptRoot\psHelpers.ps1"

# save the info in this excel
$excelFilePath = 'C:\temp\PowerBIServiceOverview.xlsx'

# only look into this workspace - .* for all workspaces
$workspaceNameRegex = '.*'

<#
            _     _        _              
           | |   | |      | |             
  __ _  ___| |_  | |_ ___ | | _____ _ __  
 / _` |/ _ \ __| | __/ _ \| |/ / _ \ '_ \ 
| (_| |  __/ |_  | || (_) |   <  __/ | | |
 \__, |\___|\__|  \__\___/|_|\_\___|_| |_|
  __/ |                                   
 |___/
#>

# token is required for further REST API queries
# a browser window with login credentials should pop up

Write-Host "Checking PBI access token" -ForegroundColor Gray
$token = Get-PowerBIAccessToken -ErrorAction SilentlyContinue

If(-not $token) {
    Write-Host "Getting new access token " -ForegroundColor Gray
    Connect-PowerBIServiceAccount | Out-Null
    $token = Get-PowerBIAccessToken
}

<#
            _           _ _                      _                                  
           | |         | | |                    | |                                 
  __ _  ___| |_    __ _| | | __      _____  _ __| | _____ _ __   __ _  ___ ___  ___ 
 / _` |/ _ \ __|  / _` | | | \ \ /\ / / _ \| '__| |/ / __| '_ \ / _` |/ __/ _ \/ __|
| (_| |  __/ |_  | (_| | | |  \ V  V / (_) | |  |   <\__ \ |_) | (_| | (_|  __/\__ \
 \__, |\___|\__|  \__,_|_|_|   \_/\_/ \___/|_|  |_|\_\___/ .__/ \__,_|\___\___||___/
  __/ |                                                  | |                        
 |___/                                                   |_|
#>

$uri = 'https://api.powerbi.com/v1.0/myorg/groups'
$workspaces = Invoke-RestMethod -Method Get -Uri $uri -Headers $token | Select-Object -ExpandProperty value

$workspaces = $workspaces | Where-Object name -match $workspaceNameRegex

<#
            _                               _                         _       _       _                 _       
           | |                             | |                       | |     | |     | |               | |      
  __ _  ___| |_   _ __ ___ _ __   ___  _ __| |_ ___    __ _ _ __   __| |   __| | __ _| |_ __ _ ___  ___| |_ ___ 
 / _` |/ _ \ __| | '__/ _ \ '_ \ / _ \| '__| __/ __|  / _` | '_ \ / _` |  / _` |/ _` | __/ _` / __|/ _ \ __/ __|
| (_| |  __/ |_  | | |  __/ |_) | (_) | |  | |_\__ \ | (_| | | | | (_| | | (_| | (_| | || (_| \__ \  __/ |_\__ \
 \__, |\___|\__| |_|  \___| .__/ \___/|_|   \__|___/  \__,_|_| |_|\__,_|  \__,_|\__,_|\__\__,_|___/\___|\__|___/
  __/ |                   | |                                                                                   
 |___/                    |_|
#>

$reports = @()
$datasets = @()
ForEach($ws in $workspaces) {
    $wsId = $ws.id
    $wsName = $ws.name

    Write-Host "Collecting metadata for reports" -ForegroundColor Gray 
    $uri = 'https://api.powerbi.com/v1.0/myorg/groups/{0}/reports' -f $wsId
    Write-Host ($uri) -ForegroundColor Yellow 

    $result = Invoke-RestMethod -Method Get -Uri $uri -Headers $token | Select-Object -ExpandProperty value
    $reports += $result | Select-Object @{N='wsName';E={$wsName}}, *

    Write-Host "Collecting metadata for datasets" -ForegroundColor Gray 
    $uri = 'https://api.powerbi.com/v1.0/myorg/groups/{0}/datasets' -f $wsId
    Write-Host ($uri) -ForegroundColor Yellow 

    $result = Invoke-RestMethod -Method Get -Uri $uri -Headers $token | Select-Object -ExpandProperty value
    $datasets += $result | Select-Object @{N='wsName';E={$wsName}}, *
}

Write-Host "Reports found: " -ForegroundColor Gray -NoNewline
Write-Host $reports.Count

Write-Host "Datasets found: " -ForegroundColor Gray -NoNewline
Write-Host $datasets.Count

<#
            _                        _                                                       
           | |                      | |                                                      
  __ _  ___| |_  __      _____  _ __| | _____ _ __   __ _  ___ ___   _   _ ___  ___ _ __ ___ 
 / _` |/ _ \ __| \ \ /\ / / _ \| '__| |/ / __| '_ \ / _` |/ __/ _ \ | | | / __|/ _ \ '__/ __|
| (_| |  __/ |_   \ V  V / (_) | |  |   <\__ \ |_) | (_| | (_|  __/ | |_| \__ \  __/ |  \__ \
 \__, |\___|\__|   \_/\_/ \___/|_|  |_|\_\___/ .__/ \__,_|\___\___|  \__,_|___/\___|_|  |___/
  __/ |                                      | |                                             
 |___/                                       |_|
#>

$uri = 'https://api.powerbi.com/v1.0/myorg/groups'
$workspaces = Invoke-RestMethod -Method Get -Uri $uri -Headers $token | Select-Object -ExpandProperty value

Write-Host "Getting PBI workspace users" -ForegroundColor Gray 

$wsUsers = @()
ForEach($ws in $workspaces) {
    $wsName = $ws.name
    Write-Host "Getting users for workspace " -ForegroundColor Gray 
    Write-Host ($wsName) -ForegroundColor Magenta 
    
    $wsId = $ws.id # workspace ID will be used for the REST API
    $uri = 'https://api.powerbi.com/v1.0/myorg/groups/{0}/users' -f $wsId

    Write-Host ($uri) -ForegroundColor Yellow 

    $result = Invoke-RestMethod -Method Get -Uri $uri -Headers $token | Select-Object -ExpandProperty value
    $wsUsers += $result | Select-Object @{N='wsName';E={$wsName}}, *
}

<#
                  _                       _       _                 _       
                 | |                     | |     | |               | |      
  ___  _ __ _ __ | |__   __ _ _ __     __| | __ _| |_ __ _ ___  ___| |_ ___ 
 / _ \| '__| '_ \| '_ \ / _` | '_ \   / _` |/ _` | __/ _` / __|/ _ \ __/ __|
| (_) | |  | |_) | | | | (_| | | | | | (_| | (_| | || (_| \__ \  __/ |_\__ \
 \___/|_|  | .__/|_| |_|\__,_|_| |_|  \__,_|\__,_|\__\__,_|___/\___|\__|___/
           | |                                                              
           |_|
#>

$datasetsReports = Join-Object -JoinType Left -Discern "Reports." -LeftObject $datasets -RightObject $reports -OnExpression {$Left.id -eq $Right.datasetId}
$orphanDatasets = $datasetsReports | Where-Object "Reports.name" -eq $null

Write-Host "Orphanated datasets found: " -ForegroundColor Gray -NoNewline
Write-Host $orphanDatasets.Count
Write-Host ($orphanDatasets.name -join ', ') -ForegroundColor Magenta

<#
                       _                            _ 
                      | |                          | |
 ___  __ ___   _____  | |_ ___     _____  _____ ___| |
/ __|/ _` \ \ / / _ \ | __/ _ \   / _ \ \/ / __/ _ \ |
\__ \ (_| |\ V /  __/ | || (_) | |  __/>  < (_|  __/ |
|___/\__,_| \_/ \___|  \__\___/   \___/_/\_\___\___|_|
#>

Remove-Item $excelFilePath -Force -ErrorAction SilentlyContinue

$objectToSheetMapping = @(
    [PSCustomObject]@{SheetName = "Workspaces"; Data = $workspaces},
    [PSCustomObject]@{SheetName = "Users"; Data = $wsUsers},
    [PSCustomObject]@{SheetName = "Reports"; Data = $reports},
    [PSCustomObject]@{SheetName = "Datasets"; Data = $datasets},
    [PSCustomObject]@{SheetName = "OrphanDatasets"; Data = $orphanDatasets}
)

Save-ObjectsToExcelSheets -excelFilePath $excelFilePath -objectToSheetMapping $objectToSheetMapping
