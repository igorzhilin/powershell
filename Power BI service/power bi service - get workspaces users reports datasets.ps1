<#
    First steps in automation of PBI service (e.g. Power BI Premium)
    Make an overview of all workspaces, reports, datasets, and orphanated datasets (report was deleted, but underlying dataset was not)

    As a start, we connect to PBI service and get the token. Using this token, further REST calls are made.
    To get the creation date of report and dataset, imports API is used.

    The result is saved into an Excel.

    Limitations:
        - You will only be able to see objects to which you have access. For other, you will have (401) Unauthorized.
        - The importCreatedDate & importUpdatedDate (not datetime) are exported to Excel as a string, this is the limitation of ImportExcel module

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
                     _         _                    _               
                    | |       | |                  | |              
  ___ _ __ ___  __ _| |_ ___  | |__   ___  __ _  __| | ___ _ __ ___ 
 / __| '__/ _ \/ _` | __/ _ \ | '_ \ / _ \/ _` |/ _` |/ _ \ '__/ __|
| (__| | |  __/ (_| | ||  __/ | | | |  __/ (_| | (_| |  __/ |  \__ \
 \___|_|  \___|\__,_|\__\___| |_| |_|\___|\__,_|\__,_|\___|_|  |___/
#>

$headers = @{
    'Content-type' = 'application/json'
}

$headers += $token

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
$workspaces = Invoke-RestMethod -Method Get -Uri $uri -Headers $headers | Select-Object -ExpandProperty value

$workspaces = $workspaces | Where-Object name -match $workspaceNameRegex

<#
            _     _                            _       
           | |   (_)                          | |      
  __ _  ___| |_   _ _ __ ___  _ __   ___  _ __| |_ ___ 
 / _` |/ _ \ __| | | '_ ` _ \| '_ \ / _ \| '__| __/ __|
| (_| |  __/ |_  | | | | | | | |_) | (_) | |  | |_\__ \
 \__, |\___|\__| |_|_| |_| |_| .__/ \___/|_|   \__|___/
  __/ |                      | |                       
 |___/                       |_|

imports are used to get the creation date of report and dataset
#>

$imports = @()

ForEach($ws in $workspaces) {
    $wsId = $ws.id
    $wsName = $ws.name

    Write-Host "Collecting imports in workspace " -ForegroundColor Gray -NoNewline
    Write-Host $wsName -ForegroundColor Magenta 

    $uri = 'https://api.powerbi.com/v1.0/myorg/groups/{0}/imports' -f $wsId
    Write-Host ($uri) -ForegroundColor Yellow 

    $result = Invoke-RestMethod -Method Get -Uri $uri -Headers $headers | Select-Object -ExpandProperty value -ErrorAction SilentlyContinue
    $imports += $result | Select-Object @{N='wsName';E={$wsName}}, *
}

# filter out imports not containing reports or datasets
$imports = $imports | Where-Object {$_.reports -or $_.datasets}

# collect datasets/reports with their dates
$datasetsFromImports = @()
$reportsFromImports = @()

ForEach($i in $imports) {
    
    $wsName = $i.wsName
    $importName = $i.name
    $importState = $i.importState
    $createdDateTime = $i.createdDateTime
    $updatedDateTime = $i.updatedDateTime

    ForEach($ds in $i.datasets) {
        $dsName = $ds.name
        $dsId = $ds.id

        $result = [PSCustomObject]@{
            importWsName = $wsName
            importName = $importName
            importState = $importState

            importCreatedDateTime = Get-Date($createdDateTime)
            importCreatedDate = (Get-Date($createdDateTime) -Format d)
            
            importUpdatedDateTime = Get-Date($updatedDateTime)
            importUpdatedDate = (Get-Date($updatedDateTime) -Format d)
            
            datasetId = $dsId
            datasetName = $dsName
        }

        $datasetsFromImports += $result
    }

    ForEach($rep in $i.reports) {
        $repName = $rep.name
        $repId = $rep.id

        $result = [PSCustomObject]@{
            importWsName = $wsName
            importName = $importName
            importState = $importState

            importCreatedDateTime = Get-Date($createdDateTime)
            importCreatedDate = (Get-Date($createdDateTime) -Format d)
            
            importUpdatedDateTime = Get-Date($updatedDateTime)
            importUpdatedDate = (Get-Date($updatedDateTime) -Format d)

            reportId = $repId
            reportName = $repName
        }

        $reportsFromImports += $result
    }

}

<#
            _         _       _                 _                                            _       
           | |       | |     | |               | |         ___                              | |      
  __ _  ___| |_    __| | __ _| |_ __ _ ___  ___| |_ ___   ( _ )    _ __ ___ _ __   ___  _ __| |_ ___ 
 / _` |/ _ \ __|  / _` |/ _` | __/ _` / __|/ _ \ __/ __|  / _ \/\ | '__/ _ \ '_ \ / _ \| '__| __/ __|
| (_| |  __/ |_  | (_| | (_| | || (_| \__ \  __/ |_\__ \ | (_>  < | | |  __/ |_) | (_) | |  | |_\__ \
 \__, |\___|\__|  \__,_|\__,_|\__\__,_|___/\___|\__|___/  \___/\/ |_|  \___| .__/ \___/|_|   \__|___/
  __/ |                                                                    | |                       
 |___/                                                                     |_|
#>

$reports = @()
$datasets = @()

ForEach($ws in $workspaces) {
    $wsId = $ws.id
    $wsName = $ws.name

    Write-Host "Collecting metadata for reports" -ForegroundColor Gray 
    $uri = 'https://api.powerbi.com/v1.0/myorg/groups/{0}/reports' -f $wsId
    Write-Host ($uri) -ForegroundColor Yellow 

    $result = Invoke-RestMethod -Method Get -Uri $uri -Headers $headers | Select-Object -ExpandProperty value -ErrorAction SilentlyContinue
    $reports += $result | Select-Object @{N='wsName';E={$wsName}}, @{N='wsId';E={$wsId}}, *

    Write-Host "Collecting metadata for datasets" -ForegroundColor Gray 
    $uri = 'https://api.powerbi.com/v1.0/myorg/groups/{0}/datasets' -f $wsId
    Write-Host ($uri) -ForegroundColor Yellow 

    $result = Invoke-RestMethod -Method Get -Uri $uri -Headers $headers | Select-Object -ExpandProperty value -ErrorAction SilentlyContinue
    $datasets += $result | Select-Object @{N='wsName';E={$wsName}}, @{N='wsId';E={$wsId}}, *
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
$workspaces = Invoke-RestMethod -Method Get -Uri $uri -Headers $headers | Select-Object -ExpandProperty value

Write-Host "Getting PBI workspace users" -ForegroundColor Gray 

$wsUsers = @()
ForEach($ws in $workspaces) {
    $wsName = $ws.name
    $wsId = $ws.id # workspace ID will be used for the REST API

    Write-Host "Getting users for workspace " -ForegroundColor Gray 
    Write-Host ($wsName) -ForegroundColor Magenta 
    
    $uri = 'https://api.powerbi.com/v1.0/myorg/groups/{0}/users' -f $wsId

    Write-Host ($uri) -ForegroundColor Yellow 

    $result = Invoke-RestMethod -Method Get -Uri $uri -Headers $headers | Select-Object -ExpandProperty value
    $wsUsers += $result | Select-Object @{N='wsName';E={$wsName}}, @{N='wsId';E={$wsId}}, *
}

<#
            _         _       _                 _                             
           | |       | |     | |               | |                            
  __ _  ___| |_    __| | __ _| |_ __ _ ___  ___| |_   _   _ ___  ___ _ __ ___ 
 / _` |/ _ \ __|  / _` |/ _` | __/ _` / __|/ _ \ __| | | | / __|/ _ \ '__/ __|
| (_| |  __/ |_  | (_| | (_| | || (_| \__ \  __/ |_  | |_| \__ \  __/ |  \__ \
 \__, |\___|\__|  \__,_|\__,_|\__\__,_|___/\___|\__|  \__,_|___/\___|_|  |___/
  __/ |                                                                       
 |___/
#>

Write-Host "Getting PBI dataset users" -ForegroundColor Gray 

$dsUsers = @()
ForEach($ds in $datasets) {
    $wsName = $ds.wsName
    
    $dsName = $ds.name
    $dsId = $ds.id 

    Write-Host "Getting users for dataset " -ForegroundColor Gray 
    Write-Host ($dsName) -ForegroundColor Magenta 
    
    $uri = 'https://api.powerbi.com/v1.0/myorg/datasets/{0}/users' -f $dsId

    Write-Host ($uri) -ForegroundColor Yellow 

    $result = Invoke-RestMethod -Method Get -Uri $uri -Headers $headers | Select-Object -ExpandProperty value -ErrorAction SilentlyContinue
    $dsUsers += $result | Select-Object @{N='wsName';E={$wsName}}, @{N='wsId';E={$wsId}}, @{N='dsName';E={$dsName}}, @{N='dsId';E={$dsId}}, *
}

<#
     _                                                               _           _       
    | |                                                             | |         (_)      
  __| |___   _   _ ___  ___ _ __ ___   _ __   ___  _ __     __ _  __| |_ __ ___  _ _ __  
 / _` / __| | | | / __|/ _ \ '__/ __| | '_ \ / _ \| '_ \   / _` |/ _` | '_ ` _ \| | '_ \ 
| (_| \__ \ | |_| \__ \  __/ |  \__ \ | | | | (_) | | | | | (_| | (_| | | | | | | | | | |
 \__,_|___/  \__,_|___/\___|_|  |___/ |_| |_|\___/|_| |_|  \__,_|\__,_|_| |_| |_|_|_| |_|

as a way to get the users that are assigned directly to dataset:
- get ws admin users - these will be in every dataset 
- subtract these users from all ds users
#>

$wsUsersAdmin = $wsUsers | Where-Object groupUserAccessRight -eq 'Admin' | Select-Object wsName, wsId, identifier -Unique

$dsUsersNonAdmin = Join-Object -JoinType Left -LeftObject $dsUsers -RightObject $wsUsersAdmin `
    -OnExpression {$Left.wsName -eq $Right.wsName -and $Left.identifier -eq $Right.identifier} `
    -Discern 'Admin.' |`
    Where-Object "Admin.identifier" -EQ $null |`
    Select-Object wsName, wsId, dsName, dsId, identifier -Unique

<#
                           _                                                             _           _       
                          | |                                                           | |         (_)      
 _ __ ___ _ __   ___  _ __| |_   _   _ ___  ___ _ __ ___   _ __   ___  _ __     __ _  __| |_ __ ___  _ _ __  
| '__/ _ \ '_ \ / _ \| '__| __| | | | / __|/ _ \ '__/ __| | '_ \ / _ \| '_ \   / _` |/ _` | '_ ` _ \| | '_ \ 
| | |  __/ |_) | (_) | |  | |_  | |_| \__ \  __/ |  \__ \ | | | | (_) | | | | | (_| | (_| | | | | | | | | | |
|_|  \___| .__/ \___/|_|   \__|  \__,_|___/\___|_|  |___/ |_| |_|\___/|_| |_|  \__,_|\__,_|_| |_| |_|_|_| |_|
         | |                                                                                                 
         |_|
#>

$reportUsersNonAdmin = Join-Object -JoinType Left -LeftObject $reports -RightObject $dsUsersNonAdmin `
    -OnExpression {$Left.datasetId -eq $Right.dsId -and $Left.datasetWorkspaceId -eq $Right.wsId} `
    -Discern 'NonAdmin.' |`
    Select-Object wsName, name, identifier, webUrl |`
    Where-Object identifier -ne $null

<#
           _     _                       _       _____        _       
          | |   | |                     | |     |  __ \      | |      
  __ _  __| | __| |   ___ _ __ ___  __ _| |_ ___| |  | | __ _| |_ ___ 
 / _` |/ _` |/ _` |  / __| '__/ _ \/ _` | __/ _ \ |  | |/ _` | __/ _ \
| (_| | (_| | (_| | | (__| | |  __/ (_| | ||  __/ |__| | (_| | ||  __/
 \__,_|\__,_|\__,_|  \___|_|  \___|\__,_|\__\___|_____/ \__,_|\__\___|
from imports
#>

$propertiesToExclude = @"
importWsName
reportId
reportName
users
subscriptions
embedUrl
qnaEmbedURL
createReportEmbedURL
datasetId
datasetName
upstreamDatasets
"@ -Split [System.Environment]::NewLine | Where-Object {$_}

$reports = Join-Object -JoinType Left -LeftObject $reports -RightObject $reportsFromImports -OnExpression {$Left.id -eq $Right.reportId} |`
    Select-Object * -ExcludeProperty $propertiesToExclude

$datasets = Join-Object -JoinType Left -LeftObject $datasets -RightObject $datasetsFromImports -OnExpression {$Left.id -eq $Right.datasetId} |`
    Select-Object * -ExcludeProperty $propertiesToExclude

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
    [PSCustomObject]@{SheetName = "WorkspaceUsers"; Data = $wsUsers},
    [PSCustomObject]@{SheetName = "Reports"; Data = $reports},
    [PSCustomObject]@{SheetName = "Datasets"; Data = $datasets},
    [PSCustomObject]@{SheetName = "DatasetUsers"; Data = $dsUsers},
    [PSCustomObject]@{SheetName = "ReportUsersNonAdmin"; Data = $reportUsersNonAdmin}
)

Save-ObjectsToExcelSheets -excelFilePath $excelFilePath -objectToSheetMapping $objectToSheetMapping
