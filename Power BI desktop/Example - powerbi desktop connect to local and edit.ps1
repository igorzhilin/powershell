<#
    This script connects to PBI data model and allows to access the metadata, i.e. tables, columns, measures, relationships.
    Allows to edit PBI data model in batch.

    Process:
        - open pbix file in PBI desktop, wait for it to load
        - specify $pbixFilePathRegex below 
        - set up, uncomment commands
    
    Dependencies:
        psPbiDesktop.ps1 must be in the same folder as this script - all functions for PBI are there
            -> it has dependencies too

    See function examples below.
    See also function logic in the psPbiDesktop.ps1
#>

. "$PSScriptRoot\psPbiDesktop.ps1"

cls
<#
                       _                   _   
                      (_)                 | |  
 _   _ ___  ___ _ __   _ _ __  _ __  _   _| |_ 
| | | / __|/ _ \ '__| | | '_ \| '_ \| | | | __|
| |_| \__ \  __/ |    | | | | | |_) | |_| | |_ 
 \__,_|___/\___|_|    |_|_| |_| .__/ \__,_|\__|
                              | |              
                              |_|
#>

# You can get the pbi process with file name fulfilling the criteria
$pbixFilePathRegex = '.*.*'

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
           _                                               _                               
          | |                                             | |                              
  ___ ___ | |_   _ _ __ ___  _ __  ___   ______   ___  ___| |_   _ __  _ __ ___  _ __  ___ 
 / __/ _ \| | | | | '_ ` _ \| '_ \/ __| |______| / __|/ _ \ __| | '_ \| '__/ _ \| '_ \/ __|
| (_| (_) | | |_| | | | | | | | | \__ \          \__ \  __/ |_  | |_) | | | (_) | |_) \__ \
 \___\___/|_|\__,_|_| |_| |_|_| |_|___/          |___/\___|\__| | .__/|_|  \___/| .__/|___/
                                                                | |             | |        
                                                                |_|             |_|
Use case: hide columns
#>

# look for tables with this name pattern - keep .* for all tables
$tableNameRegex = '.*'
# look for cols with this name pattern
$columnNameRegex = '^_|^\$|.*surrogate.*'
#$columnNameRegex = '.*surrogate.*'
# set these properties - refer to psPbiDesktop.ps1 for possible properties and their names
$columnProperties = @{
    IsHidden = $true
}

Set-SSASColumnProperties -db $db -tableNameRegex $tableNameRegex -columnNameRegex $columnNameRegex -columnProperties $columnProperties

<#
           _                                                                          
          | |                                                                         
  ___ ___ | |_   _ _ __ ___  _ __  ___   ______   _ __ ___ _ __   __ _ _ __ ___   ___ 
 / __/ _ \| | | | | '_ ` _ \| '_ \/ __| |______| | '__/ _ \ '_ \ / _` | '_ ` _ \ / _ \
| (_| (_) | | |_| | | | | | | | | \__ \          | | |  __/ | | | (_| | | | | | |  __/
 \___\___/|_|\__,_|_| |_| |_|_| |_|___/          |_|  \___|_| |_|\__,_|_| |_| |_|\___|
Use case: rename columns
#>

# look for tables with this name pattern - keep .* for all tables
$tableNameRegex = '.*'

# look for cols with this name pattern
# case-sensitve regex flag: (?-i) at the start of pattern
# refer to psPbiDesktop.ps1 for the structure of this param
$findReplaceList = @(
    @{Find = '(?-i)\(DesC\)'; Replace = '(desc.)'}
)

Set-SSASColumnNamesFindReplace -db $db -tableNameRegex $tableNameRegex -findReplaceList $findReplaceList

<#
                                                                  _                               
                                                                 | |                              
 _ __ ___   ___  __ _ ___ _   _ _ __ ___  ___   ______   ___  ___| |_   _ __  _ __ ___  _ __  ___ 
| '_ ` _ \ / _ \/ _` / __| | | | '__/ _ \/ __| |______| / __|/ _ \ __| | '_ \| '__/ _ \| '_ \/ __|
| | | | | |  __/ (_| \__ \ |_| | | |  __/\__ \          \__ \  __/ |_  | |_) | | | (_) | |_) \__ \
|_| |_| |_|\___|\__,_|___/\__,_|_|  \___||___/          |___/\___|\__| | .__/|_|  \___/| .__/|___/
                                                                       | |             | |        
                                                                       |_|             |_|
Use case: change formatting of measures
#>

# look for tables with this name pattern - keep .* for all tables
$tableNameRegex = '.*kpi.*'
# look for measures with this name pattern
$measureNameRegex = '%'

$measureProperties = @{
    IsHidden = $false
    FormatString = '0%;-0%;0%'
}

Set-SSASMeasureProperties -db $db -tableNameRegex $tableNameRegex -measureNameRegex $measureNameRegex -measureProperties $measureProperties


<#
                                                                                             
                                                                                             
 _ __ ___   ___  __ _ ___ _   _ _ __ ___  ___   ______   _ __ ___ _ __   __ _ _ __ ___   ___ 
| '_ ` _ \ / _ \/ _` / __| | | | '__/ _ \/ __| |______| | '__/ _ \ '_ \ / _` | '_ ` _ \ / _ \
| | | | | |  __/ (_| \__ \ |_| | | |  __/\__ \          | | |  __/ | | | (_| | | | | | |  __/
|_| |_| |_|\___|\__,_|___/\__,_|_|  \___||___/          |_|  \___|_| |_|\__,_|_| |_| |_|\___|
Use case: bring measure names to a common naming convention
#>

# look for tables with this name pattern - keep .* for all tables
$tableNameRegex = '.*3wm.*'

# look for measures with this name pattern
# case-sensitve regex flag: (?-i) at the start of pattern
# refer to psPbiDesktop.ps1 for the structure of this param

$findReplaceRegexList = @(
    #@{Find = 'rep\.'; Replace = 'reporting'}
    @{Find = 'No match performed %'; Replace = 'Revenue (rep) not matched %'}
)

Set-SSASMeasureNamesFindReplace -db $db -tableNameRegex $tableNameRegex -findReplaceRegexList $findReplaceRegexList


<#
                                                                             _       
                                                                            | |      
 _ __ ___   ___  __ _ ___ _   _ _ __ ___  ___   ______    ___ _ __ ___  __ _| |_ ___ 
| '_ ` _ \ / _ \/ _` / __| | | | '__/ _ \/ __| |______|  / __| '__/ _ \/ _` | __/ _ \
| | | | | |  __/ (_| \__ \ |_| | | |  __/\__ \          | (__| | |  __/ (_| | ||  __/
|_| |_| |_|\___|\__,_|___/\__,_|_|  \___||___/           \___|_|  \___|\__,_|\__\___|
Use case: mass create correctly formatted measures or replace existing measures
#>

$tableName = '_Measures - KPIs'

# hardcode

$measureDefinitionList = @(
<#
    FormatString:
        Short Date
        General Date
        #,0
        0%;-0%;0%
        <can leave empty for string measures>

#>
    <#
    @{
        Name = "Analysis period start date"
        Expression = "MIN('Analysis Details'[Period start])"
        FormatString = "Short Date"
    },
    @{
        Name = "Analysis period end date"
        Expression = "MAX('Analysis Details'[Period end])"
        FormatString = "Short Date"
    },
    @{
        Name = "TestMeasure3"
        Expression = "DIVIDE([TestMeasure1],[TestMeasure2])"
        FormatString = "#.#%"
    }
    #>
    @{
        Name = "Entity name"
        Expression = "MAX('Analysis Details'[Client])"
        FormatString = ""
    }
)

# Copypaste and parse tabseparated list

$measureDataList = @"
3-way match (no diffs)	Revenue (rep) 3-way match no differences
3-way match (no diffs) %	Revenue (rep) 3-way match no differences %
3-way match (diffs)	Revenue (rep) 3-way match differences
3-way match (diffs) %	Revenue (rep) 3-way match differences %
2-way match	Revenue (rep) 2-way match
2-way match %	Revenue (rep) 2-way match %
No match performed	Revenue (rep) not matched
No match performed %	Revenue (rep) not matched %
"@ -Split([Environment]::NewLine)

$measureDataList = @"
Revenue 3-way match	Revenue (rep) 3-way match
3-way match	Revenue (rep) 3-way match
2-way match %	Revenue (rep) 2-way match %
3-way match %	Revenue (rep) 3-way match %
"@ -Split([Environment]::NewLine)

$measureDefinitionList = @()
ForEach($measureData in $measureDataList) {
    $MeasureDefinition = @{}
    $MeasureDefinition['Name'] = $measureData.Split("`t")[0]
    $MeasureDefinition['Expression'] = '[{0}]' -f $measureData.Split("`t")[1]
    
    $MeasureDefinition['FormatString'] = "#,0"

    If($MeasureDefinition['Name'] -like '*%*') {
        $MeasureDefinition['FormatString'] = "0%;-0%;0%"   
    }
        
    $measureDefinitionList += $MeasureDefinition
}

# delimited

$measureDataList = @"
3rd party#CALCULATE([Total revenue (rep)], 'Fact - Sales, Revenue Header'[_Intercompany (entity)] = "3rd Party")
Non-3rd party#CALCULATE([Total revenue (rep)], 'Fact - Sales, Revenue Header'[_Intercompany (entity)] <> "3rd Party")
Total revenue#[Total revenue (rep)]
Companies#[Nr companies]
"@ -Split([Environment]::NewLine)

$measureDefinitionList = @()
ForEach($measureData in $measureDataList) {
    $delim = '#'
    $MeasureDefinition = @{}
    $MeasureDefinition['Name'] = $measureData.Split($delim)[0]
    $MeasureDefinition['Expression'] = '{0}' -f $measureData.Split($delim)[1]
    
    $MeasureDefinition['FormatString'] = "#,0"

    If($MeasureDefinition['Name'] -like '*%*') {
        $MeasureDefinition['FormatString'] = "0%;-0%;0%"   
    }
        
    $measureDefinitionList += $MeasureDefinition
}

New-SSASMeasures -db $db -tableName $tableName -measureDefinitionList $measureDefinitionList -replace


<#
 _        _     _                               _                               
| |      | |   | |                             | |                              
| |_ __ _| |__ | | ___  ___   ______   ___  ___| |_   _ __  _ __ ___  _ __  ___ 
| __/ _` | '_ \| |/ _ \/ __| |______| / __|/ _ \ __| | '_ \| '__/ _ \| '_ \/ __|
| || (_| | |_) | |  __/\__ \          \__ \  __/ |_  | |_) | | | (_) | |_) \__ \
 \__\__,_|_.__/|_|\___||___/          |___/\___|\__| | .__/|_|  \___/| .__/|___/
                                                     | |             | |        
                                                     |_|             |_|
Use case: hide tables
#>

$tableNameRegex = '^_helper|^_key|^Row-level'
$tableProperties = @{
    IsHidden = $true
}

Set-SSASTableProperties -db $db -tableNameRegex $tableNameRegex -tableProperties $tableProperties

<#
 _        _     _                                          _       
| |      | |   | |                                        | |      
| |_ __ _| |__ | | ___  ___   ______    ___ _ __ ___  __ _| |_ ___ 
| __/ _` | '_ \| |/ _ \/ __| |______|  / __| '__/ _ \/ _` | __/ _ \
| || (_| | |_) | |  __/\__ \          | (__| | |  __/ (_| | ||  __/
 \__\__,_|_.__/|_|\___||___/           \___|_|  \___|\__,_|\__\___|
use case: DAX tables
#>

$tableDefinitionList = @(
    @{
        Name = '_key_AnalysisId_KOKRS_PRCTR'
        IsHidden = $true
        Expression = "SUMMARIZE('Dim - Profit center hierarchy', [_key_AnalysisId_KOKRS_PRCTR])"
    }
)

New-SSASTables -db $db -tableDefinitionList $tableDefinitionList

<#
 _        _     _                                                        
| |      | |   | |                                                       
| |_ __ _| |__ | | ___  ___   ______   _ __ ___ _ __ ___   _____   _____ 
| __/ _` | '_ \| |/ _ \/ __| |______| | '__/ _ \ '_ ` _ \ / _ \ \ / / _ \
| || (_| | |_) | |  __/\__ \          | | |  __/ | | | | | (_) \ V /  __/
 \__\__,_|_.__/|_|\___||___/          |_|  \___|_| |_| |_|\___/ \_/ \___|
this can seriously damage the data model
#>

$tableNameRegex = '^Calendar'
 Remove-SSASTables -db $db -tableNameRegex $tableNameRegex

<#
          _       _   _                 _     _                                          _       
         | |     | | (_)               | |   (_)                                        | |      
 _ __ ___| | __ _| |_ _  ___  _ __  ___| |__  _ _ __  ___   ______    ___ _ __ ___  __ _| |_ ___ 
| '__/ _ \ |/ _` | __| |/ _ \| '_ \/ __| '_ \| | '_ \/ __| |______|  / __| '__/ _ \/ _` | __/ _ \
| | |  __/ | (_| | |_| | (_) | | | \__ \ | | | | |_) \__ \          | (__| | |  __/ (_| | ||  __/
|_|  \___|_|\__,_|\__|_|\___/|_| |_|___/_| |_|_| .__/|___/           \___|_|  \___|\__,_|\__\___|
                                               | |                                               
                                               |_|
Use case: restore relationships after data model changes. 
Can be restored from saved result of Get-SSASTabularModelRelationships
#>

<#
    IMPORTANT: 
        "from" IS MANY 
        "to"   IS ONE
        
        otherwise it will throw an error

    CrossFilteringBehavior possible values:
        BothDirections
        OneDirection
        Automatic
#>

$relationshipDefinitionList = @(
    @{
        fromTableName = 'Dim - HRC keys'
        fromColumnName = '_key_AnalysisId_HRCGroup'
        toTableName = 'Dim - HRC descriptions'
        toColumnName = '_key_AnalysisId_HRCGroup'
        CrossFilteringBehavior = 'BothDirections'
    }

)

# OR loading and reproducing the relationships from a metadata Excel file
# which is created via Get-SSASTabularModelRelationships
# Using this file works out of the box because column names are the same

$excelFilePath = 'metadata file with relationships.xlsx'
$excelSheetName = 'Relationships'

# $relationshipDefinitionList = Import-Excel -Path $excelFilePath -WorksheetName $excelSheetName

New-SSASRelationships -db $db -relationshipDefinitionList $relationshipDefinitionList