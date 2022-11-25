# Manage Power BI desktop by script

## Problem
If you work with Power BI (PBI) desktop, you are familiar with the "Working on it" window.

![Power BI desktop Working on it](https://i.imgur.com/PS03SLQ.png)

After you do **any change in the data model** (e.g. create or rename a measure, change a relationship cross-filtering direction, hide a column, measure or table, change measure format string), you spend a few seconds looking at this window waiting for PBI desktop to perform the work.

When you need to take a **mass action** in PBI desktop, there is no way other than doing the actions one by one. And each of these actions will be accompanied by "Working on it".

On a related note, if you need to **provide an overview of all PBI desktop model measures, relationships and transform scripts** (e.g. for a peer review, source control or simply for documentation), there is no way to do this out of PBI desktop.

## Solution

While working on a large analytical project, I found a way to access and edit the model behind the PBI desktop using a powershell script. 

I needed to perform the following operations, and I managed to script them:
- Mass hide table columns by name pattern
- Mass rename columns and measures by find-replace
- Mass change format strings of measures (e.g. remove decimal comma, set percentage or thousand separators)
- Mass create measures, tables, relationships
- Mass remove tables
- Extract the PBI desktop model metadata (e.g. for source control or for a review): relationships, measures, table loading code, etc.
- Re-create PBI desktop model objects from model metadata

## How this works
When a pbix file is opened in PBI desktop, the `PBIDesktop.exe` spawns an Analysis Services Tabular process `msmdsrv.exe`, and you can connect to this SSAS Tabular model and change it using .NET namespace `Microsoft.AnalysisServices.Tabular` (coming from PS `sqlserver` module). After the script makes the changes in the model and saves them, the desktop automatically picks them up.

## Files that you need
- [`psPbiDesktop.ps1`](./psPbiDesktop.ps1) contains the functions to connect to local SSAS Tabular instance created by PBI desktop
- [`psHelpers.ps1`](../psHelpers/psHelpers.ps1) contains a bit of PS boilerplate

## Dependencies on PS modules
- sqlserver
- ImportExcel
- JoinModule

## What needs further polishing
1. When a measure is renamed, the script then cascades the change to other measures referencing the old name of the measure and does a find-replace in their DAX formulas. This usually works. But I have not tested this thoroughly.
3. Currently renamed columns are not cascaded to measures. This means, column is renamed, measure cannot find it and breaks. Cascading column name changes is in the plans.
2. Currently any change done on the PBI desktop data model is saved in the data model immediately (each function calls `$db.Model.SaveChanges()` at the end). Further time savings may be gained if `$db.Model.SaveChanges()` is called after all required changes have been made. 
3. For the same reason, und of changes is not implemented - however, in theory the SSAS object model provides method [`Model.UndoLocalChanges`](https://learn.microsoft.com/en-us/dotnet/api/microsoft.analysisservices.tabular.model.undolocalchanges?view=analysisservices-dotnet).

## Examples

### First detect local PBI SSAS Tabular instance and get the DB object
This is the SSAS Tabular DB behind the openen PBI desktop file.

```PowerShell
# The functions are here
. "$PSScriptRoot\psPbiDesktop.ps1"

# if you have a single pbix open, this will work
# otherwise specify here the name of opened pbix file
$pbixFilePathRegex = '.*' 

# Get local PBI desktop instance
$localPBIDesktopInstance = Get-LocalPBIDesktopInstances -pbixFilePathRegex $pbixFilePathRegex

# Get the SSAS Tabular host and port
$port = $localPBIDesktopInstance.LocalPort
$SSASTabularHostPort = "localhost:$port"

# Get the SSAS Tabular DB behind the PBI desktop - there is a single DB
$db = Get-SSASTabularDBs -SSASTabularHostPort $SSASTabularHostPort
```
Then you use this `$db` object to make mass changes.

### Mass hide Power BI desktop columns
If column name begins with _ or $, hide them in all tables.
```PowerShell
# look for tables with this name pattern - keep .* for all tables
$tableNameRegex = '.*'

# look for cols with this name pattern - in this case, columns starting with _ or $
$columnNameRegex = '^_|^\$'

# set these properties - refer to psPbiDesktop.ps1 for possible properties and their names
$columnProperties = @{
    IsHidden = $true
}

Set-SSASColumnProperties -db $db -tableNameRegex $tableNameRegex -columnNameRegex $columnNameRegex -columnProperties $columnProperties
```

### Mass rename Power BI desktop columns
Replace "(DesC)" with "(desc.)" in column names of all tables.
```PowerShell
# look for tables with this name pattern - keep .* for all tables
$tableNameRegex = '.*'

# look for cols with this name pattern
# case-sensitve regex flag: (?-i) at the start of pattern
# refer to psPbiDesktop.ps1 for the structure of this param
$findReplaceList = @(
    @{Find = '(?-i)\(DesC\)'; Replace = '(desc.)'}
)

Set-SSASColumnNamesFindReplace -db $db -tableNameRegex $tableNameRegex -findReplaceList $findReplaceList
```

### Mass change format string of Power BI desktop measures
Look up measures in table with "kpi" in name that have "%" in measure name and set their formatting to percentage.
```PowerShell
# look for tables with this name pattern - keep .* for all tables
$tableNameRegex = '.*kpi.*'
# look for measures with this name pattern
$measureNameRegex = '%'

$measureProperties = @{
    IsHidden = $false
    FormatString = '0%;-0%;0%'
}

Set-SSASMeasureProperties -db $db -tableNameRegex $tableNameRegex -measureNameRegex $measureNameRegex -measureProperties $measureProperties
```

### Mass rename Power BI desktop measures
Look up measures in table with "3wm" in name. Replace "rep." with "reporting" and "No match performed %" with "Revenue (rep) not matched %".
```PowerShell
# look for tables with this name pattern - keep .* for all tables
$tableNameRegex = '.*3wm.*'

# look for measures with this name pattern
# case-sensitve regex flag: (?-i) at the start of pattern
# refer to psPbiDesktop.ps1 for the structure of this param

$findReplaceRegexList = @(
    @{Find = 'rep\.'; Replace = 'reporting'},
    @{Find = 'No match performed %'; Replace = 'Revenue (rep) not matched %'}
)

Set-SSASMeasureNamesFindReplace -db $db -tableNameRegex $tableNameRegex -findReplaceRegexList $findReplaceRegexList
```

### Mass create Power BI desktop measures from definitions
```PowerShell
# measures will be created in this table
$tableName = '_Measures - KPIs'

# create an array of measure definitions
$measureDefinitionList = @(
<#
    Possible values for FormatString:
        Short Date
        General Date
        #,0 - no decimals, thousand separator
        0%;-0%;0% - percentage
         - can leave empty for string measures
#>    
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
    },
    @{
        Name = "Entity name"
        Expression = "MAX('Analysis Details'[Client])"
        FormatString = ""
    }
)

# create the measures and replace currently existing measures
New-SSASMeasures -db $db -tableName $tableName -measureDefinitionList $measureDefinitionList -replace
```

### Mass hide Power BI desktop tables
Hide tables with names starting with "_helper", "_key" or "Row-level"
```PowerShell
$tableNameRegex = '^_helper|^_key|^Row-level'
$tableProperties = @{
    IsHidden = $true
}

Set-SSASTableProperties -db $db -tableNameRegex $tableNameRegex -tableProperties $tableProperties
```

### Mass create Power BI desktop calculated tables from definitions
Create a hidden table with a DAX function.
```PowerShell
$tableDefinitionList = @(
    @{
        Name = '_key_AnalysisId_KOKRS_PRCTR'
        IsHidden = $true
        Expression = "SUMMARIZE('Dim - Profit center hierarchy', [_key_AnalysisId_KOKRS_PRCTR])"
    }
)

New-SSASTables -db $db -tableDefinitionList $tableDefinitionList
```

### Mass remove Power BI desktop tables - ‼ warning: this can seriously damage your data model
Remove tables with names starting with "Calendar"
```PowerShell
$tableNameRegex = '^Calendar'
Remove-SSASTables -db $db -tableNameRegex $tableNameRegex
```

### Mass create Power BI desktop relationships from definitions
You can create many relationships.
```PowerShell
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
        fromTableName = 'Fact - Sales, Revenue Item'
        fromColumnName = '_key_AnalysisId_KOKRS_PRCTR'
        toTableName = '_key_AnalysisId_KOKRS_PRCTR'
        toColumnName = '_key_AnalysisId_KOKRS_PRCTR'
        CrossFilteringBehavior = 'BothDirections'
    },
    @{
        fromTableName = 'Dim - HRC keys'
        fromColumnName = '_key_AnalysisId_HRCGroup'
        toTableName = 'Dim - HRC descriptions'
        toColumnName = '_key_AnalysisId_HRCGroup'
        CrossFilteringBehavior = 'BothDirections'
    }
)

New-SSASRelationships -db $db -relationshipDefinitionList $relationshipDefinitionList
```

### Export Power BI desktop model metadata (tables, columns, load script, relationships, measures) to an excel file
See [this file](Example%20-%20powerbi%20desktop%20connect%20to%20local%20-%20extract%20model%20json%20and%20metadata.ps1).
