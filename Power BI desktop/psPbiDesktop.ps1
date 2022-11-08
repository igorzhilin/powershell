<#
    This script contains the functions to:
        - connect to PBI desktop data model and allows to access the metadata, i.e. tables, columns, relationships
        - edit PBI data model in batch (e.g. rename columns, hide columns, update measures...)
    
    Basically, access and change PBI objects that are SSAS-based.

    This script does NOT allow to change the visuals.

    Process:
        - open pbix file in PBI desktop, wait for it to load
        - run commands
        
    How it works:
        - find SSAS process msmdsrv.exe and its network port
        - get SSAS's parent process command line - it should be PBIDesktop.exe with the path to pbix file
        - connect to the SSAS instance - then you can use Microsoft.AnalysisServices.Tabular

    Dependencies:
        sqlserver 22.0.20 or higher
        joinmodule

    The beautiful ascii labels are created with figletcmd wrapped in an autohotkey script.
#>
cls

Import-Module sqlserver -MinimumVersion "22.0.20"
Import-Module JoinModule -DisableNameChecking # -DisableNameChecking to suppress unapproved verbs warning

<#
            _     _                 _         _     _ 
           | |   | |               | |       | |   (_)
  __ _  ___| |_  | | ___   ___ __ _| |  _ __ | |__  _ 
 / _` |/ _ \ __| | |/ _ \ / __/ _` | | | '_ \| '_ \| |
| (_| |  __/ |_  | | (_) | (_| (_| | | | |_) | |_) | |
 \__, |\___|\__| |_|\___/ \___\__,_|_| | .__/|_.__/|_|
  __/ |                                | |            
 |___/                                 |_|
#>

Function Get-LocalPBIDesktopInstances {
    [CmdletBinding()]
    param(
        [Parameter (Mandatory = $false)]$pbixFilePathRegex
    )

    # look for the local SSAS process
    $processNamesToLookFor = @"
msmdsrv.exe
"@ -Split([Environment]::NewLine) | Where-Object {$_} 

    Write-Host "Getting PIDs: $($processNamesToLookFor -join ', ')" -ForegroundColor Gray

    # find the source app PIDs
    $processes = @()
    ForEach($processName in $processNamesToLookFor) {
        $processInfo = Get-WMIObject -Class win32_process -Filter "Name = '$($processName)'"
        $processes += $processInfo | Select-Object Name, ProcessId, ParentProcessId, CommandLine
    }

    # get parent process details as well
    $processesWithDetails = @()
    ForEach($process in $processes) {
        $parentProcessInfo = Get-WMIObject -Class win32_process -Filter "ProcessId = '$($process.ParentProcessId)'"

        $processesWithDetails += $process | Select-Object *, @{N="Parent";E={$parentProcessInfo}}
    }

    # when there is a pbix file opened in pbi desktop, the command line of pbidesktop.exe would be: "PBIDesktop.exe path" "pbix file path"
    $processesWithDetails = $processesWithDetails |`
        Select-Object *, `
            @{N="ParentName";E={$_.Parent.Name}}, `
            @{N="ParentCommandLine";E={$_.Parent.CommandLine}}, `
            @{N="PbixFilePath";E={[regex]::Replace($_.Parent.CommandLine, '"(.*)" "(.*)"', '$2')}}
    
    # get the port where SSAS is running
    Write-Host "Getting connections for PIDs: $($processesWithDetails.ProcessId -join ', ')" -ForegroundColor Gray

    $connections = Get-NetTCPConnection -OwningProcess $processesWithDetails.ProcessId |` 
        Select-Object `
            OwningProcess, 
            LocalPort -Unique
    
    # add process data to connections
    $connectionsWithProcessDetails = Join-Object -LeftObject $connections -RightObject $processesWithDetails -OnExpression {$Left.OwningProcess -eq $Right.ProcessId}

    # if file path pattern supplied, get the matching process
    If($pbixFilePathRegex) {
        $connectionsWithProcessDetails = $connectionsWithProcessDetails | Where-Object PbixFilePath -Match $pbixFilePathRegex
    }

    $connectionsWithProcessDetails
}

<#
            _     _        _           _              _____  ____      
           | |   | |      | |         | |            |  __ \|  _ \     
  __ _  ___| |_  | |_ __ _| |__  _   _| | __ _ _ __  | |  | | |_) |___ 
 / _` |/ _ \ __| | __/ _` | '_ \| | | | |/ _` | '__| | |  | |  _ </ __|
| (_| |  __/ |_  | || (_| | |_) | |_| | | (_| | |    | |__| | |_) \__ \
 \__, |\___|\__|  \__\__,_|_.__/ \__,_|_|\__,_|_|    |_____/|____/|___/
  __/ |                                                                
 |___/
#>

Function Get-SSASTabularDBs {
    param(
        $SSASTabularHostPort
    )

    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.AnalysisServices.Tabular") | Out-Null
    $tabularServer = New-Object Microsoft.AnalysisServices.Tabular.Server
    
    Write-Host "Attempting to connect to local SSAS Tabular: $SSASTabularHostPort" -ForegroundColor Gray
    
    $dbs = $null
    
    $tabularServer.Connect("Data source=$SSASTabularHostPort")
    
    If($tabularServer.Connected) {
        Write-Host "Success" -ForegroundColor Green
        $dbs = $tabularServer.Databases
    }
    else {
        Write-Host "Failed" -ForegroundColor Red

    }

    $dbs
}

<#
########  ########    ###    ########  
##     ## ##         ## ##   ##     ## 
##     ## ##        ##   ##  ##     ## 
########  ######   ##     ## ##     ## 
##   ##   ##       ######### ##     ## 
##    ##  ##       ##     ## ##     ## 
##     ## ######## ##     ## ########
#>

<#
            _         _ _                _     
           | |       | | |              | |    
  __ _  ___| |_    __| | |__    _ __ ___| |___ 
 / _` |/ _ \ __|  / _` | '_ \  | '__/ _ \ / __|
| (_| |  __/ |_  | (_| | |_) | | | |  __/ \__ \
 \__, |\___|\__|  \__,_|_.__/  |_|  \___|_|___/
  __/ |                                        
 |___/
#>

Function Get-SSASTabularModelRelationships {
    param(
        $SSASTabularDb
    )

    Write-Host "Opened SSAS Tabular DB: $($SSASTabularDb.Name)" -ForegroundColor Gray
    Write-Host "Collecting relationships"
    
    $relationships = @()
    ForEach($relationship in $SSASTabularDb.Model.Relationships) {
        $relationshipInfo = New-Object PSCustomObject
    
        $relationshipInfo | Add-Member -MemberType NoteProperty -Name "FromCardinality" -Value $relationship.FromCardinality
        $relationshipInfo | Add-Member -MemberType NoteProperty -Name "FromTableName" -Value $relationship.FromColumn.Table.Name
        $relationshipInfo | Add-Member -MemberType NoteProperty -Name "FromColumnName"  -Value $relationship.FromColumn.Name

        $relationshipInfo | Add-Member -MemberType NoteProperty -Name "ToCardinality" -Value  $relationship.ToCardinality
        $relationshipInfo | Add-Member -MemberType NoteProperty -Name "ToTableName" -Value  $relationship.ToColumn.Table.Name
        $relationshipInfo | Add-Member -MemberType NoteProperty -Name "ToColumnName" -Value  $relationship.ToColumn.Name

        $relationships += $relationshipInfo
    }

    $relationships
}

<#
            _                         _      _        _  _____  ____  _   _ 
           | |                       | |    | |      | |/ ____|/ __ \| \ | |
  __ _  ___| |_   _ __ ___   ___   __| | ___| |      | | (___ | |  | |  \| |
 / _` |/ _ \ __| | '_ ` _ \ / _ \ / _` |/ _ \ |  _   | |\___ \| |  | | . ` |
| (_| |  __/ |_  | | | | | | (_) | (_| |  __/ | | |__| |____) | |__| | |\  |
 \__, |\___|\__| |_| |_| |_|\___/ \__,_|\___|_|  \____/|_____/ \____/|_| \_|
  __/ |                                                                     
 |___/
#>

Function Get-SSASTabularModelJSON {
    param(
        $SSASTabularDb
    )

    Write-Host "Opened SSAS Tabular DB: $($SSASTabularDb.Name)" -ForegroundColor Gray
    Write-Host "Scripting model as JSON"
    
    $scriptJson = [Microsoft.AnalysisServices.Tabular.JsonScripter]::ScriptCreate($SSASTabularDb)
    
    $scriptJson
}

<#
            _     _        _     _              __ _      _     _     
           | |   | |      | |   | |            / _(_)    | |   | |    
  __ _  ___| |_  | |_ __ _| |__ | | ___  ___  | |_ _  ___| | __| |___ 
 / _` |/ _ \ __| | __/ _` | '_ \| |/ _ \/ __| |  _| |/ _ \ |/ _` / __|
| (_| |  __/ |_  | || (_| | |_) | |  __/\__ \ | | | |  __/ | (_| \__ \
 \__, |\___|\__|  \__\__,_|_.__/|_|\___||___/ |_| |_|\___|_|\__,_|___/
  __/ |                                                               
 |___/
#>

Function Get-SSASTabularTablesFields {
    param(
        $SSASTabularDb
    )
    Write-Host "Opened SSAS Tabular DB: $($SSASTabularDb.Name)" -ForegroundColor Gray
    Write-Host "Collecting tables and fields"

    $tablesFields = @()
    ForEach($table in $SSASTabularDb.Model.Tables) {
        
        $tablesFields += $table.Columns | Select-Object @{N="TableName";E={$table.Name}}, Name, Type, Expression, DataType, IsHidden, IsUnique, IsKey
    }

    $tablesFields

}

<#
            _                                                  
           | |                                                 
  __ _  ___| |_   _ __ ___   ___  __ _ ___ _   _ _ __ ___  ___ 
 / _` |/ _ \ __| | '_ ` _ \ / _ \/ _` / __| | | | '__/ _ \/ __|
| (_| |  __/ |_  | | | | | |  __/ (_| \__ \ |_| | | |  __/\__ \
 \__, |\___|\__| |_| |_| |_|\___|\__,_|___/\__,_|_|  \___||___/
  __/ |                                                        
 |___/
#>

Function Get-SSASTabularMeasures {
    param(
        $SSASTabularDb
    )
    Write-Host "Opened SSAS Tabular DB: $($SSASTabularDb.Name)" -ForegroundColor Gray
    Write-Host "Collecting measures"

    $measures = @()
    ForEach($measure in $SSASTabularDb.Model.Tables.Measures) {
        
        $measures += $measure | Select-Object Name, Expression, FormatString, IsHidden, @{N="TableName";E={$_.Table.Name}}
    }

    $measures
}

<#
            _      _____ _____          _____                  _        
           | |    / ____/ ____|  /\    / ____|                | |       
  __ _  ___| |_  | (___| (___   /  \  | (___    _ __ ___   ___| |_ __ _ 
 / _` |/ _ \ __|  \___ \\___ \ / /\ \  \___ \  | '_ ` _ \ / _ \ __/ _` |
| (_| |  __/ |_   ____) |___) / ____ \ ____) | | | | | | |  __/ || (_| |
 \__, |\___|\__| |_____/_____/_/    \_\_____/  |_| |_| |_|\___|\__\__,_|
  __/ |                                                                 
 |___/
#>

Function Get-SSASDBMetadata {
    param(
        $SSASTabularDb
    )
    Write-Host "Opened SSAS Tabular DB: $($SSASTabularDb.Name)" -ForegroundColor Gray
    Write-Host "Collecting DB metadata"

    $dbMetadata = $SSASTabularDb | Select-Object Name, Server, EstimatedSize, State, CreatedTimestamp, LastUpdate, LastSchemaUpdate, LastProcessed
    $dbMetadata
}

<#
            _     __  __     _______          __   __                _       _       
           | |   |  \/  |   / /  __ \   /\    \ \ / /               (_)     | |      
  __ _  ___| |_  | \  / |  / /| |  | | /  \    \ V /   ___  ___ _ __ _ _ __ | |_ ___ 
 / _` |/ _ \ __| | |\/| | / / | |  | |/ /\ \    > <   / __|/ __| '__| | '_ \| __/ __|
| (_| |  __/ |_  | |  | |/ /  | |__| / ____ \  / . \  \__ \ (__| |  | | |_) | |_\__ \
 \__, |\___|\__| |_|  |_/_/   |_____/_/    \_\/_/ \_\ |___/\___|_|  |_| .__/ \__|___/
  __/ |                                                               | |            
 |___/                                                                |_|
#>

Function Get-SSASDBTableScripts {
    param(
        $SSASTabularDb
    )

    Write-Host "Opened SSAS Tabular DB: $($SSASTabularDb.Name)" -ForegroundColor Gray
    Write-Host "Collecting table load M/DAX scripts"

    $tableScripts = @()

    ForEach($table in $db.Model.Tables) {
        $tableName = $table.Name
        $tableIsHidden = $table.IsHidden
        
        ForEach($partition in $table.Partitions) {
            $sourceType = $partition.SourceType
            $expression = $partition.Source.Expression
            $tableScripts += 1 | Select-Object @{N="TableName";E={$tableName}}, @{N="IsHidden";E={$tableIsHidden}}, @{N="SourceType";E={$sourceType}}, @{N="Expression";E={$expression}}
        }
    }

    $tableScripts
}

<#
######## ########  #### ######## 
##       ##     ##  ##     ##    
##       ##     ##  ##     ##    
######   ##     ##  ##     ##    
##       ##     ##  ##     ##    
##       ##     ##  ##     ##    
######## ########  ####    ##
#>

<#
          _               _                             
         | |             | |                            
 ___  ___| |_    ___ ___ | |  _ __  _ __ ___  _ __  ___ 
/ __|/ _ \ __|  / __/ _ \| | | '_ \| '__/ _ \| '_ \/ __|
\__ \  __/ |_  | (_| (_) | | | |_) | | | (_) | |_) \__ \
|___/\___|\__|  \___\___/|_| | .__/|_|  \___/| .__/|___/
                             | |             | |        
                             |_|             |_|
#>

Function Set-SSASColumnProperties {

    param(
        $db,
        $tableNameRegex,
        $columnNameRegex,
        $columnProperties
    )

    # Change column properties
    #$tableNameRegex = '.*dim.*product.*'
    #$columnNameRegex = '^prod.*surr.*'

    Write-Host "Looking for tables & columns: "  -NoNewline
    Write-Host "'$($tableNameRegex)'[$($columnNameRegex)]" -ForegroundColor Yellow 

    <#
    In, $columnProperties, the key names and value datatypes should match the $column property names and their datatypes respectively
    The properties should be settable. $column | Get-Member -MemberType Properties | Select-Object Name, Definition | Where-Object Definition -Match '.*\bset\b.*'
    Which are:
        Alignment            Microsoft.AnalysisServices.Tabular.Alignment Alignment {get;set;}                      
        AlternateOf          Microsoft.AnalysisServices.Tabular.AlternateOf AlternateOf {get;set;}                  
        DataCategory         string DataCategory {get;set;}                                                         
        DataType             Microsoft.AnalysisServices.Tabular.DataType DataType {get;set;}                        
        Description          string Description {get;set;}                                                          
        DisplayFolder        string DisplayFolder {get;set;}                                                        
        DisplayOrdinal       int DisplayOrdinal {get;set;}                                                          
        EncodingHint         Microsoft.AnalysisServices.Tabular.EncodingHintType EncodingHint {get;set;}            
        FormatString         string FormatString {get;set;}                                                         
        IsAvailableInMDX     bool IsAvailableInMDX {get;set;}                                                       
        IsDataTypeInferred   bool IsDataTypeInferred {get;set;}                                                     
        IsDefaultImage       bool IsDefaultImage {get;set;}                                                         
        IsDefaultLabel       bool IsDefaultLabel {get;set;}                                                         
        IsHidden             bool IsHidden {get;set;}                                                               
        IsKey                bool IsKey {get;set;}                                                                  
        IsNullable           bool IsNullable {get;set;}                                                             
        IsUnique             bool IsUnique {get;set;}                                                               
        KeepUniqueRows       bool KeepUniqueRows {get;set;}                                                         
        LineageTag           string LineageTag {get;set;}                                                           
        Name                 string Name {get;set;}                                                                 
        RelatedColumnDetails Microsoft.AnalysisServices.Tabular.RelatedColumnDetails RelatedColumnDetails {get;set;}
        SortByColumn         Microsoft.AnalysisServices.Tabular.Column SortByColumn {get;set;}                      
        SourceColumn         string SourceColumn {get;set;}                                                         
        SourceLineageTag     string SourceLineageTag {get;set;}                                                     
        SourceProviderType   string SourceProviderType {get;set;}                                                   
        SummarizeBy          Microsoft.AnalysisServices.Tabular.AggregateFunction SummarizeBy {get;set;}            
        TableDetailPosition  int TableDetailPosition {get;set;}   
    
    DataType:
        [Microsoft.AnalysisServices.Tabular.DataType]::Automatic
        [Microsoft.AnalysisServices.Tabular.DataType]::Decimal
        [Microsoft.AnalysisServices.Tabular.DataType]::Int64
        [Microsoft.AnalysisServices.Tabular.DataType]::Binary
        [Microsoft.AnalysisServices.Tabular.DataType]::String
        [Microsoft.AnalysisServices.Tabular.DataType]::DateTime                                  

    FormatString: commonly used
        #,0 - no decimals, with thousand separator
        0%;-0%;0% - percentages, no decimals

    #>

    $tables = $db.Model.Tables | Where-Object Name -Match $tableNameRegex 
    ForEach($table in $tables) {
        $columns =  $table.Columns | Where-Object Name -Match $columnNameRegex
        ForEach($column in $columns) {
            Write-Host "'$($table.Name)'[$($column.Name)]" -ForegroundColor Yellow
            ForEach($propName in $columnProperties.Keys) {
                Write-Host "`t" -NoNewline
                Write-Host "$($propName)" -ForegroundColor Magenta -NoNewline
                Write-Host " = " -ForegroundColor Gray -NoNewline
                Write-Host "$($columnProperties[$propName])" -ForegroundColor Cyan

                $column.$propName = $columnProperties[$propName]
            }
        }
    }

    Write-Host "Saving model changes" -ForegroundColor Gray
    $db.Model.SaveChanges() | Out-Null
}

<#
                                                 _                           
                                                | |                          
 _ __ ___ _ __   __ _ _ __ ___   ___    ___ ___ | |_   _ _ __ ___  _ __  ___ 
| '__/ _ \ '_ \ / _` | '_ ` _ \ / _ \  / __/ _ \| | | | | '_ ` _ \| '_ \/ __|
| | |  __/ | | | (_| | | | | | |  __/ | (_| (_) | | |_| | | | | | | | | \__ \
|_|  \___|_| |_|\__,_|_| |_| |_|\___|  \___\___/|_|\__,_|_| |_| |_|_| |_|___/
#>

Function Set-SSASColumnNamesFindReplace {
    param(
        $db, 
        $tableNameRegex, 
        $findReplaceList 
        <#
        expects array of hashtables with keys Find Replace
        by default, regex matching is case-insensitive
        for case-sensitive regex, add (?-i) at the beginning of the string
        $findReplaceList = @(
            @{Find = '(?-i)\(DesC\)'; Replace = '(desc.)'}
        )
        #>
    )

    ForEach($findReplace in $findReplaceList) {
        $find = $findReplace.Find
        $replace = $findReplace.Replace

        $tables = $db.Model.Tables | Where-Object Name -Match $tableNameRegex 
        ForEach($table in $tables) {
            $columns =  $table.Columns | Where-Object Name -Match $find
            ForEach($column in $columns) {
                Write-Host "'$($table.Name)'[$($column.Name)]" -ForegroundColor Yellow -NoNewline
                Write-Host " -> " -ForegroundColor White -NoNewline
                $newName = [regex]::Replace($column.Name, $find, $replace)
                Write-Host "'$($table.Name)'[$newName]" -ForegroundColor Magenta

                $column.Name = $newName
            }
        }
    }

    Write-Host "Saving model changes" -ForegroundColor Gray
    $db.Model.SaveChanges() | Out-Null
}

<#
 _ __ ___   ___  __ _ ___ _   _ _ __ ___   _ __  _ __ ___  _ __  ___ 
| '_ ` _ \ / _ \/ _` / __| | | | '__/ _ \ | '_ \| '__/ _ \| '_ \/ __|
| | | | | |  __/ (_| \__ \ |_| | | |  __/ | |_) | | | (_) | |_) \__ \
|_| |_| |_|\___|\__,_|___/\__,_|_|  \___| | .__/|_|  \___/| .__/|___/
                                          | |             | |        
                                          |_|             |_|
#>

Function Set-SSASMeasureProperties {

    param(
        $db,
        $tableNameRegex,
        $measureNameRegex,
        $measureProperties
    )

    <#
    $db.Model.Tables['_Measures - entity'].Measures[0] | Get-Member -MemberType Properties | Select-Object Name, Definition | Where-Object Definition -Match '.*\bset\b.*'
        DataCategory           string DataCategory {get;set;}                                                             
        Description            string Description {get;set;}                                                              
        DetailRowsDefinition   Microsoft.AnalysisServices.Tabular.DetailRowsDefinition DetailRowsDefinition {get;set;}    
        DisplayFolder          string DisplayFolder {get;set;}                                                            
        Expression             string Expression {get;set;}                                                               
        FormatString           string FormatString {get;set;}                                                             
        FormatStringDefinition Microsoft.AnalysisServices.Tabular.FormatStringDefinition FormatStringDefinition {get;set;}
        IsHidden               bool IsHidden {get;set;}                                                                   
        IsSimpleMeasure        bool IsSimpleMeasure {get;set;}                                                            
        KPI                    Microsoft.AnalysisServices.Tabular.KPI KPI {get;set;}                                      
        LineageTag             string LineageTag {get;set;}                                                               
        Name                   string Name {get;set;}                                                                     
        SourceLineageTag       string SourceLineageTag {get;set;}
    
    #>

    Write-Host "Looking for tables & measures: "  -NoNewline
    Write-Host "'$($tableNameRegex)'[$($measureNameRegex)]" -ForegroundColor Yellow 


    $tables = $db.Model.Tables | Where-Object Name -Match $tableNameRegex 
    ForEach($table in $tables) {
        $measures =  $table.Measures | Where-Object Name -Match $measureNameRegex
        ForEach($measure in $measures) {
            Write-Host "'$($table.Name)'[$($measure.Name)]" -ForegroundColor Yellow
            
            ForEach($propName in $measureProperties.Keys) {
                Write-Host "`t" -NoNewline
                Write-Host "$($propName)" -ForegroundColor Magenta -NoNewline
                Write-Host " = " -ForegroundColor Gray -NoNewline
                Write-Host "$($measureProperties[$propName])" -ForegroundColor Cyan

                $measure.$propName = $measureProperties[$propName]
            }
            
        }
    }

    Write-Host "Saving model changes" -ForegroundColor Gray
    $db.Model.SaveChanges() | Out-Null
}

<#
 _ __ ___ _ __   __ _ _ __ ___   ___   _ __ ___   ___  __ _ ___ _   _ _ __ ___  ___ 
| '__/ _ \ '_ \ / _` | '_ ` _ \ / _ \ | '_ ` _ \ / _ \/ _` / __| | | | '__/ _ \/ __|
| | |  __/ | | | (_| | | | | | |  __/ | | | | | |  __/ (_| \__ \ |_| | | |  __/\__ \
|_|  \___|_| |_|\__,_|_| |_| |_|\___| |_| |_| |_|\___|\__,_|___/\__,_|_|  \___||___/
#>

Function Set-SSASMeasureNamesFindReplace {
    <#
    !!! this function generally works but it needs more polishing
    ToDo:    
        - update references to renamed measure in calc columns
        - double-check that it really works as expected
    #>

    param(
        $db, 
        $tableNameRegex, 
        $findReplaceRegexList
        <#
        expects array of hashtables with keys Find Replace
        by default, regex matching is case-insensitive
        for case-sensitive regex, add (?-i) at the beginning of the string
        $findReplaceList = @(
            @{Find = '(?-i)\(DesC\)'; Replace = '(desc.)'}
        )
        #>
    )

    # get tables in scope
    $tables = $db.Model.Tables | Where-Object Name -Match $tableNameRegex
    
    ForEach($findReplace in $findReplaceRegexList) {

        $find = $findReplace.Find
        $replace = $findReplace.Replace

        # get measures fulfilling the find criteria
        
        $measures = @()

        ForEach($table in $tables) {
            $measures += $table.Measures | Where-Object Name -Match $find    
        }

        # save old measure names - we will need them to update expressions depending on old names
        # rename measures
        
        Write-Host "Renaming measures"
        $oldNewMeasureNames = @()

        ForEach($measure in $measures) {
            $oldName = $measure.Name
            $newName = [regex]::Replace($oldName, $find, $replace)

            Write-Host "'$($table.Name)'[$oldName]" -ForegroundColor Yellow -NoNewline
            Write-Host " -> " -ForegroundColor White -NoNewline
            Write-Host "'$($table.Name)'[$newName]" -ForegroundColor Magenta

            $measure.Name = $newName

            $oldNewMeasureNames += 1 | Select-Object @{N="OldName";E={$oldName}}, @{N="NewName";E={$newName}}
        }

        Write-Host "Saving model changes" -ForegroundColor Gray
        $db.Model.SaveChanges() | Out-Null


        # update expressions depending on the old names

        Write-Host "Collecting measures depening on old measure names"

        $dependentMeasures = @()
        ForEach($table in $tables) {
            ForEach($oldNew in $oldNewMeasureNames) {
                $oldName = $oldNew.OldName
                $oldNameCLike = '*{0}*' -f $oldName
                
                $newName = $oldNew.NewName

                $dependentMeasures += $table.Measures | Where-Object Expression -CLike $oldNameCLike
            }
        }

        Write-Host "Updating the expressions of dependent measures with new measure names"

        ForEach($dependentMeasure in $dependentMeasures) {
            
            Write-Host "Updating expression " -ForegroundColor Gray -NoNewline
            Write-Host "'$($dependentMeasure.Table.Name)'[$($dependentMeasure.Name)]" -ForegroundColor Yellow

            ForEach($oldNew in $oldNewMeasureNames) {
                $oldExpression = $dependentMeasure.Expression
                $oldName = '[{0}]' -f $oldNew.OldName
                $newName = '[{0}]' -f $oldNew.NewName

                # [regex]::escape is needed because -creplace would otherwise treat this as regex
                $newExpression = $oldExpression -creplace [regex]::escape($oldName), $newName
                
                $dependentMeasure.Expression = $newExpression

                Write-Host $oldExpression -ForegroundColor Gray
                Write-Host $newExpression -ForegroundColor Magenta
            }
        }

        Write-Host "Saving model changes" -ForegroundColor Gray
        $db.Model.SaveChanges() | Out-Null
    }
}

<#
                     _                                                      
                    | |                                                     
  ___ _ __ ___  __ _| |_ ___   _ __ ___   ___  __ _ ___ _   _ _ __ ___  ___ 
 / __| '__/ _ \/ _` | __/ _ \ | '_ ` _ \ / _ \/ _` / __| | | | '__/ _ \/ __|
| (__| | |  __/ (_| | ||  __/ | | | | | |  __/ (_| \__ \ |_| | | |  __/\__ \
 \___|_|  \___|\__,_|\__\___| |_| |_| |_|\___|\__,_|___/\__,_|_|  \___||___/
#>

Function New-SSASMeasures {
    param(
        $db, 
        $tableName, 
        $measureDefinitionList,
        [switch]$replace

        <#
        expects array of hashtables:

        $measureDefinitionList = @(
            @{
                Name = "TestMeasure1"
                Expression = "COUNTROWS('Analysis Details')"
                FormatString = "#,0"
            },
            @{
                Name = "TestMeasure2"
                Expression = "COUNTROWS('Analysis Details') + 999"
                FormatString = "#,0"
            },
            @{
                Name = "TestMeasure3"
                Expression = "DIVIDE([TestMeasure1],[TestMeasure2])"
                FormatString = "#.#%"
            }
        )
        #>
    )

    # get tables in scope
    $table = $db.Model.Tables | Where-Object Name -eq $tableName
    
    If(-not $table) {
        Write-Host "Table not found: " -ForegroundColor Gray -NoNewline
        Write-Host "'$tableName'" -ForegroundColor Yellow

        return
    }

    Write-Host "Adding measures to table " -ForegroundColor Gray -NoNewline
    Write-Host "'$tableName'" -ForegroundColor Yellow

    $measuresToAdd = @()
    ForEach($measureDefinition in $measureDefinitionList) {
        $newMeasure = New-Object Microsoft.AnalysisServices.Tabular.Measure
        $newMeasure.Name = $measureDefinition.Name
        $newMeasure.Expression = $measureDefinition.Expression
        $newMeasure.FormatString = $measureDefinition.FormatString
        
        $measuresToAdd += $newMeasure
    }

    # if -replace switch is on, then delete the existing measure
    If($replace) {
        Write-Host "-replace " -ForegroundColor Cyan -NoNewline
        Write-Host "switch is ACTIVE, existing measures WILL BE REPLACED" -ForegroundColor Gray

        ForEach($measure in $table.Measures | Where-Object Name -in $measuresToAdd.Name) {
            Write-Host "Deleting existing measure " -ForegroundColor Gray -NoNewline
            Write-Host "'$tableName'" -ForegroundColor Yellow -NoNewline
            Write-Host "[$($measure.Name)]" -ForegroundColor Magenta
            $table.Measures.Remove($measure.Name)
        }
    }

    # if -replace switch is NOT on, then do not deploy the new measure
    If(-not $replace) {        
        Write-Host "-replace " -ForegroundColor Cyan -NoNewline
        Write-Host "switch NOT active, existing measures will NOT be replaced" -ForegroundColor Gray
        $filteredMeasuresToAdd = @()
        ForEach($measure in $measuresToAdd) {
            If($measure.Name -in $Table.Measures.Name) {
                Write-Host "Measure already exists and will NOT be replaced " -ForegroundColor Gray -NoNewline
                Write-Host "'$tableName'" -ForegroundColor Yellow -NoNewline
                Write-Host "[$($measure.Name)]" -ForegroundColor Magenta
                
            } else {
                $filteredMeasuresToAdd += $measure
            }
        }
        $measuresToAdd = $filteredMeasuresToAdd
    }

    If(-not $measuresToAdd) {
        Write-Host "No measures to add. Exiting " -ForegroundColor Gray
        return
    }
    

    ForEach($measure in $measuresToAdd) {
        Write-Host "Adding " -ForegroundColor Gray -NoNewline
        Write-Host "'$tableName'" -ForegroundColor Yellow -NoNewline
        Write-Host "[$($measure.Name)]" -ForegroundColor Magenta
        Write-Host "$($measure.Expression)" -ForegroundColor White
        $table.Measures.Add($measure)
    }

    Write-Host "Saving model changes" -ForegroundColor Gray
    $db.Model.SaveChanges() | Out-Null
}

<#
          _     _        _     _                                  
         | |   | |      | |   | |                                 
 ___  ___| |_  | |_ __ _| |__ | | ___   _ __  _ __ ___  _ __  ___ 
/ __|/ _ \ __| | __/ _` | '_ \| |/ _ \ | '_ \| '__/ _ \| '_ \/ __|
\__ \  __/ |_  | || (_| | |_) | |  __/ | |_) | | | (_) | |_) \__ \
|___/\___|\__|  \__\__,_|_.__/|_|\___| | .__/|_|  \___/| .__/|___/
                                       | |             | |        
                                       |_|             |_|
#>

Function Set-SSASTableProperties {

    param(
        $db,
        $tableNameRegex,
        $tableProperties
    )

    # Change column properties

    Write-Host "Looking for tables "  -NoNewline
    Write-Host "'$($tableNameRegex)'" -ForegroundColor Yellow 

    <#
    $db.Model.Tables['_Measures - entity'] | Get-Member -MemberType Properties | Select-Object Name, Definition | Where-Object Definition -Match '.*\bset\b.*'
    
    Settable properties are:
        AlternateSourcePrecedence   int AlternateSourcePrecedence {get;set;}                                                      
        CalculationGroup            Microsoft.AnalysisServices.Tabular.CalculationGroup CalculationGroup {get;set;}               
        DataCategory                string DataCategory {get;set;}                                                                
        DefaultDetailRowsDefinition Microsoft.AnalysisServices.Tabular.DetailRowsDefinition DefaultDetailRowsDefinition {get;set;}
        Description                 string Description {get;set;}                                                                 
        ExcludeFromModelRefresh     bool ExcludeFromModelRefresh {get;set;}                                                       
        IsHidden                    bool IsHidden {get;set;}                                                                      
        IsPrivate                   bool IsPrivate {get;set;}                                                                     
        LineageTag                  string LineageTag {get;set;}                                                                  
        Name                        string Name {get;set;}                                                                        
        RefreshPolicy               Microsoft.AnalysisServices.Tabular.RefreshPolicy RefreshPolicy {get;set;}                     
        ShowAsVariationsOnly        bool ShowAsVariationsOnly {get;set;}                                                          
        SourceLineageTag            string SourceLineageTag {get;set;}                                                            
        SystemManaged               bool SystemManaged {get;set;}  
    #>

    $tables = $db.Model.Tables | Where-Object Name -Match $tableNameRegex 
    ForEach($table in $tables) {
        Write-Host "'$($table.Name)'" -ForegroundColor Yellow
        ForEach($propName in $tableProperties.Keys) {
            Write-Host "`t" -NoNewline
            Write-Host "$($propName)" -ForegroundColor Magenta -NoNewline
            Write-Host " = " -ForegroundColor Gray -NoNewline
            Write-Host "$($tableProperties[$propName])" -ForegroundColor Cyan

            $table.$propName = $tableProperties[$propName]
        }
    }

    Write-Host "Saving model changes" -ForegroundColor Gray
    $db.Model.SaveChanges() | Out-Null
}

<#
                     _                  _       _   _                 _     _           
                    | |                | |     | | (_)               | |   (_)          
  ___ _ __ ___  __ _| |_ ___   _ __ ___| | __ _| |_ _  ___  _ __  ___| |__  _ _ __  ___ 
 / __| '__/ _ \/ _` | __/ _ \ | '__/ _ \ |/ _` | __| |/ _ \| '_ \/ __| '_ \| | '_ \/ __|
| (__| | |  __/ (_| | ||  __/ | | |  __/ | (_| | |_| | (_) | | | \__ \ | | | | |_) \__ \
 \___|_|  \___|\__,_|\__\___| |_|  \___|_|\__,_|\__|_|\___/|_| |_|___/_| |_|_| .__/|___/
                                                                             | |        
                                                                             |_|
#>

Function New-SSASRelationships {
    param(
        $db,
        $relationshipDefinitionList
    )

    <#
        From must be MANY
        To must be ONE
        
        CrossFilteringBehavior:
        https://learn.microsoft.com/en-us/dotnet/api/microsoft.analysisservices.tabular.crossfilteringbehavior?view=analysisservices-dotnet           
            Automatic 	3 	
            The engine will analyze the relationships and choose one of the behaviors by using heuristics.
            BothDirections 	2 	
            Filters on either end of the relationship will automatically filter the other table.
            OneDirection 	1 	
            The rows selected in the 'To' end of the relationship will automatically filter scans of the table in the 'From' end of the relationship.
    #>
    $crossFilteringBehaviorEnum = @{
        'OneDirection' = 1, ' -> '
        'BothDirections' = 2, ' <-> '
        'Automatic' = 3, ' <~> '
        'default' = 2, ' <-> '
    }


    ForEach($relationshipDefinition in $relationshipDefinitionList) {
        $fromTableName              = $relationshipDefinition.fromTableName
        $fromColumnName             = $relationshipDefinition.fromColumnName
        $toTableName                = $relationshipDefinition.toTableName
        $toColumnName               = $relationshipDefinition.toColumnName
        $crossFilteringBehaviorText = $relationshipDefinition.crossFilteringBehavior

        try {
            $crossFilteringBehavior = $crossFilteringBehaviorEnum[$crossFilteringBehaviorText][0]
            $crossFilteringArrow    = $crossFilteringBehaviorEnum[$crossFilteringBehaviorText][1]
        }
        catch {
            Write-Host "Did not find cross-filtering enum " -ForegroundColor Gray -NoNewline
            Write-Host "$crossFilteringBehaviorText " -ForegroundColor Yellow 

            $crossFilteringBehaviorText = 'default'
            $crossFilteringBehavior = $crossFilteringBehaviorEnum[$crossFilteringBehaviorText][0]
            $crossFilteringArrow    = $crossFilteringBehaviorEnum[$crossFilteringBehaviorText][1]

            Write-Host "Default cross-filtering will be used: " -ForegroundColor Gray -NoNewline
            Write-Host "$crossFilteringArrow" -ForegroundColor Yellow 
        }
        
        Write-Host "'$fromTableName'" -ForegroundColor Magenta -NoNewline
        Write-Host "[$fromColumnName]" -ForegroundColor Cyan -NoNewline
        Write-Host $crossFilteringArrow -NoNewline
        Write-Host "'$toTableName'" -ForegroundColor Magenta -NoNewline
        Write-Host "[$toColumnName]" -ForegroundColor Cyan 

        $fromTable = $db.Model.Tables | Where-Object Name -eq $fromTableName
        If(-not $fromTable) {
            Write-Host "Table not found: " -ForegroundColor Red -NoNewline
            Write-Host "'$fromTableName'" -ForegroundColor Magenta 
            Write-Host "Skipping this relationship" -ForegroundColor Gray 
            continue
        }
        
       $fromColumn = $fromTable.Columns | Where-Object Name -eq $fromColumnName
        If(-not $fromColumn) {
            Write-Host "Column not found: " -ForegroundColor Red -NoNewline
            Write-Host "'$fromTableName'" -ForegroundColor Magenta -NoNewline
            Write-Host "[$fromColumnName]" -ForegroundColor Cyan
            Write-Host "Skipping this relationship" -ForegroundColor Gray 
            continue
        }

        $toTable = $db.Model.Tables | Where-Object Name -eq $toTableName
        If(-not $toTable) {
            Write-Host "Table not found: " -ForegroundColor Red -NoNewline
            Write-Host "'$toTableName'" -ForegroundColor Magenta 
            Write-Host "Skipping this relationship" -ForegroundColor Gray 
            continue
        }
        
       $toColumn = $toTable.Columns | Where-Object Name -eq $toColumnName
        If(-not $toColumn) {
            Write-Host "Column not found: " -ForegroundColor Red -NoNewline
            Write-Host "'$toTableName'" -ForegroundColor Magenta -NoNewline
            Write-Host "[$toColumnName]" -ForegroundColor Cyan
            Write-Host "Skipping this relationship" -ForegroundColor Gray 
            continue
        }        
        
        $relationshipExists = $false

        # check if rel exists
        ForEach($relationship in $db.Model.Relationships) {
            If(
                (
                    $relationship.FromTable -eq $fromTable -and
                    $relationship.ToTable -eq $toTable -and
                    $relationship.FromColumn -eq $fromColumn -and
                    $relationship.ToColumn -eq $toColumn
                ) -or 
                # check opposite direction as well
                (
                    $relationship.ToTable -eq $fromTable -and
                    $relationship.FromTable -eq $toTable -and
                    $relationship.ToColumn -eq $fromColumn -and
                    $relationship.FromColumn -eq $toColumn
                )
            ) {
                $relationshipExists = $true
                continue
            }
        }

        If($relationshipExists -and -not $replace) {
            Write-Host "Relationship already exists" -ForegroundColor Red
            Write-Host "Skipping this relationship" -ForegroundColor Gray
            continue
        }
        
        Write-Host "Validating the relationship" -ForegroundColor Gray 

        # https://stackoverflow.com/questions/40501374/unable-to-create-relationships-in-ssas-2016-using-tabular-model-programming-for

        $relationship = New-Object Microsoft.AnalysisServices.Tabular.SingleColumnRelationship
        $relationship.FromColumn = $fromColumn
        $relationship.ToColumn   = $toColumn
        $relationship.CrossFilteringBehavior = $crossFilteringBehavior

        $validationResult = $relationship.Validate()

        If($validationResult.ContainsErrors) {
            Write-Host "Validation failed" -ForegroundColor Red
            ForEach($error in $validationResult.Errors) {
                Write-Host $error -ForegroundColor Red
            }
            
            Write-Host "Skipping this relationship" -ForegroundColor Gray

            continue
        }

        Write-Host "Validation successful" -ForegroundColor Green
        $db.Model.Relationships.Add($relationship)

        Write-Host "Relationship added" -ForegroundColor Green
    }

    Write-Host "Saving model changes" -ForegroundColor Gray
    $db.Model.SaveChanges() | Out-Null
}

<#
                     _         _        _     _           
                    | |       | |      | |   | |          
  ___ _ __ ___  __ _| |_ ___  | |_ __ _| |__ | | ___  ___ 
 / __| '__/ _ \/ _` | __/ _ \ | __/ _` | '_ \| |/ _ \/ __|
| (__| | |  __/ (_| | ||  __/ | || (_| | |_) | |  __/\__ \
 \___|_|  \___|\__,_|\__\___|  \__\__,_|_.__/|_|\___||___/
#>

Function New-SSASTables {
    param(
        $db,
        $tableDefinitionList
    )

    $tableCollection = @()

    ForEach($tableDefinition in $tableDefinitionList) {
        $table = New-Object Microsoft.AnalysisServices.Tabular.Table
        
        Write-Host "Creating table" -ForegroundColor Gray -NoNewline
        Write-Host "'$($tableDefinition.Name)'" -ForegroundColor Magenta
        Write-Host "'$($tableDefinition.Expression)'" -ForegroundColor Yellow

        $table.Name = $tableDefinition.Name
        $table.IsHidden = $tableDefinition.IsHidden

        $partition = New-Object Microsoft.AnalysisServices.Tabular.Partition
        
        $partitionSource = New-Object Microsoft.AnalysisServices.Tabular.CalculatedPartitionSource
        $partitionSource.Expression = $tableDefinition.Expression

        $partition.Source = $partitionSource
        $table.Partitions.Add($partition)

        Write-Host "Adding table to model" -ForegroundColor Gray

        $db.Model.Tables.Add($table)

        Write-Host "Refreshing table" -ForegroundColor Gray

        # to activate a calc table, we need to refresh it
        $db.Model.Tables[$table.Name].RequestRefresh(1)
    }

    Write-Host "Saving model changes" -ForegroundColor Gray
    $db.Model.SaveChanges() | Out-Null
}

<#
 _        _     _                                                        
| |      | |   | |                                                       
| |_ __ _| |__ | | ___  ___   ______   _ __ ___ _ __ ___   _____   _____ 
| __/ _` | '_ \| |/ _ \/ __| |______| | '__/ _ \ '_ ` _ \ / _ \ \ / / _ \
| || (_| | |_) | |  __/\__ \          | | |  __/ | | | | | (_) \ V /  __/
 \__\__,_|_.__/|_|\___||___/          |_|  \___|_| |_| |_|\___/ \_/ \___|
#>
Function Remove-SSASTables {
    param(
        $db,
        $tableNameRegex
    )

    Write-Host "Removing tables by name regex: " -ForegroundColor Gray -NoNewline
    Write-Host "'$tableNameRegex'" -ForegroundColor Yellow

    $tableCollection = $db.Model.Tables | Where-Object Name -Match $tableNameRegex

    ForEach($table in $tableCollection) {
        Write-Host "Removing table: " -ForegroundColor Gray -NoNewline
        Write-Host "'$($table.Name)'" -ForegroundColor Magenta
        $db.Model.Tables.Remove($table) | Out-Null
    }

    Write-Host "Saving model changes" -ForegroundColor Gray
    $db.Model.SaveChanges() | Out-Null
}