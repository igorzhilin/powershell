<#
    Boilerplate collection
#>

Import-Module ImportExcel

Function Save-ObjectsToExcelSheets {
    <#
        saves a list of objects into excel sheets
        
        $objectToSheetMapping is array of PSCustomObject. SheetName = sheet name, Data = psobject with data
        
        dependencies:
            importexcel

        example:
            $excelFilePath = 'C:\temp\test.xlsx'
            Remove-Item $excelFilePath
            $items1 = Get-ChildItem 'C:\temp'
            $items2 = Get-Date | Select-Object *
            $objectToSheetMapping = @(
                [PSCustomObject]@{SheetName = "item1"; Data = $items1},
                [PSCustomObject]@{SheetName = "item2"; Data = $items2}
            )
            Save-ObjectsToExcelSheets -excelFilePath $excelFilePath -objectToSheetMapping $objectToSheetMapping
    #>

    param($excelFilePath, $objectToSheetMapping)

    Write-Host "Writing to file " -ForegroundColor Gray -NoNewline
    Write-Host $excelFilePath -ForegroundColor Magenta 

    ForEach($sheetData in $objectToSheetMapping) {
        $sheetName = $sheetData.SheetName
        $sheetData = $sheetData.Data | Select-Object *

        Write-Host "Writing excel sheet " -ForegroundColor Gray -NoNewline
        Write-Host $sheetName -ForegroundColor Magenta 
        
        $sheetData | Export-Excel $excelFilePath -WorksheetName $sheetName -TableName $sheetName -AutoSize -AutoFilter 
    }
}