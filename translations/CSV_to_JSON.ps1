param($trainingName)
$basepath = $PSScriptRoot + "\"  
$xlsxfile = $basePath + "csvContent.xlsx"

if ((Test-Path $xlsxfile) -eq $true) {
    $objExcel = New-Object -ComObject Excel.Application
    $workbook = $objExcel.Workbooks.Open($xlsxfile)
    $workbook.refreshall()
                  
    Start-Sleep -s 10
    $sheet = $WorkBook.sheets.item("material")

    $totalNoOfRecords = ($sheet.UsedRange.Rows).count 
    $col = 0
    while ($null -ne $sheet.Cells.Item(1, $col).value2) {
        if ($col -gt 0) {
            $jsonBase = @{}

            for ($i = 2; $i -lt $totalNoOfRecords; $i++) {
                if (!$jsonBase.ContainsKey($sheet.cells.item($i, 1).value2)) {
                    $jsonBase.Add($sheet.cells.item($i, 1).value2, $sheet.cells.item($i, $col).value2)
                }
            }

            $countryCode = ($sheet.Cells.Item(1, $col).text).ToLower()
            if ($countryCode -ne "") {
                $outFile = $basepath + $countrycode + "_translation_$trainingName.json"
                $jsonBase | ConvertTo-Json -Depth 10 | Out-File $outFile -Encoding UTF8
            }
        }
        $col += 1
    }
    $objExcel.Quit()	
}