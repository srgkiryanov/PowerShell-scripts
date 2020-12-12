Import-Module -Name HPEOACmdlets
$Credential = Get-Credential
$BladeFile = "$($PSScriptRoot)\Blades.xlsx"
$SheetsName = @("ЦОД", "ЦОД(copy)")
$deshdesh = "https://"
$ShassiForEdit = "10.1.2.5" # YourOAIp
$ExcelObj = New-Object -ComObject Excel.Application
$Workbook = $ExcelObj.workbooks.open($BladeFile)
$Worksheets = $Workbooks.worksheets
foreach ($SheetName in $SheetsName){
    $WorkSheet = $WorkBook.sheets.item($SheetName)
    $MaxRows = ($WorkSheet.UsedRange.Rows).count
    $MaxColumns = ($WorkSheet.UsedRange.Columns).count
    #Write-Host "MaxRows = " $MaxRows
    #Write-Host "MaxColumns = " $MaxColumns
    for ($row = 1; $row -le $MaxRows; $row++) {
        for ($col = 1; $col -lt ($MaxColumns+3); $col++) {
            [string]$CellsValue = $WorkSheet.UsedRange.Cells($row,$col).Text
            $CellsValue = $CellsValue.ToLower()
            if ($CellsValue.Contains($deshdesh)){
                $ipShassi = $CellsValue.Substring($CellsValue.LastIndexOf("/")+1)
                $ipShassi
                if($ipShassi -eq $ShassiForEdit){
                    Write-Host "Шасси найдено " $ipshassi
                    $Session = Connect-HPEOA -OA $ipShassi -Credential $Credential
                    $Shassi = Get-HPEOAServerInfo -Connection $Session # $Shassi = Get-HPEOAServerName -Connection $Session
                    $j = 0; $ilo = 0
                    for ($i = 0; $i -lt 16; $i++){
                        if ($Shassi.ServerBlade.ServerName[$i] -ne $null){
                            $strbay = "Bay"+$Shassi.ServerBlade.Bay[$i]; $strbay
                            $ServerBladeServerName = $Shassi.ServerBlade.ServerName[$i]
                            if ($ServerBladeServerName.contains(".")){$ServerBladeServerName = $ServerBladeServerName.Substring(0, $ServerBladeServerName.IndexOf("."))}
                            $ServerBladeProductName = $Shassi.ServerBlade.ProductName[$i]
                            $bladememory = $Shassi.ServerBlade.memory[$i]; $bladememory = $bladememory.Substring(0, $bladememory.IndexOf(" ")) # оперативка в МБ
                            $bladememory = [Math]::Round($bladememory/1024,0) # оперативка в ГБ
                            [string]$bladememory1 = "$bladememory Gb RAM"
                            $cpu1 = $Shassi.ServerBlade.cpu.value[$j]
                            $cpu2 = $Shassi.ServerBlade.cpu.value[$j+1]
                            if ($cpu1 -eq $cpu2){$allcpu = "2 x "+$cpu1}
                            $ServerBladeSerialNumber = $Shassi.ServerBlade.SerialNumber[$i]
                            $iLoType = $Shassi.ServerBlade.ManagementProcessorInformation.Type[$ilo]
                            $strforWrite = $ServerBladeServerName.ToUpper() + " " + $ServerBladeProductName + " " + $allcpu + " " + $bladememory1 + " " + $iLoType
                            for ($bayrow = $row; $bayrow -le ($row + 12); $bayrow++) {
                                for ($baycol = 1; $baycol -lt ($MaxColumns+3); $baycol++) {
                                    [string]$BayCellsValue = $WorkSheet.UsedRange.Cells($bayrow,$baycol).Text
                                    if ($BayCellsValue -eq $strbay){
                                        $WorkSheet.Columns.Item($baycol).Rows.Item($bayrow+1) = $strforWrite
                                        $WorkSheet.Columns.Item($baycol).Rows.Item($bayrow+1).font.bold = $false
                                        $WorkSheet.Columns.Item($baycol).Rows.Item($bayrow+2) = $ServerBladeSerialNumber
                                        $WorkSheet.Columns.Item($baycol).Rows.Item($bayrow+2).font.bold = $false
                                    }
                                }
                            }
                            $j = $j + 2; $ilo++
                        }
                    }
                    $Shassi = $null
                }
            }
        }
    }
}
$Workbook.SaveAs("$PSScriptRoot\Blades.xlsx")
$Workbook.Close($true)
$ExcelObj.Quit()