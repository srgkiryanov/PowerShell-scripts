$path = "путь вашего сетевого ресурса" # путь вашего сетевого ресурса
$include = @("*.mov", "*.avi", "*.mpeg", "*.wmv", "*.mp4") # видео-файлы
$includefoto = @("*.jpg", "*.jpeg", "*.tiff", "*.gif", "*.png") # фотки
$exclude = @("*.mov", "*.avi", "*.mpeg", "*.wmv", "*.mp4", "*.jpg", "*.jpeg", "*.tiff", "*.gif", "*.png") # .doc, .xls и прочие файлы
$SumFileSizeLess3GBPerMonth = 0; $SumFileSizeLess3GBMoreMonth = 0
$SumFileSizeMore3GBPerMonth = 0; $SumFileSizeMore3GBMoreMonth = 0
$sumFotoFileSizePerMonth = 0; $sumFotoFileSizeMoreMonth = 0
$sumOtherFilePerMonth = 0; $sumOtherFileMoreMonth = 0
$i1 = 0; $i2 = 0; $i3 = 0; $i4 = 0; $i5 = 0; $i6 = 0; $i7 = 0; $i8 = 0
$sumAllFiles = 0; $kolAllFiles = 0
#----Видео-файлы размером меньше 3Gb и датой создания за последний месяц
$filesLess3GBPerMonth = Get-ChildItem -Path $path -Recurse -Include $include | Where-Object -FilterScript {($_.LastWriteTime -ge "2020-06-28") -and ($_.Length -lt 3gb)}
foreach ($file in $filesLess3GBPerMonth){
    $SumFileSizeLess3GBPerMonth = $SumFileSizeLess3GBPerMonth + [math]::Round($file.length /1GB,2)
    $i1++
}
#----Видео-файлы размером меньше 3Gb и датой создания более месяца назад
$filesLess3GBMoreMonth = Get-ChildItem -Path $path -Recurse -Include $include | Where-Object -FilterScript {($_.LastWriteTime -lt "2020-06-28") -and ($_.Length -lt 3gb)}
foreach ($file in $filesLess3GBMoreMonth){
    $SumFileSizeLess3GBMoreMonth = $SumFileSizeLess3GBMoreMonth + [math]::Round($file.length /1GB,2)
    $i2++
}
#----Видео-файлы размером более 3Gb и датой создания за последний месяц
$filesMore3GBPerMonth = Get-ChildItem -Path $path -Recurse -Include $include | Where-Object -FilterScript {($_.LastWriteTime -ge "2020-06-28") -and ($_.Length -ge 3gb)}
foreach ($file in $filesMore3GBPerMonth){
    #$file #можно раскомментить здесь, чтобы увидеть где лежат эти файлы
    $SumFileSizeMore3GBPerMonth = $SumFileSizeMore3GBPerMonth + [math]::Round($file.length /1GB,2)
    $i3++
}
#----Видео-файлы размером более 3Gb и датой создания более месяца назад
$filesMore3GBMoreMonth = Get-ChildItem -Path $path -Recurse -Include $include | Where-Object -FilterScript {($_.LastWriteTime -lt "2020-06-28") -and ($_.Length -ge 3gb)}
foreach ($file in $filesMore3GBMoreMonth){
    #$file #можно раскомментить здесь, чтобы увидеть где лежат эти файлы
    $SumFileSizeMore3GBMoreMonth = $SumFileSizeMore3GBMoreMonth + [math]::Round($file.length /1GB,2)
    $i4++
}
#----Фото-файлы датой создания за последний месяц
$FotoFilesPerMonth = Get-ChildItem -Path $path -Recurse -Include $includefoto | Where-Object -FilterScript {($_.LastWriteTime -ge "2020-06-28")}
foreach ($file in $FotoFilesPerMonth){
    $sumFotoFileSizePerMonth = $sumFotoFileSizePerMonth + [math]::Round($file.length /1GB,4)
    $i5++
}
#----Фото-файлы датой создания более месяца назад
$FotoFilesMoreMonth = Get-ChildItem -Path $path -Recurse -Include $includefoto | Where-Object -FilterScript {($_.LastWriteTime -lt "2020-06-28")}
foreach ($file in $FotoFilesMoreMonth){
    $sumFotoFileSizeMoreMonth = $sumFotoFileSizeMoreMonth + [math]::Round($file.length /1GB,4)
    $i6++
}
#----Прочие файлы датой создания за последний месяц
$otherfilesPerMonth = Get-ChildItem -Path $path -Recurse -Exclude $exclude | Where-Object -FilterScript {($_.LastWriteTime -ge "2020-06-28")}
foreach ($file in $otherfilesPerMonth){
    $sumOtherFilePerMonth = $sumOtherFilePerMonth + [math]::Round($file.length /1GB,4)
    $i7++
}
#----Прочие файлы датой создания более месяца назад
$otherfilesMoreMonth = Get-ChildItem -Path $path -Recurse -Exclude $exclude | Where-Object -FilterScript {($_.LastWriteTime -lt "2020-06-28")}
foreach ($file in $otherfilesMoreMonth){
    $sumOtherFileMoreMonth = $sumOtherFileMoreMonth + [math]::Round($file.length /1GB,4)
    $i8++
}
$sumAllFiles = $SumFileSizeLess3GBPerMonth + $SumFileSizeLess3GBMoreMonth + $SumFileSizeMore3GBPerMonth + $SumFileSizeMore3GBMoreMonth + $sumFotoFileSizePerMonth + $sumFotoFileSizeMoreMonth + $sumOtherFilePerMonth + $sumOtherFileMoreMonth
$kolAllFiles = $i1 +$i2 +$i3 +$i4 +$i5 +$i6 +$i7 +$i8
$SumFileSizeLess3GBPerMonth = [math]::Round($SumFileSizeLess3GBPerMonth,2); $SumFileSizeLess3GBMoreMonth = [math]::Round($SumFileSizeLess3GBMoreMonth,2); $SumFileSizeMore3GBPerMonth = [math]::Round($SumFileSizeMore3GBPerMonth,2)
$SumFileSizeMore3GBMoreMonth = [math]::Round($SumFileSizeMore3GBMoreMonth,2); $sumFotoFileSizePerMonth = [math]::Round($sumFotoFileSizePerMonth,2); $sumFotoFileSizeMoreMonth = [math]::Round($sumFotoFileSizeMoreMonth,2)
$sumOtherFilePerMonth = [math]::Round($sumOtherFilePerMonth,2); $sumOtherFileMoreMonth = [math]::Round($sumOtherFileMoreMonth,2); $sumAllFiles = [math]::Round($sumAllFiles,2)
write-host "Всего на данном ресурсе $kolAllFiles файлов, общим объёмом $sumAllFiles Gb"
write-host "Из них:"
write-host "Видео-файлы размером меньше 3Gb и датой создания за последний месяц: $SumFileSizeLess3GBPerMonth Gb. Количество Файлов: $i1"
write-host "Видео-файлы размером меньше 3Gb и датой создания более месяца назад: $SumFileSizeLess3GBMoreMonth Gb. Количество Файлов: $i2"
write-host "Видео-файлы размером более 3Gb и датой создания за последний месяц: $SumFileSizeMore3GBPerMonth Gb. Количество Файлов: $i3"
write-host "Видео-файлы размером более 3Gb и датой создания более месяца назад: $SumFileSizeMore3GBMoreMonth Gb. Количество Файлов: $i4"
write-host "Фото-файлы датой создания за последний месяц: $sumFotoFileSizePerMonth Gb. Количество Файлов: $i5"
write-host "Фото-файлы датой создания более месяца назад: $sumFotoFileSizeMoreMonth Gb. Количество Файлов: $i6"
write-host "Прочие файлы датой создания за последний месяц: $sumOtherFilePerMonth Gb. Количество Файлов: $i7"
write-host "Прочие файлы датой создания более месяца назад: $sumOtherFileMoreMonth Gb. Количество Файлов: $i8"

# load the appropriate assemblies
[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")
# create chart object
$Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart
$Chart.Width = 500
$Chart.Height = 400
$Chart.Left = 40
$Chart.Top = 30
# create a chartarea to draw on and add to chart
$ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
$Chart.ChartAreas.Add($ChartArea)
# add data to chart
[void]$Chart.Series.Add("Data")
# display the chart on a form
$Chart.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right -bor
                [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
# add data to chart
$Processes = Get-Process | Sort-Object -Property WS | Select-Object Name,WS,ID -Last 5
$ProcNames = @(foreach($Proc in $Processes){$Proc.Name + "_" + $Proc.ID})
$ProcNames = @("Видео-файлы <3Gb за месяц $SumFileSizeLess3GBPerMonth Gb", "Видео-файлы <3Gb более месяца $SumFileSizeLess3GBMoreMonth Gb", "Видео-файлы >3Gb за месяц $SumFileSizeMore3GBPerMonth Gb", "Видео-файлы >3Gb более месяца $SumFileSizeMore3GBMoreMonth Gb", "Фото-файлы за месяц $sumFotoFileSizePerMonth Gb", "Фото-файлы более месяца $sumFotoFileSizeMoreMonth Gb", "Прочие-файлы за месяц $sumOtherFilePerMonth Gb", "Прочие-файлы более месяца $sumOtherFileMoreMonth Gb")
$WS = @($SumFileSizeLess3GBPerMonth, $SumFileSizeLess3GBMoreMonth, $SumFileSizeMore3GBPerMonth, $SumFileSizeMore3GBMoreMonth, $sumFotoFileSizePerMonth, $sumFotoFileSizeMoreMonth, $sumOtherFilePerMonth, $sumOtherFileMoreMonth)
[void]$Chart.Titles.Clear()
[void]$Chart.Titles.Add("Сетевой ресурс $path")
#$ChartArea.AxisX.Title = "Примеры"
#$ChartArea.AxisY.Title = "Значения"
$Chart.Series["Data"].Points.DataBindXY($ProcNames, $WS)
######################
$Chart.Series["Data"].Sort([System.Windows.Forms.DataVisualization.Charting.PointSortOrder]::Descending, "Y")
# Find point with max/min values and change their colour
$maxValuePoint = $Chart.Series["Data"].Points.FindMaxByValue()
$maxValuePoint.Color = [System.Drawing.Color]::Red
$minValuePoint = $Chart.Series["Data"].Points.FindMinByValue()
$minValuePoint.Color = [System.Drawing.Color]::Green
# make bars into 3d cylinders
$Chart.Series["Data"]["DrawingStyle"] = "Cylinder"
# set chart type
$Chart.Series["Data"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Pie
# display the chart on a form
###################
# set chart options
$Chart.Series["Data"]["PieLabelStyle"] = "Outside"
$Chart.Series["Data"]["PieLineColor"] = "Black"
$Chart.Series["Data"]["PieDrawingStyle"] = "Concave"
($Chart.Series["Data"].Points.FindMaxByValue())["Exploded"] = $true
# display the chart on a form
$Form = New-Object Windows.Forms.Form
$Form.Text = "Анализ сетевого диска"
$Form.Width = 600
$Form.Height = 600
$Form.controls.add($Chart)
$Form.Add_Shown({$Form.Activate()})
$Form.ShowDialog()