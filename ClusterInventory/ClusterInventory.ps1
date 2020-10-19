$User="yourdomain\yourusername"
$Pass = "YourPasword"
$Password = $pass|ConvertTo-SecureString -AsPlainText -Force
$Credential = New-Object -TypeName System.Management.Automation.PSCredential($User, $Password)
$vcentresdme = @("vcenter1", "vcenter2", "vcenter3", "vcenter4", "vcenter5")
#---
$ExcelObj = New-Object -comobject Excel.Application
$xlsClInvWorkBook = $ExcelObj.Workbooks.Add()
$xlsClInWorkSheet = $xlsClInvWorkBook.Worksheets.Item(1)
$xlsClInWorkSheet.Name = 'Hardware'; $xlsClInWorkSheet.Cells.Item(1,1) = 'Host'; $xlsClInWorkSheet.Cells.Item(1,2) = 'Model'; $xlsClInWorkSheet.Cells.Item(1,3) = 'Vcenter'
$xlsClInWorkSheet.Cells.Item(1,4) = 'Кластер'; $xlsClInWorkSheet.Cells.Item(1,5) = 'PowerState'; $xlsClInWorkSheet.Cells.Item(1,6) = 'ConnectionState'
$xlsClInWorkSheet.Rows.Item(1).Font.Bold = $true; $xlsClInWorkSheet.Rows.Item(1).Font.size = 15
$xlsClInWorkSheet.Columns.Item(1).ColumnWidth = 16.14; $xlsClInWorkSheet.Columns.Item(2).ColumnWidth = 27.43; $xlsClInWorkSheet.Columns.Item(3).ColumnWidth = 14.14
$xlsClInWorkSheet.Columns.Item(4).ColumnWidth = 21.86; $xlsClInWorkSheet.Columns.Item(5).ColumnWidth = 14; $xlsClInWorkSheet.Columns.Item(6).ColumnWidth = 19.86
#---
[int]$i = 2
foreach ($vcentreorg in $vcentresorg) {
    Connect-VIServer $vcentreorg -Credential $Credential -Force -SaveCredentials
    $clustersorg = Get-Cluster
    foreach ($clusterorg in $clustersorg){
        $hostsorg = Get-VMHost -Location $clusterorg.name
        $xlsClInWorkSheet2 = $xlsClInvWorkBook.worksheets.Add()
        $xlsClInWorkSheet2.Name = $clusterorg.Name
        [int]$j = 2
        foreach ($hostorg in $hostsorg){
            $svmname = $hostorg.Name.Substring(0, $hostorg.Name.IndexOf("."))
            [string]$PowerState = $hostorg.PowerState; [string]$ConnectionState = $hostorg.ConnectionState
            $xlsClInWorkSheet.Columns.Item(1).Rows.Item($i) = $svmname; $xlsClInWorkSheet.Columns.Item(2).Rows.Item($i) = $hostorg.Model; $xlsClInWorkSheet.Columns.Item(3).Rows.Item($i) = $vcentreorg
            $xlsClInWorkSheet.Columns.Item(4).Rows.Item($i) = $clusterorg.name; $xlsClInWorkSheet.Columns.Item(5).Rows.Item($i) = $PowerState; $xlsClInWorkSheet.Columns.Item(6).Rows.Item($i) = $ConnectionState
            $i++
            $vmsorg = Get-VM -Location $hostorg
            foreach ($vmorg in $vmsorg){
                $xlsClInWorkSheet2.Cells.Item(1,1) = 'Name'; $xlsClInWorkSheet2.Cells.Item(1,2) = 'State'; $xlsClInWorkSheet2.Cells.Item(1,3) = 'CurrentHost'
                $xlsClInWorkSheet2.Rows.Item(1).Font.Bold = $true; $xlsClInWorkSheet2.Rows.Item(1).Font.size = 15
                $xlsClInWorkSheet2.Columns.Item(1).ColumnWidth = 36.5; $xlsClInWorkSheet2.Columns.Item(2).ColumnWidth = 12.86; $xlsClInWorkSheet2.Columns.Item(3).ColumnWidth = 16.14
                $xlsClInWorkSheet2.Columns.Item(1).Rows.Item($j) = $vmorg.Name; [string]$PowerStateVM = $vmorg.PowerState
                $xlsClInWorkSheet2.Columns.Item(2).Rows.Item($j) = $PowerStateVM; $xlsClInWorkSheet2.Columns.Item(3).Rows.Item($j) = $svmname
                $j++
            }
        }
    }
}
$xlsClInvWorkBook.SaveAs("$PSScriptRoot\ClusterInventory.xlsx")
$xlsClInvWorkBook.close($true)