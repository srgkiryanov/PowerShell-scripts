function Generate-Report ($masBadPermission) {
    foreach ($iden in $masBadPermission){Write-Output "<tr><td>$($iden.printername)</td><td>$($iden.strpermissions)</td><td>$($iden.groupstoadd)</td></tr>"}
} 
function SnapMail {
    $msg = new-object Net.Mail.MailMessage; $smtp = new-object Net.Mail.SmtpClient($smtpServer)
    $msg.From = $MailFrom; $msg.To.Add($Mailto); $msg.Subject = "Выгрузка корректности опубликованных принтеров"
    $msg.IsBodyHtml = 1; $msg.Body = $mailreport; $smtp.Send($msg)
}
$smtpServer = "YoursmtpServer"; $MailFrom = "CorpMailAdressForSend"
$MailTo = "YourCorpMailAdress.ru"
$mailreport = $null
$printservers = @("printserver1", "printserver2", "printserver3")
foreach($vprint in $printservers){
    $FullListPrinters = (Get-Printer -ComputerName $vprint).Name
    $mailreport += Write-Output "<table cellpadding=""4"" border=""1""><tr class=""Title""><td colspan=""6"">$vprint</td></tr><tr class=""Title""><td>Printer name</td><td>Bad permissions (groups to del)</td><td>Groups to add</td></tr>"
    $adm="Administrators"
    foreach($printer in $FullListPrinters){
        $tipprav = (Get-Printer -ComputerName $vprint -Name $printer -Full).PermissionSDDL
        $MasUserwithPermission = @(); $strGroupsToAdd = ""
        $ACLObject = New-Object -TypeName System.Security.AccessControl.DirectorySecurity
        $ACLObject.SetSecurityDescriptorSddlForm($tipprav)
        $imeyutdostup = $ACLObject.Access.IdentityReference
        foreach($userfromshare in $imeyutdostup){
            $positiondrob = $userfromshare.Value.LastIndexOf("\")
            $user = $userfromshare.Value.Substring($positiondrob+1, $userfromshare.Value.Length - $positiondrob - 1)
            $user = $user.ToLower(); $printer = $printer.ToLower(); $adm = $adm.ToLower()
            if (!($MasUserwithPermission -contains $user)){$MasUserwithPermission = $MasUserwithPermission + $user}
        }
        if (!($MasUserwithPermission -contains $printer) -or !($MasUserwithPermission -contains $adm) -or($MasUserwithPermission.Count -ne 2)){
            $strMasUserwithPermission = $MasUserwithPermission -join ' '
            $strMasUserwithPermission = $strMasUserwithPermission.Replace($adm, ""); $strMasUserwithPermission = $strMasUserwithPermission.Replace($printer, "")
            if (!($MasUserwithPermission -contains $printer)){$strGroupsToAdd = $strGroupsToAdd + $printer + " "}
            if (!($MasUserwithPermission -contains $adm)){$strGroupsToAdd = $strGroupsToAdd + $adm + " "}
            if ($strGroupsToAdd -eq ""){$strGroupsToAdd = "-"}
            $masBadPermissions = @([pscustomobject]@{printername = $printer; strpermissions = $strMasUserwithPermission; groupstoadd = $strGroupsToAdd})
            $mailreport += Generate-Report $masBadPermissions
            Write-Host $vprint "\" $printer
        }
        $masBadPermissions = $null
    }
} #--
$mailreport += Write-Output "</table>"
$mailreport += "</body></html>"
SnapMail