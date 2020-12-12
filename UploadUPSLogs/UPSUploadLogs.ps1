# ================================
# = проверка работы branch в git =
# ================================
$UPSis = @("123.123.12.34", "234.567.12.34") # UPS IP-adress
$user = 'UPSLogin'
$password = 'UPSPassword'
$UploadFolder = "$($PSScriptRoot)\"
$FileforUpload1 = "data.txt"
$FileforUpload2 = "event.txt"
foreach ($UPS in $UPSis){
    $connection = New-Object System.Net.WebClient
    $connection.Credentials = New-Object System.Net.NetworkCredential("$user", "$password")
    $DataTime = Get-Date -Format "dd/MM/yyyy HH-mm"
    if ($UPS -eq "123.123.12.34"){$logfile = "_Your_1UPSName_"; $logfolder = "Your_1UPSName"}
    else {$logfile = "_Your_2UPSName_"; $logfolder = "Your_2UPSName"}
    $logfilename = $logfile + $DataTime
    $connection.DownloadFile("ftp://$UPS/$FileforUpload1", "$UploadFolder\$logfolder\$("data" + $logfilename).txt")
    $connection.DownloadFile("ftp://$UPS/$FileforUpload2", "$UploadFolder\$logfolder\$("event" + $logfilename).txt")
}