function FindAdmRuk ([string]$login) {
    $AdminID = 'abc'; $MasAdmRuk = @()
    [string]$psd = 'SimvolniyIdentifikator'; $realytop = $false
    $srv = "YourServerName"; $dbase = "YourDBName"
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = "Server=$srv;Database=$dbase;Integrated Security=True"
    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
    $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $SqlConnection.Open()
    WHILE ($AdminID -ne "") {
        if ($psd -eq $AdminID){$realytop = $true}
        $query = "SELECT adminpersonuid FROM [YourDBName].[dbo].[YourTableName1] WHERE [domainaccountname] = '$login' AND [domainname] = 'DME'"
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand; $SqlCmd.Connection = $SqlConnection
        $SqlCmd.CommandText = $query; $SqlAdapter.SelectCommand = $SqlCmd
        $DataSet = New-Object System.Data.DataSet
        $SqlAdapter.Fill($DataSet); [string]$AdminID = $DataSet.Tables[0].adminpersonuid
#Поиск среди дублей истинного адм. руководителя. Можно изменить запрос, добавив значение [status] = 'Работает'", но...
#...тогда на шарах уволенных сотрудников не увидим адм. руководителей и уволенных адм. руководителей, также не увидем.
        $count = $DataSet.Tables[0].adminpersonuid.count - 1 # количество дублей adminpersonuid
        $AdminID = $AdminID.TrimEnd(); $AdminID = $AdminID.TrimStart()
        if ($count -ge 1) { # если количество дублей >= 1
            # ищем фамилию его руководителя по его странице в тел. справочнике
            $subquery = "SELECT adminperson FROM [YourDBName].[dbo].[YourTableName1] WHERE [domainaccountname] = '$login'"
            $SqlCmd = New-Object System.Data.SqlClient.SqlCommand; $SqlCmd.Connection = $SqlConnection
            $SqlCmd.CommandText = $subquery; $SqlAdapter.SelectCommand = $SqlCmd
            $DataSet = New-Object System.Data.DataSet
            $SqlAdapter.Fill($DataSet); [string]$HisPerson = $DataSet.Tables[0].adminperson
            $subprobel = $HisPerson.IndexOf(" ")
            $HisPerson = $HisPerson.Substring(0,$subprobel)
            #==============================================================================================
            for ($i = 0; $i -ne $count; $i++ ){
                $Length = $AdminID.Length; $probel = $AdminID.IndexOf(" "); if ($probel -lt 0) {$probel = $Length}
                $Substring = $AdminID.Substring(0,$probel)  # выделяем один из ID из дублей adminpersonuid
                # и получаем фамилию из дубля =============================================================
                $subquery2 = "SELECT lastname FROM [YourDBName].[dbo].[YourTableName1] WHERE [unid] = '$Substring'"
                $SqlCmd = New-Object System.Data.SqlClient.SqlCommand; $SqlCmd.Connection = $SqlConnection
                $SqlCmd.CommandText = $subquery2; $SqlAdapter.SelectCommand = $SqlCmd
                $DataSet = New-Object System.Data.DataSet
                $SqlAdapter.Fill($DataSet); [string]$MayBeAdmin = $DataSet.Tables[0].lastname
                #========================================================================================
                if ($HisPerson -eq $MayBeAdmin){$RealyHisAdmin = $Substring} # сравниваем значения фамилий по логину в ТС и по дублю adminpersonuid. Если равно - это реальная фамилия его адм. руководителя по ТС 
                $AdminID = $AdminID.Substring($probel, ($Length - $probel))
                $AdminID = $AdminID.TrimStart()
            }
            $AdminID = $RealyHisAdmin # присваиваем значение для дальнейшей обработки
        }
        #=======================================================================================================
        $query2 = "SELECT domainaccountname FROM [YourDBName].[dbo].[YourTableName1] WHERE [universalid] = '$AdminID'"
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand; $SqlCmd.Connection = $SqlConnection
        $SqlCmd.CommandText = $query2; $SqlAdapter.SelectCommand = $SqlCmd
        $DataSet = New-Object System.Data.DataSet
        $SqlAdapter.Fill($DataSet); [string]$login = $DataSet.Tables[0].domainaccountname # получаем один из логинов из цепочки адм. руководителей
        [string]$login =$login.Trim()
        If (($login -ne "") -and ($realytop -eq $false)){$MasAdmRuk += $login}
    }
    $SqlConnection.Close()
    return $MasAdmRuk
}
$manepaths = @("\\Path1\share1$\", "\\Path1\share2$\", "\\Path1\share3$\", "\\Path2\share4$\", "\\Path2\share5$\", "\\Path2\share6$\") # расскоментить, чтоб запустить на цикл все шары
$DataTime = Get-Date -Format "dd/MM/yyyy HH-mm"; $logfilenameAdm = "log_modify_AdmRuk_rights_" + $DataTime + " $env:username"; $logfilenameOther = "log_delete_Bad_rights_" + $DataTime + " $env:username"
foreach($manepath in $manepaths){
    $massive = Get-ChildItem -Path $manepath
    foreach($user in $massive){
        $path = $manepath+$user
        $maneuser = $path.Substring($path.LastIndexOf("\") + 1, $path.Length - $path.LastIndexOf("\") - 1)
        $ACL = Get-Acl -path $path
        Import-Module ActiveDirectory
        $imeyutdostup = $ACL.Access.IdentityReference
        $schetchikimen=0
        foreach($userfromshare in $imeyutdostup){
            $positiondrob = $userfromshare.Value.LastIndexOf("\")
            $user = $userfromshare.Value.Substring($positiondrob + 1, $userfromshare.Value.Length - $positiondrob - 1)
            $userfromad = Get-ADUser -Filter {samaccountname -like $user}
            if (($maneuser -ne $userfromad.samaccountname) -and (-not !$userfromad)){
                $pravakotorieimeetUZ = $ACL.Access.FileSystemRights.Item($schetchikimen)#тип прав на шару, которые имеет УЗ $userfromad.samaccountname/ она же $userfromshare
                $MassiveAdmRukovoditelei = FindAdmRuk ($maneuser) #получаем цепочку адм.руководителей
                $userfromad = $userfromad.samaccountname #логин УЗ, которая имеет права на шару
                if ($MassiveAdmRukovoditelei -contains $userfromad){ # если массив адм. руководителей включает в себя $userfromshare, проводим проверку на корректность галок
                    if(($pravakotorieimeetUZ -eq "FullControl") -or($pravakotorieimeetUZ -eq "Modify, Synchronize") -or($pravakotorieimeetUZ -eq "Write, ReadAndExecute, Synchronize") -or($pravakotorieimeetUZ -eq "ReadAndExecute, Synchronize")){
                        $rule = New-Object System.Security.AccessControl.FileSystemAccessRule($userfromad, ‘ReadAndExecute’, 'ContainerInherit, ObjectInherit', 'None', ’Allow’) #выставление на чтение и выполнение
                        #$rule = New-Object System.Security.AccessControl.FileSystemAccessRule($userfromad, ‘Read, listdirectory’, 'ContainerInherit, ObjectInherit', 'None', ’Allow’) #выставление только на чтение
                        $acl.ResetAccessRule($rule); $ACL | Set-Acl $path
                        Write-Host "Права адм.-ого руководителя $userfromad к шаре $path пользователя $maneuser изменены на доступ на чтение" #закомментить при желании
                        $strbadright = "Права адм.-ого руководителя $userfromad к шаре $path пользователя $maneuser изменены на доступ на чтение"
                        Add-content $PSScriptRoot\"$logfilenameAdm".txt -value $strbadright
                    }
                    else {#закомментить при желании
                        Write-Host "УЗ адм.-ого руководителя $userfromad имеет коректные права на шару $maneuser $path" #закомментить при желании
                        $strrightright = "УЗ адм.-ого руководителя $userfromad имеет коректные права на шару $maneuser $path"#закомментить при желании
                        Add-content $PSScriptRoot\Right_rights_adm_ruk.txt -value $strrightright #закомментить при желании
                    }#закомментить при желании

                }
                else { # во всех остальных случаях - грубое нарушение
                    $rule = New-Object System.Security.AccessControl.FileSystemAccessRule($userfromad, ‘FullControl, Modify, ReadAndExecute, Write’, 'ContainerInherit, ObjectInherit', 'None', ’Allow’)
                    $acl.RemoveAccessRule($rule); $ACL | Set-Acl $path
                    Write-Host "УЗ $userfromad лишена прав доступа к шаре $path пользователя $maneuser"
                    $strbadright = "УЗ $userfromad лишена прав доступа к шаре $path пользователя $maneuser"
                    Add-content $PSScriptRoot\"$logfilenameOther".txt -value $strbadright
                }
            }
            $schetchikimen++
        }
    }
}