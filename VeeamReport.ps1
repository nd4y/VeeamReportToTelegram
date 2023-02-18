param (
    [Parameter(Position=0,mandatory=$false)]
    [string]$ConfigPath = "C:\scripts\VeeamReport.conf"
)
$error.Clear()
#region Functions
function Get-VBRLatestRestorePointDate { # Получает дату последней точки восстановления
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [System.String]$VBRBackupJobName,
        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateSet('VMware Backup', 'File Backup', 'Linux Agent Backup', 'Hyper-V Backup', 'Backup Copy', 'Hyper-V Backup Copy')] # Другие типы бекапов добавлю по мере необходимости
        [System.String]$VBRBackupType
    )
    $InputDateFormat = 'M/d/yyyy h:mm:ss tt'
    switch  ($VBRBackupType) {
        'VMware Backup' {
            $result = try {
                Get-Date(
                    [datetime]::parseexact(
                        (Get-VBRJob -Name $VBRBackupJobName).GetLastBackup().LastPointCreationTime, $InputDateFormat, $null
                    )
                )
            }
            catch {
                'No restore points'
            }
        }
        'Hyper-V Backup' {
            $result = try {
                Get-Date(
                    [datetime]::parseexact(
                        (Get-VBRJob -Name $VBRBackupJobName).GetLastBackup().LastPointCreationTime, $InputDateFormat, $null
                    )
                )
            }
            catch {
                'No restore points'
            }
        }
        'File Backup' {
            $result = try {
                Get-Date(
                    (Get-VBRNASBackup -Name $VBRBackupJobName).LastRestorePointCreationTime | Sort-Object -Descending | Select-Object -First 1
                    )
            }
            catch {
                'No restore points'
            }
        }
        'Backup Copy' {
            $result = try {
                Get-Date(
                    (Get-VBRJob -Name $VBRBackupJobName).GetLastBackup().MetaUpdateTime
                    )
            }
            catch {
                'No restore points'
            } 
        }
        default {
            $result = try {
                Get-Date(
                    (Get-VBRJob -Name $VBRBackupJobName).GetLastBackup().CreationTime
                    )
            }
            catch {
                'No restore points'
            } 
        }
    }
    return $result
}
function Get-VBRRecoveryPointObjective { #Выводит округленное время в часах между текущей датой и переданной в параметре датой
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        $LatestRestorePointDate
    )
    $result = try {
        [math]::round(
            ((Get-Date) - $LatestRestorePointDate).TotalHours
            )
        }
        catch {
            '9999'
        }
    return $result
}
function Get-FormattedRPO { # Конвертирует значение RPO в строку с Emoji
    Param(
        [Parameter(Mandatory = $true, Position = 0)]
        [int]$RPO,
        [Parameter(Mandatory = $true, Position = 1)]
        [hashtable]$RPOMap
    )
    foreach ($Element in ($RPOMap.GetEnumerator() | Sort-Object -Property 'Key')) {
        if ($RPO -le $Element.Key) {
            $Color = $Element.Value
            break
        }
        else {
            $Color = 'Red'
        }
    }
    $result = Add-EmojiAtTheBegginingOfTheString -Color $Color -String ("$RPO" + 'h')
    return $result
}
function Import-Config { #Импортирует параметры скрипта и проверяет наличие обязательных параметров.
    param (
        [Parameter(Position=0,mandatory=$true)]
        [string]$ConfigPath
    )
    try {
        $result = @{}
        Get-Content -Path $ConfigPath | Foreach-Object {
            if ($_.Split('=')[0] -notmatch "^;|#.*") { # Исключить закомментированные строки
                $result += [hashtable]@{
                    $_.Split('=')[0] = $_.Split('=')[1]
                }
            }
        }
    }
    catch {
        break
    }
    #region Precheck Imported Params
    $RequiredParamsList = @(
        'TelegramBotToken'
        'TelegramChatId'
    )
    $RequiredParamsList | ForEach-Object {
        if ($result.GetEnumerator().Name -contains $_) {
        }
        else {
            break
        } 
    }
    #endregion Precheck Imported Params
    return $result
}
function Send-MessageToTelegramChatViaBot { # Отправляет сообщение в телеграм. 
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [String]$BotToken,
        [Parameter(Mandatory = $true, Position = 1)]
        [String]$ChatId,
        [Parameter(Mandatory = $true, Position = 2)]
        [String]$Message,
        [Parameter(Mandatory = $false, Position = 3)]
        [String]$ParseMode = 'html'
    )
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $i = [int]0
    do {
        $Failed = $false
        $i++
        try {
            Write-Host "Trying to send a message. Attempt number $i"
            $null = Invoke-RestMethod -Uri "https://api.telegram.org/bot$($BotToken)/sendMessage?chat_id=$($ChatId)&parse_mode=$($ParseMode)&text=$($Message)"
        }
        catch {
            $ErrorLog = $_
            $Failed = $true
            Start-Sleep -Seconds 10
        }
    } 
    while (
        $Failed -and ($i -lt 3)
        )
        
    if ($Failed) {
        $result = $ErrorLog
    }
    else {
        $result = 'Message sent successfully'
    }
    return $result
}
function Get-FormattedDate { # Конвертирует дату из datetime в строку в удобночитаемом виде
    Param(
        [Parameter(Mandatory = $true, Position = 0)]
        $InputDate
    )
    if ($InputDate.GetType().Name -eq 'DateTime') {
        $DateFormat = 'd MMMM yyyy HH:mm'
        $result = (Get-Date($InputDate) -Format $DateFormat)
    }
    else {
        $result = $InputDate
    }    
    return $result
}
function Get-FormattedLastResult { # Конвертирует значение статуса в строку с Emoji
    Param(
        [Parameter(Mandatory = $true, Position = 0)]
        [String]$LastResult
    )    
    $LastResultsMap = [ordered]@{
        'Success' = 'Green'
        'Warning' = 'Yellow'
        'Failed'  = 'Red'
        'None'    = 'Yellow'
    }
    foreach ($Element in $LastResultsMap.GetEnumerator()) {
        if ($LastResult -eq $Element.Key) {
            $Color = $Element.Value
            break
        }
        else {
            $Color = 'Red'
        }
    }
    $result = Add-EmojiAtTheBegginingOfTheString -Color $Color -String $LastResult
    return $result
    
} 
function Get-VBRJobTotalBackupSize { # Расчитывает значение размеров всех Restore Points в рамках одной Backup Job. Возвращает значение в [int] в байтах. Работает неправильно, если есть две джобы где одна джоба полностью включает в себя часть названия другой джобы.
    Param(
        [Parameter(Mandatory = $true, Position = 0)]
        [String]$VBRBackupJobName,
        [Parameter(Mandatory = $true, Position = 1)]
        [String]$VBRBackupType
    )

    switch ($VBRBackupType) {
        'File Backup' {
            $result = try {
                (Get-VBRJob -Name $VBRBackupJobName).FindLastSession().Info.BackupTotalSize
            }
            catch {
                $_
            }
        }
        'Backup Copy' {
            $result = try {
                (Get-VBRNASBackupCopyJob -Name $VBRBackupJobName).FindLastSession().Info.Progress.TotalUsedSize
            }
            catch {
                $_
            }
        }
        default {
            $result = try {
                (((Get-VBRBackup -Name "$VBRBackupJobName*").GetAllStorages().Stats.BackupSize) | Measure-Object -Sum).Sum
            }
            catch {
                $_
            }
        }
    }
    return $result 
}
function Add-EmojiAtTheBegginingOfTheString { #Добавляет URL encoded Emoji для Telegram
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [String]$String,
        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateSet('Green', 'Yellow', 'Red')]
        [String]$Color
        )  
    $EmojiMap = @{
        'Green'     = '%E2%9C%85'
        'Yellow'    = '%E2%9A%A0%EF%B8%8F'
        'Red'       = '%E2%9D%8C'
    }
    $result = $EmojiMap.$Color + $String
    return $result
}
#endregion Functions
#Main Script        
$BackupStatistics = @()
Get-VBRJob | ForEach-Object {
    $RPOMap = [ordered]@{ # Максимально допустимое время в часах для получения зеленого, желтого или красного значка напротив значения RPO. Красный значек используется, если не выполняются условия для зеленого или желтого.
        '24'        = 'Green'
        '48'        = 'Yellow'
        #'999999'    = 'Red'
    }
    #region Custom RPO Settings
    $CustomRPOMap = @( # Переопределение дефолтного RPO (24 часа) для некоторых видов бекапов
        [PSCustomObject]@{JobName = 'Custom Backup';   RPOMap = [ordered]@{'168' = 'Green'; '336' = 'Yellow'}}
        [PSCustomObject]@{JobName = 'Custom 2 Backup'; RPOMap = [ordered]@{'9998' = 'Green'; '9999' = 'Yellow'}}
    )
    foreach ($Element in $CustomRPOMap) {
        if ($_.Name -eq $Element.JobName) {
            $RPOMap = $Element.RPOMap
        }
    }
    #endregion Custom RPO Settings
    $BackupStatistics += [PSCustomObject]@{
        'Name'                        = $_.Name 
        'Job Type'                    = $_.TypeToString
        'Job status'                  = $_.GetLastState()
        'RPO'                         = Get-FormattedRPO -RPO (Get-VBRRecoveryPointObjective -LatestRestorePointDate (Get-VBRLatestRestorePointDate -VBRBackupJobName $_.Name -VBRBackupType $_.TypeToString)) -RPOMap $RPOMap
        'Last result'                 = Get-FormattedLastResult -LastResult ($_.Info.LatestStatus)
        'Latest restore point'        = Get-FormattedDate -InputDate (Get-VBRLatestRestorePointDate -VBRBackupJobName $_.Name -VBRBackupType $_.TypeToString)
        'Total backup size'           = "$([math]::round((Get-VBRJobTotalBackupSize -VBRBackupJobName $_.Name -VBRBackupType $_.TypeToString)/1GB))GB"
    }
}
$BackupStatistics
$Header  = 'Veeam backup report for ' + (Get-FormattedDate -InputDate (Get-Date))
$Tail    = '[DEBUG] Number of data processing errors: ' + $error.Count 
$Message = $Header + '<pre>' + $($BackupStatistics | Sort-Object -Property 'Name' | Format-List | Out-String) + '</pre>' + $Tail
Send-MessageToTelegramChatViaBot -BotToken $((Import-Config -ConfigPath $ConfigPath).TelegramBotToken) -ChatId $((Import-Config -ConfigPath $ConfigPath).TelegramChatId) -Message $Message
#Endregion Main Script