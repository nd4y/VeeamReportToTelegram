#region Functions
$error.Clear()
function Get-VBRLatestRestorePointDate { # Получает дату последней точки восстановления
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [System.String]
        $VBRBackupName,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateSet('VMware Backup', 'File Backup', 'Linux Agent Backup')] # Другие типы бекапов добавлю по мере необходимости
        [System.String]
        $VBRBackupType
    )

    $InputDateFormat = 'MM/d/yyyy h:mm:ss tt'

    if ($VBRBackupType -eq 'VMware Backup') {
        $result = try {
            Get-Date(
                [datetime]::parseexact(
                    (Get-VBRJob -Name $VBRBackupName).GetLastBackup().LastPointCreationTime, $InputDateFormat, $null
                )
            )
        }
        catch {
            'No restore points'
        }
    }

    if ($VBRBackupType -eq 'File Backup') {
        $result = try {
            Get-Date(
                (Get-VBRNASBackup -Name $VBRBackupName).LastRestorePointCreationTime
                )
        }
        catch {
            'No restore points'
        }
    }

   
    if ($VBRBackupType -eq 'Linux Agent Backup') {  # Выводит дату последнего запуска джобы бекапа, даже если бекап был создан неудачно. Вероятно, в интерфейсе Veeam Console это тоже работает по этому принципу.
                                                    #  Это не работает https://helpcenter.veeam.com/docs/backup/powershell/get-vbrrestorepoint.html?ver=110
        $result = try {
            Get-Date(
                (Get-VBRJob -Name $VBRBackupName).FindLastSession().CreationTime
            )
        }
        catch {
            'No restore points'
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
function Add-EmojiAtTheBegginingOfTheString { #Добавляет URL encoded Emoji для Telegram
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [String]
        $String,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateSet('Green', 'Yellow', 'Red')]
        [String]
        $Color
        )
        
    $EmojiMap = @{
        'Green'     = '%E2%9C%85'
        'Yellow'    = '%E2%9A%A0%EF%B8%8F'
        'Red'       = '%E2%9D%8C'
    }
    
    $result = $EmojiMap.$Color + $String

    return $result
}
function Send-MessageToTelegramChatViaBot { # Отправляет сообщение в телеграм. 
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [String]
        $BotToken,
        [Parameter(Mandatory = $true, Position = 1)]
        [String]
        $ChatId,
        [Parameter(Mandatory = $true, Position = 2)]
        [String]
        $Message,
        [Parameter(Mandatory = $false, Position = 3)]
        [String]
        $ParseMode = 'html'


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
function Get-FormattedRPO { # Конвертирует значение RPO в строку с Emoji
    Param(
        [Parameter(Mandatory = $true, Position = 0)]
        [int]
        $RPO,
        [Parameter(Mandatory = $true, Position = 1)]
        [hashtable]
        $RPOMap
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
function Get-FormattedLastResult { # Конвертирует значение статуса в строку с Emoji
    Param(
        [Parameter(Mandatory = $true, Position = 0)]
        [String]
        $LastResult
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
        [String]
        $VBRBackupJobName,
        [Parameter(Mandatory = $true, Position = 1)]
        [String]
        $VBRBackupType
    )

    if ($VBRBackupType -eq 'File Backup') {
        $result = try {
            (Get-VBRJob -Name $VBRBackupJobName).FindLastSession().Info.BackupTotalSize
        }
        catch {
            $_
        }
    }


    else {
        $result = try {
            (((Get-VBRBackup -Name "$VBRBackupJobName*").GetAllStorages().Stats.BackupSize) | Measure-Object -Sum).Sum
        }
        catch {
            $_
        }
    }
    return $result 
}
#endregion Functions

#Main Script        
$RPOMap = [ordered]@{ # Максимально допустимое время в часах для получения зеленого, желтого или красного значка напротив значения RPO. Красный значек используется, если не выполняются условия для зеленого или желтого.
    '24'        = 'Green'
    '48'        = 'Yellow'
    #'999999'    = 'Red'
}

$BackupStatistics = @()
Get-VBRJob | ForEach-Object {

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
        'RPO'                         = Get-FormattedRPO -RPO (Get-VBRRecoveryPointObjective -LatestRestorePointDate (Get-VBRLatestRestorePointDate -VBRBackupName $_.Name -VBRBackupType $_.TypeToString)) -RPOMap $RPOMap
        'Job status'                  = $_.GetLastState()
        'Last result'                 = Get-FormattedLastResult -LastResult ($_.Info.LatestStatus)
        'Latest restore point'        = Get-FormattedDate -InputDate (Get-VBRLatestRestorePointDate -VBRBackupName $_.Name -VBRBackupType $_.TypeToString)
        'Total backup size'           = "$([math]::round((Get-VBRJobTotalBackupSize -VBRBackupJobName $_.Name -VBRBackupType $_.TypeToString)/1GB))GB"
    }

}
$BackupStatistics
$Header  = 'Veeam backup report for ' + (Get-FormattedDate -InputDate (Get-Date))
$Tail    = '[DEBUG] Number of data processing errors: ' + $error.Count 
$Message = $Header + '<pre>' + $($BackupStatistics | Sort-Object -Property 'Name' | Format-List | Out-String) + '</pre>' + $Tail
Send-MessageToTelegramChatViaBot -BotToken 'XXX' -ChatId 'YYY' -Message $Message
#Endregion Main Script