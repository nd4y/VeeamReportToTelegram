$error.Clear()
#region Functions
function Get-VBRLatestRestorePointDate { # Получает дату последней точки восстановления
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [System.String] $VBRBackupName,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateSet('VMware Backup', 'File Backup')] # Другие типы бекапов добавлю по мере необходимости
        [System.String] $VBRBackupType
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
        [String] $String,

        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateSet('Green', 'Yellow', 'Red')]
        [String] $Color
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
        [String] $BotToken,
        [Parameter(Mandatory = $true, Position = 1)]
        [String] $ChatId,
        [Parameter(Mandatory = $true, Position = 2)]
        [String] $Message,
        [Parameter(Mandatory = $false, Position = 3)]
        [String] $ParseMode = 'html',
        [Parameter(Mandatory = $false, Position = 4)]
        [int] $Attempts = '3', # Количество попыток отправки сообщения
        [Parameter(Mandatory = $false, Position = 5)]
        [int] $Timeout = '10' # Таймаут между попытками в секундах
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
            Start-Sleep -Seconds $Timeout
        }
    } 
    while (
        $Failed -and ($i -lt $Attempts)
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
        [int] $RPO
    )
        
    $RPOMap = [ordered]@{
        '24'        = 'Green'
        '48'        = 'Yellow'
        #'999999'    = 'Red'
    }

    foreach ($Element in $RPOMap.GetEnumerator()) {
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
        [String] $LastResult
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
#endregion Functions


$BackupStatistics = @()
Get-VBRJob | ForEach-Object {

    $BackupStatistics += [PSCustomObject]@{
        'Name'                        = $_.Name 
        'RPO'                         = Get-FormattedRPO -RPO (Get-VBRRecoveryPointObjective -LatestRestorePointDate (Get-VBRLatestRestorePointDate -VBRBackupName $_.Name -VBRBackupType $_.TypeToString))
        'Job status'                  = $_.GetLastState()
        'Latest result'               = Get-FormattedLastResult -LastResult ($_.Info.LatestStatus)
        'Latest restore point'        = Get-FormattedDate -InputDate (Get-VBRLatestRestorePointDate -VBRBackupName $_.Name -VBRBackupType $_.TypeToString)
        'Total backup size'           = "$([math]::round((((Get-VBRJob -Name $_.Name).FindLastSession().Info.BackupTotalSize)/1GB)))GB"
    }

}
$BackupStatistics
$Header  = 'Veeam backup report for ' + (Get-FormattedDate -InputDate (Get-Date))
$Tail    = '[DEBUG] Number of data processing errors: ' + $error.Count 
$Message = $Header + '<pre>' + $($BackupStatistics | Sort-Object -Property 'Name' | Format-List | Out-String) + '</pre>' + $Tail
Send-MessageToTelegramChatViaBot -BotToken 'YYYYYYYYYYYYYYYYYYYYYYY' -ChatId 'XXXXXXXXXXXXXXXXXXX' -Message $Message
