param (
    [Parameter(Position=0,mandatory=$false)]
    [string]$ConfigPath = "C:\scripts\VeeamReport.conf"
)
$error.Clear()
#region Functions
function ConvertTo-DateTime { # Приводит значение к [datetime]. Строки парсит сначала в формате Veeam (en-US), потом в локали системы, чтобы скрипт работал на серверах с любой локалью
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        $InputDate
    )
    if ($InputDate -is [datetime]) {
        return $InputDate
    }
    $InputDateString = [string]$InputDate
    $InputDateFormat = 'M/d/yyyy h:mm:ss tt' # Формат, в котором Veeam отдает даты на en-US системах
    $ParsedDate = [datetime]::MinValue
    foreach ($Culture in @([System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.CultureInfo]::CurrentCulture)) {
        if ([datetime]::TryParseExact($InputDateString, $InputDateFormat, $Culture, [System.Globalization.DateTimeStyles]::None, [ref]$ParsedDate)) {
            return $ParsedDate
        }
        if ([datetime]::TryParse($InputDateString, $Culture, [System.Globalization.DateTimeStyles]::None, [ref]$ParsedDate)) {
            return $ParsedDate
        }
    }
    throw "Cannot parse date '$InputDateString'"
}
function Get-VBRLatestRestorePointDate { # Получает дату последней точки восстановления
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        $VBRBackupJob # Объект задания из Get-VBRJob. Типы заданий без особой обработки идут через ветку default
    )
    switch ($VBRBackupJob.TypeToString) {
        { $_ -in 'VMware Backup', 'Hyper-V Backup' } {
            $result = try {
                ConvertTo-DateTime -InputDate $VBRBackupJob.GetLastBackup().LastPointCreationTime
            }
            catch {
                'No restore points'
            }
        }
        'File Backup' {
            $result = try {
                ConvertTo-DateTime -InputDate (
                    (Get-VBRNASBackup -Name $VBRBackupJob.Name).LastRestorePointCreationTime | Sort-Object -Descending | Select-Object -First 1
                    )
            }
            catch {
                'No restore points'
            }
        }
        'Backup Copy' {
            $result = try {
                ConvertTo-DateTime -InputDate $VBRBackupJob.GetLastBackup().MetaUpdateTime
            }
            catch {
                'No restore points'
            }
        }
        default {
            $result = try {
                ConvertTo-DateTime -InputDate $VBRBackupJob.GetLastBackup().CreationTime
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
            9999
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
    $Color = 'Red' # Используется, если RPO не уложился ни в один порог
    foreach ($Element in ($RPOMap.GetEnumerator() | Sort-Object -Property { [double]$_.Key })) { # Пороги сортируются как числа: строковая сортировка ставит '168' перед '48' и выбирает не тот цвет
        if ($RPO -le [double]$Element.Key) {
            $Color = $Element.Value
            break
        }
    }
    if ($Color -notin @('Green', 'Yellow', 'Red')) { # Защита от опечатки в цвете в CustomRPOMap
        $Color = 'Red'
    }
    $result = Add-EmojiAtTheBeginningOfTheString -Color $Color -String ("$RPO" + 'h')
    return $result
}
function Import-Config { #Импортирует параметры скрипта и проверяет наличие обязательных параметров.
    param (
        [Parameter(Position=0,mandatory=$true)]
        [string]$ConfigPath
    )
    if (-not (Test-Path -Path $ConfigPath)) {
        throw "Config file '$ConfigPath' not found"
    }
    $result = @{}
    Get-Content -Path $ConfigPath -ErrorAction Stop | Foreach-Object {
        if ($_.Trim() -eq '') { return } # Пропустить пустые строки
        if ($_ -notmatch '=') { return } # Пропустить строки без параметров
        if ($_.Split('=')[0] -notmatch "^;|#.*") { # Исключить закомментированные строки
            $Key, $Value = $_.Split('=', 2) # Разбить только по первому "=", чтобы не портить JSON-значения
            $result[$Key.Trim()] = $Value.Trim()
        }
    }
    #region Precheck Imported Params
    $RequiredParamsList = @(
        'TelegramBotToken'
        'TelegramChatId'
    )
    foreach ($RequiredParam in $RequiredParamsList) {
        if (-not $result.ContainsKey($RequiredParam)) {
            throw "Required parameter '$RequiredParam' is missing in '$ConfigPath'"
        }
    }
    #endregion Precheck Imported Params
    return $result
}
function ConvertTo-OrderedHashtable { # Конвертирует PSCustomObject (например, результат ConvertFrom-Json) в упорядоченный hashtable
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [PSCustomObject]$InputObject
    )
    $result = [ordered]@{}
    foreach ($Property in $InputObject.PSObject.Properties) {
        $result[$Property.Name] = $Property.Value
    }
    return $result
}
function Import-CustomRPOMap { # Парсит CustomRPOMap из JSON-строки в конфиге в массив объектов, ожидаемых Get-FormattedRPO
    param (
        [Parameter(Mandatory = $false, Position = 0)]
        [string]$CustomRPOMapJson
    )
    if ([string]::IsNullOrWhiteSpace($CustomRPOMapJson)) {
        return @()
    }
    try {
        $ParsedJson = ConvertFrom-Json -InputObject $CustomRPOMapJson
    }
    catch {
        throw "CustomRPOMap in config is not valid JSON: $($_.Exception.Message)"
    }
    $result = $ParsedJson | ForEach-Object {
        [PSCustomObject]@{
            JobName = $_.JobName
            RPOMap  = ConvertTo-OrderedHashtable -InputObject $_.RPOMap
        }
    }
    return @($result | Where-Object { $null -ne $_ })
}
function ConvertTo-TelegramHtmlText { # Экранирует спецсимволы HTML, иначе Telegram отбрасывает сообщение с parse_mode=html
    param (
        [Parameter(Mandatory = $false, Position = 0)]
        [AllowEmptyString()]
        [String]$Text = ''
    )
    $result = $Text.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
    return $result
}
function Split-MessageIntoChunks { # Разбивает блоки текста на части, каждая из которых влезает в одно сообщение Telegram
    param (
        [Parameter(Mandatory = $false, Position = 0)]
        [String[]]$TextBlocks = @(),
        [Parameter(Mandatory = $false, Position = 1)]
        [int]$ChunkSizeLimit = 3500 # Лимит Telegram - 4096 символов на сообщение, оставлен запас под заголовок и служебный текст
    )
    $Chunks = @()
    $CurrentChunk = ''
    foreach ($TextBlock in $TextBlocks) {
        if ($TextBlock.Length -gt $ChunkSizeLimit) {
            $TextBlock = ($TextBlock.Substring(0, $ChunkSizeLimit) -replace '&[A-Za-z]{0,4}$', '') # Блок длиннее лимита обрезается, недорезанная HTML-сущность в конце удаляется
        }
        if (($CurrentChunk -ne '') -and (($CurrentChunk.Length + $TextBlock.Length + 2) -gt $ChunkSizeLimit)) {
            $Chunks += $CurrentChunk
            $CurrentChunk = ''
        }
        if ($CurrentChunk -eq '') {
            $CurrentChunk = $TextBlock
        }
        else {
            $CurrentChunk = $CurrentChunk + "`n`n" + $TextBlock
        }
    }
    if ($CurrentChunk -ne '') {
        $Chunks += $CurrentChunk
    }
    return $Chunks
}
function Send-MessageToTelegramChatViaBot { # Отправляет сообщение в телеграм. Возвращает $true при успехе и $false, если отправить не удалось
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [String]$BotToken,
        [Parameter(Mandatory = $true, Position = 1)]
        [String]$ChatId,
        [Parameter(Mandatory = $true, Position = 2)]
        [String]$Message,
        [Parameter(Mandatory = $false, Position = 3)]
        [String]$ParseMode = 'html',
        [Parameter(Mandatory = $false, Position = 4)]
        [int]$MaxAttempts = 3
    )
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $Payload = [System.Text.Encoding]::UTF8.GetBytes(
        (@{
            chat_id    = $ChatId
            parse_mode = $ParseMode
            text       = $Message
        } | ConvertTo-Json) # POST с JSON-телом вместо текста в URL: спецсимволы, переносы строк и кириллица не требуют URL-кодирования
    )
    $i = [int]0
    do {
        $Failed = $false
        $i++
        try {
            Write-Host "Trying to send a message. Attempt number $i"
            $null = Invoke-RestMethod -Method Post -Uri "https://api.telegram.org/bot$($BotToken)/sendMessage" -ContentType 'application/json; charset=utf-8' -Body $Payload
        }
        catch {
            $Failed = $true
            $ErrorDescription = $_.Exception.Message
            if ($_.ErrorDetails.Message) {
                $ErrorDescription = $_.ErrorDetails.Message # В теле ответа Telegram описана причина отказа
            }
            Write-Host "Attempt $i failed: $ErrorDescription"
            if ($i -lt $MaxAttempts) {
                $SleepSeconds = 10
                if ($ErrorDescription -match '"retry_after"\s*:\s*(\d+)') { # При флуд-лимите Telegram сообщает, сколько нужно подождать
                    $SleepSeconds = [int]$Matches[1] + 1
                }
                Start-Sleep -Seconds $SleepSeconds
            }
        }
    }
    while (
        $Failed -and ($i -lt $MaxAttempts)
        )

    if ($Failed) {
        Write-Host "Failed to send message to Telegram after $MaxAttempts attempts"
        $result = $false
    }
    else {
        Write-Host 'Message sent successfully'
        $result = $true
    }
    return $result
}
function Get-FormattedDate { # Конвертирует дату из datetime в строку в удобночитаемом виде
    Param(
        [Parameter(Mandatory = $true, Position = 0)]
        $InputDate
    )
    if ($InputDate -is [datetime]) {
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
        [AllowEmptyString()]
        [String]$LastResult
    )
    $LastResultsMap = @{
        'Success' = 'Green'
        'Warning' = 'Yellow'
        'Failed'  = 'Red'
        'None'    = 'Yellow'
    }
    $Color = $LastResultsMap[$LastResult]
    if (-not $Color) { # Неизвестный статус считается проблемой
        $Color = 'Red'
    }
    $result = Add-EmojiAtTheBeginningOfTheString -Color $Color -String $LastResult
    return $result
}
function Get-VBRJobTotalBackupSize { # Рассчитывает суммарный размер всех Restore Points в рамках одной Backup Job. Возвращает размер в байтах или $null, если размер получить не удалось
    Param(
        [Parameter(Mandatory = $true, Position = 0)]
        $VBRBackupJob,
        [Parameter(Mandatory = $false, Position = 1)]
        [array]$VBRBackupList = @() # Результат Get-VBRBackup, получается один раз на весь отчет
    )
    switch ($VBRBackupJob.TypeToString) {
        'File Backup' {
            $result = try {
                $VBRBackupJob.FindLastSession().Info.BackupTotalSize
            }
            catch {
                $null
            }
        }
        'Backup Copy' {
            $result = try {
                (Get-VBRNASBackupCopyJob -Name $VBRBackupJob.Name).FindLastSession().Info.Progress.TotalUsedSize
            }
            catch {
                $null
            }
        }
        default {
            $result = try {
                $JobBackups = @($VBRBackupList | Where-Object { $_.JobId -eq $VBRBackupJob.Id }) # Поиск по JobId вместо маски "Имя*": маска захватывала чужие джобы, имена которых начинаются одинаково
                if ($JobBackups.Count -eq 0) {
                    $JobBackups = @($VBRBackupList | Where-Object { $_.Name -like "$($VBRBackupJob.Name)*" }) # Фолбек для бекапов, которые не находятся по JobId (например, дочерние бекапы агентских джоб)
                }
                $Sum = ($JobBackups | ForEach-Object { $_.GetAllStorages().Stats.BackupSize } | Measure-Object -Sum).Sum
                if ($null -eq $Sum) {
                    $Sum = 0
                }
                $Sum
            }
            catch {
                $null
            }
        }
    }
    return $result
}
function Add-EmojiAtTheBeginningOfTheString { #Добавляет Emoji в начало строки
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [String]$String,
        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateSet('Green', 'Yellow', 'Red')]
        [String]$Color
        )
    $EmojiMap = @{ # Символы заданы кодами, чтобы не зависеть от кодировки файла скрипта
        'Green'     = [string][char]0x2705                  # Галочка в зеленом квадрате
        'Yellow'    = [string][char]0x26A0 + [char]0xFE0F   # Предупреждение
        'Red'       = [string][char]0x274C                  # Красный крестик
    }
    $result = $EmojiMap.$Color + $String
    return $result
}
#endregion Functions
#Main Script
$Config = Import-Config -ConfigPath $ConfigPath
$CustomRPOMap = Import-CustomRPOMap -CustomRPOMapJson $Config.CustomRPOMap # Переопределение дефолтного RPO (24 часа) для некоторых видов бекапов, задаётся в конфиге
$DefaultRPOMap = [ordered]@{ # Максимально допустимое время в часах для получения зеленого, желтого или красного значка напротив значения RPO. Красный значек используется, если не выполняются условия для зеленого или желтого.
    '24'        = 'Green'
    '48'        = 'Yellow'
}
$DataCollectionError = $null
try {
    $AllVBRJobs = @(Get-VBRJob)
    $AllVBRBackups = @(Get-VBRBackup) # Получается один раз на весь отчет: вызовы Veeam медленные, дергать их на каждую джобу слишком дорого
}
catch {
    $DataCollectionError = $_.Exception.Message # Отчет с текстом ошибки все равно уходит в Telegram: молча умерший скрипт мониторинга хуже любой ошибки
    Write-Host "Failed to get data from Veeam: $DataCollectionError"
    $AllVBRJobs = @()
    $AllVBRBackups = @()
}
$BackupStatistics = foreach ($VBRJob in $AllVBRJobs) {
    $RPOMap = $DefaultRPOMap
    #region Custom RPO Settings
    foreach ($Element in $CustomRPOMap) {
        if ($VBRJob.Name -eq $Element.JobName) {
            $RPOMap = $Element.RPOMap
        }
    }
    #endregion Custom RPO Settings
    $LatestRestorePointDate = Get-VBRLatestRestorePointDate -VBRBackupJob $VBRJob # Получается один раз и используется и для расчета RPO, и для вывода в отчет
    $TotalBackupSize = Get-VBRJobTotalBackupSize -VBRBackupJob $VBRJob -VBRBackupList $AllVBRBackups
    if ($null -ne $TotalBackupSize) {
        $TotalBackupSizeFormatted = "$([math]::round($TotalBackupSize/1GB))GB"
    }
    else {
        $TotalBackupSizeFormatted = 'Unknown'
    }
    [PSCustomObject]@{
        'Name'                        = $VBRJob.Name
        'Job Type'                    = $VBRJob.TypeToString
        'Job status'                  = $VBRJob.GetLastState()
        'RPO'                         = Get-FormattedRPO -RPO (Get-VBRRecoveryPointObjective -LatestRestorePointDate $LatestRestorePointDate) -RPOMap $RPOMap
        'Last result'                 = Get-FormattedLastResult -LastResult ($VBRJob.Info.LatestStatus)
        'Latest restore point'        = Get-FormattedDate -InputDate $LatestRestorePointDate
        'Total backup size'           = $TotalBackupSizeFormatted
    }
}
$BackupStatistics
$Header  = 'Veeam backup report for ' + (Get-FormattedDate -InputDate (Get-Date))
$Tail    = '[DEBUG] Number of data processing errors: ' + $error.Count
$JobTextBlocks = @(
    $BackupStatistics | Sort-Object -Property 'Name' | ForEach-Object {
        ConvertTo-TelegramHtmlText -Text ($_ | Format-List | Out-String).Trim()
    }
)
$MessageChunks = @(Split-MessageIntoChunks -TextBlocks $JobTextBlocks) # Отчет длиннее лимита Telegram уходит несколькими сообщениями
if ($MessageChunks.Count -eq 0) {
    if ($DataCollectionError) {
        $MessageChunks = @(ConvertTo-TelegramHtmlText -Text "Failed to get data from Veeam: $DataCollectionError")
    }
    else {
        $MessageChunks = @('No backup jobs found')
    }
}
$SendFailed = $false
for ($ChunkIndex = 0; $ChunkIndex -lt $MessageChunks.Count; $ChunkIndex++) {
    $ChunkHeader = $Header
    if ($MessageChunks.Count -gt 1) {
        $ChunkHeader = "$Header (part $($ChunkIndex + 1)/$($MessageChunks.Count))"
    }
    $ChunkTail = ''
    if ($ChunkIndex -eq ($MessageChunks.Count - 1)) { # Строка со счетчиком ошибок добавляется только к последнему сообщению
        $ChunkTail = $Tail
    }
    $Message = $ChunkHeader + '<pre>' + "`n" + $MessageChunks[$ChunkIndex] + "`n" + '</pre>' + $ChunkTail
    if (-not (Send-MessageToTelegramChatViaBot -BotToken $Config.TelegramBotToken -ChatId $Config.TelegramChatId -Message $Message)) {
        $SendFailed = $true
    }
}
if ($SendFailed -or $DataCollectionError) {
    exit 1 # Ненулевой код возврата виден в планировщике задач как сбой
}
#Endregion Main Script
