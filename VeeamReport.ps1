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
function Get-VBRBackupForJob { # Находит объект бекапа, принадлежащий заданию. NAS-задания живут в отдельном списке Get-VBRNASBackup
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        $VBRBackupJob,
        [Parameter(Mandatory = $false, Position = 1)]
        [array]$VBRBackupList = @(),
        [Parameter(Mandatory = $false, Position = 2)]
        [array]$VBRNASBackupList = @()
    )
    $SearchList = $VBRBackupList
    if ($VBRBackupJob.JobType -in 'NasBackup', 'NasBackupCopy') {
        $SearchList = $VBRNASBackupList
    }
    $result = @($SearchList | Where-Object { $_.JobId -eq $VBRBackupJob.Id }) | Select-Object -First 1
    if (-not $result) {
        $result = @($SearchList | Where-Object { $_.Name -eq $VBRBackupJob.Name }) | Select-Object -First 1 # Фолбек по имени: у части бекапов JobId не заполнен
    }
    return $result
}
function Get-VBRLatestRestorePointDate { # Получает дату последней точки восстановления по самим точкам, а не по метаданным задания
    param (
        [Parameter(Mandatory = $false, Position = 0)]
        $VBRBackup, # Объект бекапа из Get-VBRBackup или Get-VBRNASBackup
        [Parameter(Mandatory = $false, Position = 1)]
        $VBRBackupJob # Задание нужно только для фолбеков по метаданным
    )
    # Точки восстановления - единственный надёжный источник даты. Метаданные задания врут:
    # LastPointCreationTime у многих заданий пустой, а MetaUpdateTime у Backup Copy в объектное
    # хранилище остаётся на дате создания задания и даёт RPO в тысячи часов при живых бекапах.
    if ($VBRBackup) {
        $IsNASBackup = ($VBRBackup.PSObject.Properties.Name -contains 'LastRestorePointCreationTime') # Свойство есть только у VBRNASBackup, у обычного CBackup его нет
        $RestorePoints = @()
        try {
            if (-not $IsNASBackup) {
                $RestorePoints = @(Get-VBRRestorePoint -Backup $VBRBackup -ErrorAction Stop)
            }
            elseif (Get-Command -Name 'Get-VBRUnstructuredBackupRestorePoint' -ErrorAction SilentlyContinue) {
                $RestorePoints = @(Get-VBRUnstructuredBackupRestorePoint -Backup $VBRBackup -ErrorAction Stop)
            }
            else {
                $RestorePoints = @(Get-VBRNASBackupRestorePoint -NASBackup $VBRBackup -ErrorAction Stop) # Фолбек для Veeam, где ещё нет Unstructured-командлетов
            }
        }
        catch {
            $RestorePoints = @()
        }
        $LatestPoint = $RestorePoints | Sort-Object -Property 'CreationTime' -Descending | Select-Object -First 1
        if ($LatestPoint -and $LatestPoint.CreationTime) {
            try {
                return ConvertTo-DateTime -InputDate $LatestPoint.CreationTime
            }
            catch {
            }
        }
        foreach ($FallbackProperty in 'LastRestorePointCreationTime', 'LastPointCreationTime') { # Фолбек для версий Veeam, где командлеты точек недоступны
            if ($VBRBackup.$FallbackProperty) {
                try {
                    return ConvertTo-DateTime -InputDate $VBRBackup.$FallbackProperty
                }
                catch {
                }
            }
        }
    }
    if ($VBRBackupJob) {
        try {
            return ConvertTo-DateTime -InputDate $VBRBackupJob.FindLastSession().EndTime # Последний фолбек: время завершения последней сессии
        }
        catch {
        }
    }
    return 'No restore points'
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
function Get-RPOMapForName { # Подбирает пороги RPO для конкретного задания: кастомные, если имя совпало, иначе дефолтные
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [AllowEmptyString()]
        [string]$Name,
        [Parameter(Mandatory = $true, Position = 1)]
        $DefaultRPOMap,
        [Parameter(Mandatory = $false, Position = 2)]
        [array]$CustomRPOMap = @()
    )
    foreach ($Element in $CustomRPOMap) {
        if ($Name -eq $Element.JobName) {
            return $Element.RPOMap
        }
    }
    return $DefaultRPOMap
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
function Get-VBRBackupTotalSize { # Считает суммарный размер бекапа в байтах. Возвращает $null, если размер получить не удалось
    Param(
        [Parameter(Mandatory = $false, Position = 0)]
        $VBRBackup,
        [Parameter(Mandatory = $false, Position = 1)]
        $VBRBackupJob
    )
    if ($VBRBackupJob -and $VBRBackupJob.JobType -in 'NasBackup', 'NasBackupCopy') {
        # У NAS-заданий размер берётся из последней сессии: у VBRNASBackup нет ни storages, ни свойств размера.
        # Именно BackupTotalSize, а не Progress.TotalUsedSize: последний равен ProcessedSize, то есть объёму,
        # обработанному за конкретный запуск. На инкрементальном запуске копии он давал 507GB вместо 1750GB.
        try {
            return $VBRBackupJob.FindLastSession().Info.BackupTotalSize
        }
        catch {
            return $null
        }
    }
    if (-not $VBRBackup) {
        return $null
    }
    # GetAllStorages() возвращает 0 почти для всех бекапов (в том числе для всех в объектном
    # хранилище) - реальные данные лежат в дочерних бекапах, поэтому основной источник GetAllChildrenStorages()
    foreach ($Method in 'GetAllChildrenStorages', 'GetAllStorages') {
        try {
            $Sum = (($VBRBackup.$Method().Stats.BackupSize) | Measure-Object -Sum).Sum
            if ($Sum) {
                return $Sum
            }
        }
        catch {
        }
    }
    return 0
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
$IncludeBackupsWithoutJob = $true # Бекапы, у которых не осталось задания (например, агентские), иначе просто исчезают из отчёта вместе со своим RPO
if ($Config.ContainsKey('IncludeBackupsWithoutJob')) {
    $IncludeBackupsWithoutJob = ($Config.IncludeBackupsWithoutJob -notmatch '^(false|0|no)$')
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
$AllVBRNASBackups = @()
try {
    # NAS-бекапы в Get-VBRBackup не попадают, у них свой список. Get-VBRNASBackup объявлен
    # устаревшим в пользу Get-VBRUnstructuredBackup и на каждый вызов пишет warning
    if (Get-Command -Name 'Get-VBRUnstructuredBackup' -ErrorAction SilentlyContinue) {
        $AllVBRNASBackups = @(Get-VBRUnstructuredBackup)
    }
    else {
        $AllVBRNASBackups = @(Get-VBRNASBackup)
    }
}
catch {
    Write-Host "Failed to get NAS backups: $($_.Exception.Message)"
}
$BackupStatistics = @()
$BackupStatistics += foreach ($VBRJob in $AllVBRJobs) {
    $VBRBackup = Get-VBRBackupForJob -VBRBackupJob $VBRJob -VBRBackupList $AllVBRBackups -VBRNASBackupList $AllVBRNASBackups
    $LatestRestorePointDate = Get-VBRLatestRestorePointDate -VBRBackup $VBRBackup -VBRBackupJob $VBRJob # Получается один раз и используется и для расчета RPO, и для вывода в отчет
    $TotalBackupSize = Get-VBRBackupTotalSize -VBRBackup $VBRBackup -VBRBackupJob $VBRJob
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
        'RPO'                         = Get-FormattedRPO -RPO (Get-VBRRecoveryPointObjective -LatestRestorePointDate $LatestRestorePointDate) -RPOMap (Get-RPOMapForName -Name $VBRJob.Name -DefaultRPOMap $DefaultRPOMap -CustomRPOMap $CustomRPOMap)
        'Last result'                 = Get-FormattedLastResult -LastResult ($VBRJob.Info.LatestStatus)
        'Latest restore point'        = Get-FormattedDate -InputDate $LatestRestorePointDate
        'Total backup size'           = $TotalBackupSizeFormatted
    }
}
#region Backups without a job
if ($IncludeBackupsWithoutJob) {
    $JobIds = @($AllVBRJobs | ForEach-Object { [string]$_.Id })
    $JobNames = @($AllVBRJobs | ForEach-Object { $_.Name })
    $OrphanedBackups = @($AllVBRBackups | Where-Object { ([string]$_.JobId) -notin $JobIds -and $_.Name -notin $JobNames })
    $BackupStatistics += foreach ($VBRBackup in $OrphanedBackups) {
        $LatestRestorePointDate = Get-VBRLatestRestorePointDate -VBRBackup $VBRBackup
        $TotalBackupSize = Get-VBRBackupTotalSize -VBRBackup $VBRBackup
        if ($null -ne $TotalBackupSize) {
            $TotalBackupSizeFormatted = "$([math]::round($TotalBackupSize/1GB))GB"
        }
        else {
            $TotalBackupSizeFormatted = 'Unknown'
        }
        [PSCustomObject]@{
            'Name'                        = $VBRBackup.Name
            'Job Type'                    = [string]$VBRBackup.TypeToString
            'Job status'                  = 'No job'
            'RPO'                         = Get-FormattedRPO -RPO (Get-VBRRecoveryPointObjective -LatestRestorePointDate $LatestRestorePointDate) -RPOMap (Get-RPOMapForName -Name $VBRBackup.Name -DefaultRPOMap $DefaultRPOMap -CustomRPOMap $CustomRPOMap)
            'Last result'                 = Get-FormattedLastResult -LastResult 'Unknown'
            'Latest restore point'        = Get-FormattedDate -InputDate $LatestRestorePointDate
            'Total backup size'           = $TotalBackupSizeFormatted
        }
    }
}
#endregion Backups without a job
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
