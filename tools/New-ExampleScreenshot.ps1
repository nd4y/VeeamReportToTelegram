<#
.SYNOPSIS
    Генерирует example.png для README - синтетический пример отчёта в оформлении Telegram.

.DESCRIPTION
    Картинка рисуется, а не снимается с реального чата: на скриншоте настоящего отчёта видны
    имена заданий, хостов и репозиториев, а публиковать их в репозитории незачем. Заодно
    нарисованная картинка не устаревает молча - при изменении формата отчёта достаточно
    поправить массив $ReportLines и перегенерировать.

    Требуется Windows: рисование идёт через System.Drawing (GDI+).

.PARAMETER OutputPath
    Куда записать PNG. По умолчанию - example.png в корне репозитория.

.EXAMPLE
    .\tools\New-ExampleScreenshot.ps1
#>
param (
    [Parameter(Position = 0, Mandatory = $false)]
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

if (-not $OutputPath) { # Значение по умолчанию считается здесь, а не в param: в Windows PowerShell 5.1 $PSScriptRoot там ещё не заполнен
    $ScriptRoot = $PSScriptRoot
    if (-not $ScriptRoot) { $ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path }
    $OutputPath = Join-Path (Split-Path -Parent $ScriptRoot) 'example.png'
}
try {
    Add-Type -AssemblyName System.Drawing
}
catch {
    Add-Type -AssemblyName System.Drawing.Common # В PowerShell 7 сборка называется иначе
}

#region Содержимое отчёта
# Данные вымышленные. Набор и порядок полей должны совпадать с тем, что формирует VeeamReport.ps1,
# иначе README будет показывать формат, которого уже нет. Показаны все три цвета значков,
# чтобы легенда читалась без запуска скрипта.
$HeaderText = 'Veeam backup report for 12 May 2026 08:30'
$TailText   = '[DEBUG] Number of data processing errors: 0'
$TimeText   = '08:30'
$ReportLines = @(
    @{ Text = 'Name                 : Production VMware Backup' }
    @{ Text = 'Job Type             : VMware Backup' }
    @{ Text = 'Job status           : Stopped' }
    @{ Text = 'RPO                  : '; Icon = 'Green';  Value = '8h' }
    @{ Text = 'Last result          : '; Icon = 'Green';  Value = 'Success' }
    @{ Text = 'Latest restore point : 12 May 2026 00:15' }
    @{ Text = 'Total backup size    : 2480GB' }
    @{ Text = 'Repository           : main-repo-01' }
    @{ Text = '' }
    @{ Text = 'Name                 : File Share Backup' }
    @{ Text = 'Job Type             : File Backup' }
    @{ Text = 'Job status           : Stopped' }
    @{ Text = 'RPO                  : '; Icon = 'Yellow'; Value = '31h' }
    @{ Text = 'Last result          : '; Icon = 'Yellow'; Value = 'Warning' }
    @{ Text = 'Latest restore point : 10 May 2026 23:40' }
    @{ Text = 'Total backup size    : 870GB' }
    @{ Text = 'Repository           : main-repo-01' }
    @{ Text = '' }
    @{ Text = 'Name                 : Backup Copy to object storage' }
    @{ Text = 'Job Type             : Backup Copy' }
    @{ Text = 'Job status           : Working' }
    @{ Text = 'RPO                  : '; Icon = 'Green';  Value = '9h' }
    @{ Text = 'Last result          : '; Icon = 'Green';  Value = 'Success' }
    @{ Text = 'Latest restore point : 12 May 2026 00:15' }
    @{ Text = 'Total backup size    : 5310GB' }
    @{ Text = 'Repository           : object-storage-01' }
    @{ Text = '' }
    @{ Text = 'Name                 : Legacy SQL Backup' }
    @{ Text = 'Job Type             : Hyper-V Backup' }
    @{ Text = 'Job status           : Stopped' }
    @{ Text = 'RPO                  : '; Icon = 'Red';    Value = '97h' }
    @{ Text = 'Last result          : '; Icon = 'Red';    Value = 'Failed' }
    @{ Text = 'Latest restore point : 8 May 2026 02:00' }
    @{ Text = 'Total backup size    : 145GB' }
    @{ Text = 'Repository           : main-repo-01' }
    @{ Text = '' }
    @{ Text = '--- Backups without a job ---' }
    @{ Text = '' }
    @{ Text = 'Name                 : agent-node-01' }
    @{ Text = 'Backup Type          : Linux Agent Backup' }
    @{ Text = 'Latest restore point : 14 February 2026 22:00' }
    @{ Text = 'Total backup size    : 12GB' }
    @{ Text = 'Repository           : main-repo-01' }
)
#endregion Содержимое отчёта

#region Оформление
$ColorPage   = [System.Drawing.ColorTranslator]::FromHtml('#0E1621') # Фон чата, тёмная тема Telegram
$ColorBubble = [System.Drawing.ColorTranslator]::FromHtml('#182533') # Пузырь сообщения
$ColorText   = [System.Drawing.ColorTranslator]::FromHtml('#E4EBF0')
$ColorMuted  = [System.Drawing.ColorTranslator]::FromHtml('#6B7F8F') # Время в углу пузыря
$ColorGreen  = [System.Drawing.ColorTranslator]::FromHtml('#3FBF54')
$ColorYellow = [System.Drawing.ColorTranslator]::FromHtml('#F2B317')
$ColorRed    = [System.Drawing.ColorTranslator]::FromHtml('#F1453D')
$ColorSign   = [System.Drawing.ColorTranslator]::FromHtml('#3A2E00') # Восклицательный знак внутри жёлтого треугольника

$FontMono   = New-Object System.Drawing.Font('Consolas', 12, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)
$FontHeader = New-Object System.Drawing.Font('Segoe UI', 12.5, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)
$FontTime   = New-Object System.Drawing.Font('Segoe UI', 8.5, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)

$PagePad   = 14 # Поля вокруг пузыря
$BubblePad = 16 # Поля внутри пузыря
$Radius    = 12 # Скругление углов пузыря
$IconGap   = 4  # Отступ между значком и текстом справа от него
#endregion Оформление

function Add-RoundedRectangle { # Скруглённый прямоугольник: в GraphicsPath готовой фигуры нет
    param (
        [Parameter(Mandatory = $true)]
        [System.Drawing.Drawing2D.GraphicsPath]$Path,
        [Parameter(Mandatory = $true)]
        [single]$X,
        [Parameter(Mandatory = $true)]
        [single]$Y,
        [Parameter(Mandatory = $true)]
        [single]$Width,
        [Parameter(Mandatory = $true)]
        [single]$Height,
        [Parameter(Mandatory = $true)]
        [single]$CornerRadius
    )
    $d = $CornerRadius * 2
    $Path.AddArc($X, $Y, $d, $d, 180, 90)
    $Path.AddArc($X + $Width - $d, $Y, $d, $d, 270, 90)
    $Path.AddArc($X + $Width - $d, $Y + $Height - $d, $d, $d, 0, 90)
    $Path.AddArc($X, $Y + $Height - $d, $d, $d, 90, 90)
    $Path.CloseFigure()
}

function Add-StatusIcon { # Значки рисуются фигурами, а не шрифтом: GDI+ отдаёт emoji монохромными, и зелёная галочка вышла бы серой
    param (
        [Parameter(Mandatory = $true)]
        [System.Drawing.Graphics]$Graphics,
        [Parameter(Mandatory = $true)]
        [ValidateSet('Green', 'Yellow', 'Red')]
        [string]$Kind,
        [Parameter(Mandatory = $true)]
        [single]$X,
        [Parameter(Mandatory = $true)]
        [single]$Y,
        [Parameter(Mandatory = $true)]
        [single]$Size
    )
    switch ($Kind) {
        'Green' { # Белая галочка в зелёном квадрате
            $Path = New-Object System.Drawing.Drawing2D.GraphicsPath
            Add-RoundedRectangle -Path $Path -X $X -Y $Y -Width $Size -Height $Size -CornerRadius 3
            $Graphics.FillPath((New-Object System.Drawing.SolidBrush $ColorGreen), $Path)
            $Pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::White), ([single]($Size * 0.16))
            $Pen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
            $Pen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
            # Массив приводится к PointF[] явно: иначе PowerShell 7 передаёт Object[] и не может выбрать перегрузку
            $Graphics.DrawLines($Pen, [System.Drawing.PointF[]]@(
                (New-Object System.Drawing.PointF (($X + $Size * 0.24), ($Y + $Size * 0.52)))
                (New-Object System.Drawing.PointF (($X + $Size * 0.43), ($Y + $Size * 0.70)))
                (New-Object System.Drawing.PointF (($X + $Size * 0.77), ($Y + $Size * 0.31)))
            ))
        }
        'Yellow' { # Восклицательный знак в треугольнике
            $Graphics.FillPolygon((New-Object System.Drawing.SolidBrush $ColorYellow), [System.Drawing.PointF[]]@(
                (New-Object System.Drawing.PointF (($X + $Size * 0.50), ($Y + $Size * 0.06)))
                (New-Object System.Drawing.PointF (($X + $Size * 0.98), ($Y + $Size * 0.90)))
                (New-Object System.Drawing.PointF (($X + $Size * 0.02), ($Y + $Size * 0.90)))
            ))
            $Pen = New-Object System.Drawing.Pen $ColorSign, ([single]($Size * 0.13))
            $Pen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
            $Pen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
            $Graphics.DrawLine($Pen, ($X + $Size * 0.5), ($Y + $Size * 0.36), ($X + $Size * 0.5), ($Y + $Size * 0.62))
            $Graphics.FillEllipse((New-Object System.Drawing.SolidBrush $ColorSign), ($X + $Size * 0.435), ($Y + $Size * 0.70), ($Size * 0.13), ($Size * 0.13))
        }
        'Red' { # Красный крестик
            $Pen = New-Object System.Drawing.Pen $ColorRed, ([single]($Size * 0.19))
            $Pen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
            $Pen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
            $Graphics.DrawLine($Pen, ($X + $Size * 0.17), ($Y + $Size * 0.17), ($X + $Size * 0.83), ($Y + $Size * 0.83))
            $Graphics.DrawLine($Pen, ($X + $Size * 0.83), ($Y + $Size * 0.17), ($X + $Size * 0.17), ($Y + $Size * 0.83))
        }
    }
}

# GenericTypographic вместо формата по умолчанию: тот добавляет поля вокруг строки и режет
# хвостовые пробелы, из-за чего значок уезжал бы от двоеточия
$StringFormat = [System.Drawing.StringFormat]::GenericTypographic
$StringFormat.FormatFlags = $StringFormat.FormatFlags -bor [System.Drawing.StringFormatFlags]::MeasureTrailingSpaces

#region Замер содержимого - размер картинки подгоняется под текст, а не задаётся руками
$ProbeBitmap = New-Object System.Drawing.Bitmap 1, 1
$ProbeGraphics = [System.Drawing.Graphics]::FromImage($ProbeBitmap)
$ProbeGraphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit

function Measure-LineWidth {
    param ([string]$Text, [System.Drawing.Font]$Font)
    if ([string]::IsNullOrEmpty($Text)) { return [single]0 }
    return $ProbeGraphics.MeasureString($Text, $Font, [int]10000, $StringFormat).Width
}

$IconSize   = [int][Math]::Round($FontMono.GetHeight($ProbeGraphics) * 0.85)
$LineHeight = [int][Math]::Round($FontMono.GetHeight($ProbeGraphics) * 1.32)

$ContentWidth = [Math]::Max((Measure-LineWidth $HeaderText $FontHeader), (Measure-LineWidth $TailText $FontHeader))
foreach ($Line in $ReportLines) {
    $Width = Measure-LineWidth $Line.Text $FontMono
    if ($Line.Icon) { $Width += $IconSize + $IconGap + (Measure-LineWidth $Line.Value $FontMono) }
    if ($Width -gt $ContentWidth) { $ContentWidth = $Width }
}

$HeaderHeight = [int][Math]::Round($FontHeader.GetHeight($ProbeGraphics) * 1.6)
$TailHeight   = [int][Math]::Round($FontHeader.GetHeight($ProbeGraphics) * 2.0)
$TimeHeight   = $FontTime.GetHeight($ProbeGraphics)
$ProbeGraphics.Dispose()
$ProbeBitmap.Dispose()

$BubbleWidth  = [int][Math]::Ceiling($ContentWidth) + $BubblePad * 2
$BubbleHeight = $BubblePad * 2 + $HeaderHeight + ($ReportLines.Count * $LineHeight) + $TailHeight
$ImageWidth   = $BubbleWidth + $PagePad * 2
$ImageHeight  = $BubbleHeight + $PagePad * 2
#endregion Замер содержимого

$Bitmap = New-Object System.Drawing.Bitmap $ImageWidth, $ImageHeight
$Graphics = [System.Drawing.Graphics]::FromImage($Bitmap)
$Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$Graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
$Graphics.Clear($ColorPage)

$BubblePath = New-Object System.Drawing.Drawing2D.GraphicsPath
Add-RoundedRectangle -Path $BubblePath -X $PagePad -Y $PagePad -Width $BubbleWidth -Height $BubbleHeight -CornerRadius $Radius
$Graphics.FillPath((New-Object System.Drawing.SolidBrush $ColorBubble), $BubblePath)

$BrushText = New-Object System.Drawing.SolidBrush $ColorText
$BrushMuted = New-Object System.Drawing.SolidBrush $ColorMuted

$TextLeft = $PagePad + $BubblePad
$TextTop = $PagePad + $BubblePad
$Graphics.DrawString($HeaderText, $FontHeader, $BrushText, [single]$TextLeft, [single]$TextTop, $StringFormat)
$TextTop += $HeaderHeight

foreach ($Line in $ReportLines) {
    if ($Line.Text) {
        $Graphics.DrawString($Line.Text, $FontMono, $BrushText, [single]$TextLeft, [single]$TextTop, $StringFormat)
        if ($Line.Icon) {
            $PrefixWidth = $Graphics.MeasureString($Line.Text, $FontMono, [int]10000, $StringFormat).Width
            $IconTop = $TextTop + ($FontMono.GetHeight($Graphics) - $IconSize) / 2
            Add-StatusIcon -Graphics $Graphics -Kind $Line.Icon -X ([single]($TextLeft + $PrefixWidth)) -Y ([single]$IconTop) -Size ([single]$IconSize)
            $Graphics.DrawString($Line.Value, $FontMono, $BrushText, [single]($TextLeft + $PrefixWidth + $IconSize + $IconGap), [single]$TextTop, $StringFormat)
        }
    }
    $TextTop += $LineHeight
}

$TextTop += [int]($TailHeight * 0.15)
$Graphics.DrawString($TailText, $FontHeader, $BrushText, [single]$TextLeft, [single]$TextTop, $StringFormat)

$TimeWidth = $Graphics.MeasureString($TimeText, $FontTime, [int]10000, $StringFormat).Width
$Graphics.DrawString($TimeText, $FontTime, $BrushMuted,
    [single]($PagePad + $BubbleWidth - $BubblePad - $TimeWidth),
    [single]($PagePad + $BubbleHeight - $BubblePad - $TimeHeight), $StringFormat)

$Bitmap.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
$Graphics.Dispose()
$Bitmap.Dispose()
Write-Host "Written $OutputPath ($ImageWidth x $ImageHeight)"
