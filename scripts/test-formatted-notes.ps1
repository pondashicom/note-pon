[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$projectPath = Join-Path $repositoryRoot 'NotePon.csproj'
$executablePath = Join-Path $repositoryRoot 'bin\Release\net10.0-windows\NOTE-PON.exe'
$screenshotPath = Join-Path $env:TEMP 'note-pon-formatted-notes.png'

function Release-ComObject {
    param([object]$Value)

    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

function Assert-True {
    param(
        [bool]$Condition,
        [string]$Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

function Find-TextStart {
    param(
        [string]$Text,
        [string]$Needle
    )

    $index = $Text.IndexOf($Needle, [StringComparison]::Ordinal)
    if ($index -lt 0) {
        throw "Test text '$Needle' was not found."
    }

    return $index + 1
}

function Set-TextFontProperty {
    param(
        [object]$TextRange,
        [string]$FullText,
        [string]$Needle,
        [string]$PropertyName,
        [object]$Value
    )

    $characters = $null
    $font = $null
    try {
        $start = Find-TextStart -Text $FullText -Needle $Needle
        $characters = $TextRange.Characters($start, $Needle.Length)
        $font = $characters.Font
        $font.$PropertyName = $Value
    }
    finally {
        Release-ComObject $font
        Release-ComObject $characters
    }
}

function Get-NotesBodyShape {
    param([object]$Slide)

    $notesPage = $null
    $shapes = $null
    try {
        $notesPage = $Slide.NotesPage
        $shapes = $notesPage.Shapes
        for ($index = 1; $index -le $shapes.Count; $index++) {
            $shape = $null
            $placeholderFormat = $null
            $isBodyShape = $false
            try {
                $shape = $shapes.Item($index)
                if ($shape.Type -ne 14) {
                    continue
                }

                $placeholderFormat = $shape.PlaceholderFormat
                if ($placeholderFormat.Type -eq 2) {
                    $isBodyShape = $true
                    return $shape
                }
            }
            finally {
                Release-ComObject $placeholderFormat
                if ($null -ne $shape -and -not $isBodyShape) {
                    Release-ComObject $shape
                }
            }
        }
    }
    finally {
        Release-ComObject $shapes
        Release-ComObject $notesPage
    }

    throw 'The notes body placeholder was not found.'
}

function Capture-Window {
    param(
        [IntPtr]$WindowHandle,
        [string]$Path
    )

    Add-Type -AssemblyName System.Drawing
    $previousCaptureLib = $env:LIB
    try {
        $env:LIB = ''
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class NotePonFormattingScreenshot
{
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr window, out RECT rect);

    [DllImport("user32.dll")]
    public static extern bool PrintWindow(IntPtr window, IntPtr deviceContext, uint flags);
}
'@
    }
    finally {
        $env:LIB = $previousCaptureLib
    }

    $rect = New-Object NotePonFormattingScreenshot+RECT
    Assert-True `
        ([NotePonFormattingScreenshot]::GetWindowRect($WindowHandle, [ref]$rect)) `
        'GetWindowRect failed.'

    $width = $rect.Right - $rect.Left
    $height = $rect.Bottom - $rect.Top
    $bitmap = New-Object Drawing.Bitmap $width, $height
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    try {
        $deviceContext = $graphics.GetHdc()
        try {
            Assert-True `
                ([NotePonFormattingScreenshot]::PrintWindow($WindowHandle, $deviceContext, 2)) `
                'PrintWindow failed.'
        }
        finally {
            $graphics.ReleaseHdc($deviceContext)
        }

        $bitmap.Save($Path, [Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
    throw 'Close PowerPoint before running this isolated formatting test.'
}

$previousTaskLib = $env:LIB
try {
    $env:LIB = ''
    & dotnet build $projectPath --configuration Release --no-incremental -warnaserror
    if ($LASTEXITCODE -ne 0) {
        throw "NOTE-PON build failed with exit code $LASTEXITCODE."
    }
}
finally {
    $env:LIB = $previousTaskLib
}

$powerPoint = $null
$presentation = $null
$slide = $null
$notesBody = $null
$textFrame = $null
$textRange = $null
$paragraph = $null
$paragraphFormat = $null
$bullet = $null
$slideShowWindow = $null
$worker = $null
$notePon = $null

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.Visible = -1
    $presentation = $powerPoint.Presentations.Add()
    $slide = $presentation.Slides.Add(1, 12)
    $notesBody = Get-NotesBodyShape -Slide $slide
    $textFrame = $notesBody.TextFrame
    $textRange = $textFrame.TextRange

    $noteText =
        "Plain Bold Italic Underline Red H2O x2`r" +
        "Indented bullet`r" +
        "`r" +
        "Centered paragraph"
    $textRange.Text = $noteText

    Set-TextFontProperty $textRange $noteText 'Bold' 'Bold' -1
    Set-TextFontProperty $textRange $noteText 'Italic' 'Italic' -1
    Set-TextFontProperty $textRange $noteText 'Underline' 'Underline' -1

    $redCharacters = $null
    $redFont = $null
    $redColor = $null
    try {
        $redStart = Find-TextStart -Text $noteText -Needle 'Red'
        $redCharacters = $textRange.Characters($redStart, 3)
        $redFont = $redCharacters.Font
        $redColor = $redFont.Color
        $redColor.RGB = 255
    }
    finally {
        Release-ComObject $redColor
        Release-ComObject $redFont
        Release-ComObject $redCharacters
    }

    $subscriptCharacters = $null
    $subscriptFont = $null
    $superscriptCharacters = $null
    $superscriptFont = $null
    try {
        $subscriptStart = (Find-TextStart -Text $noteText -Needle 'H2O') + 1
        $subscriptCharacters = $textRange.Characters($subscriptStart, 1)
        $subscriptFont = $subscriptCharacters.Font
        $subscriptFont.Subscript = -1

        $superscriptStart = (Find-TextStart -Text $noteText -Needle 'x2') + 1
        $superscriptCharacters = $textRange.Characters($superscriptStart, 1)
        $superscriptFont = $superscriptCharacters.Font
        $superscriptFont.Superscript = -1
    }
    finally {
        Release-ComObject $superscriptFont
        Release-ComObject $superscriptCharacters
        Release-ComObject $subscriptFont
        Release-ComObject $subscriptCharacters
    }

    $paragraph = $textRange.Paragraphs(2, 1)
    $paragraph.IndentLevel = 2
    $paragraphFormat = $paragraph.ParagraphFormat
    $bullet = $paragraphFormat.Bullet
    $bullet.Visible = -1
    $bullet.Character = 8226
    Release-ComObject $bullet
    $bullet = $null
    Release-ComObject $paragraphFormat
    $paragraphFormat = $null
    Release-ComObject $paragraph
    $paragraph = $null

    $paragraph = $textRange.Paragraphs(4, 1)
    $paragraphFormat = $paragraph.ParagraphFormat
    $paragraphFormat.Alignment = 2
    Release-ComObject $paragraphFormat
    $paragraphFormat = $null
    Release-ComObject $paragraph
    $paragraph = $null

    $slideShowWindow = $presentation.SlideShowSettings.Run()

    $startInfo = New-Object Diagnostics.ProcessStartInfo
    $startInfo.FileName = $executablePath
    $startInfo.Arguments = "--powerpoint-worker --parent-pid $PID"
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardInput = $true
    $startInfo.RedirectStandardOutput = $true
    $worker = [Diagnostics.Process]::Start($startInfo)
    $worker.StandardInput.WriteLine('poll')
    $worker.StandardInput.Flush()
    $readTask = $worker.StandardOutput.ReadLineAsync()
    Assert-True ($readTask.Wait(5000)) 'The formatting worker did not answer.'
    $snapshot = $readTask.Result | ConvertFrom-Json

    Assert-True ($snapshot.State -eq 3) 'The formatting worker did not connect to the slide show.'
    Assert-True ($snapshot.FormattedNotes.Paragraphs.Count -ge 4) 'Formatted paragraphs were not returned.'
    $runs = @($snapshot.FormattedNotes.Paragraphs | ForEach-Object { $_.Runs })
    Assert-True (@($runs | Where-Object Bold).Count -gt 0) 'Bold formatting was not returned.'
    Assert-True (@($runs | Where-Object Italic).Count -gt 0) 'Italic formatting was not returned.'
    Assert-True (@($runs | Where-Object Underline).Count -gt 0) 'Underline formatting was not returned.'
    Assert-True (@($runs | Where-Object Superscript).Count -gt 0) 'Superscript formatting was not returned.'
    Assert-True (@($runs | Where-Object Subscript).Count -gt 0) 'Subscript formatting was not returned.'
    Assert-True (@($runs | Where-Object { $_.RgbColor -eq 255 }).Count -gt 0) 'Red text was not returned.'
    Assert-True ($snapshot.FormattedNotes.Paragraphs[1].IndentLevel -eq 2) 'Paragraph indentation was not returned.'
    Assert-True ([bool]$snapshot.FormattedNotes.Paragraphs[1].BulletText) 'The bullet marker was not returned.'
    Assert-True ($snapshot.FormattedNotes.Paragraphs[2].Runs.Count -eq 0) 'The blank paragraph was not preserved.'
    Assert-True ($snapshot.FormattedNotes.Paragraphs[3].Alignment -eq 1) 'Center alignment was not returned.'

    $worker.Kill()
    $worker.WaitForExit(3000)
    $worker.Dispose()
    $worker = $null

    $notePon = Start-Process -FilePath $executablePath -PassThru
    $windowDeadline = [DateTime]::UtcNow.AddSeconds(8)
    do {
        Start-Sleep -Milliseconds 100
        $notePon.Refresh()
    } while ($notePon.MainWindowHandle -eq [IntPtr]::Zero -and
        [DateTime]::UtcNow -lt $windowDeadline)

    Assert-True ($notePon.MainWindowHandle -ne [IntPtr]::Zero) 'The NOTE-PON window did not appear.'
    Start-Sleep -Milliseconds 1200
    Capture-Window -WindowHandle $notePon.MainWindowHandle -Path $screenshotPath

    [pscustomobject]@{
        Result = 'PASS'
        ParagraphCount = $snapshot.FormattedNotes.Paragraphs.Count
        RunCount = $runs.Count
        Screenshot = $screenshotPath
    }
}
finally {
    if ($worker -and -not $worker.HasExited) {
        $worker.Kill()
        $worker.WaitForExit(3000)
    }
    if ($worker) {
        $worker.Dispose()
    }

    if ($notePon -and -not $notePon.HasExited) {
        [void]$notePon.CloseMainWindow()
        if (-not $notePon.WaitForExit(3000)) {
            $notePon.Kill()
            $notePon.WaitForExit(3000)
        }
    }
    if ($notePon) {
        $notePon.Dispose()
    }

    if ($slideShowWindow) {
        try {
            $slideShowWindow.View.Exit()
        }
        catch {
        }
    }
    if ($presentation) {
        try {
            $presentation.Close()
        }
        catch {
        }
    }
    if ($powerPoint) {
        try {
            $powerPoint.Quit()
        }
        catch {
        }
    }

    Release-ComObject $bullet
    Release-ComObject $paragraphFormat
    Release-ComObject $paragraph
    Release-ComObject $slideShowWindow
    Release-ComObject $textRange
    Release-ComObject $textFrame
    Release-ComObject $notesBody
    Release-ComObject $slide
    Release-ComObject $presentation
    Release-ComObject $powerPoint
}
