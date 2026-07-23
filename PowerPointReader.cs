using System.Runtime.InteropServices;

namespace NotePon;

internal enum PowerPointState
{
    WaitingForPowerPoint,
    WaitingForPresentation,
    WaitingForSlideShow,
    Connected,
    Reconnecting,
    NoteReadError
}

internal sealed record PowerPointSnapshot(
    PowerPointState State,
    string StatusText,
    int? SlideNumber = null,
    int? TotalSlides = null,
    bool NotesChanged = false,
    string? Notes = null,
    FormattedNoteDocument? FormattedNotes = null);

internal sealed class PowerPointReader
{
    private const int MsoShapeTypePlaceholder = 14;
    private const int PpPlaceholderBody = 2;
    private const int PpPlaceholderVerticalBody = 6;
    private const int PpBulletUnnumbered = 1;
    private const int PpBulletNumbered = 2;
    private const int PpBulletPicture = 3;
    private const int PpAlignLeft = 1;
    private const int PpAlignCenter = 2;
    private const int PpAlignRight = 3;
    private const int PpAlignJustify = 4;
    private const int PpAlignDistribute = 5;
    private const int PpAlignThaiDistribute = 6;
    private const int PpAlignJustifyLow = 7;

    private bool _hasConnected;
    private string? _lastSlideKey;

    public PowerPointSnapshot Poll()
    {
        object? application = null;
        object? presentations = null;
        object? slideShowWindows = null;
        object? slideShowWindow = null;
        object? view = null;
        object? slide = null;
        object? presentation = null;
        object? slides = null;

        try
        {
            if (!TryGetRunningPowerPoint(out application))
            {
                ResetCurrentSlide();
                return _hasConnected
                    ? Snapshot(PowerPointState.Reconnecting, "PowerPointへ再接続しています")
                    : Snapshot(PowerPointState.WaitingForPowerPoint, "PowerPointを待っています");
            }

            _hasConnected = true;
            dynamic powerPoint = application!;
            presentations = powerPoint.Presentations;
            if ((int)((dynamic)presentations).Count == 0)
            {
                ResetCurrentSlide();
                return Snapshot(PowerPointState.WaitingForPresentation, "プレゼンテーションを待っています");
            }

            slideShowWindows = powerPoint.SlideShowWindows;
            if ((int)((dynamic)slideShowWindows).Count == 0)
            {
                ResetCurrentSlide();
                return Snapshot(PowerPointState.WaitingForSlideShow, "スライドショーを開始してください");
            }

            slideShowWindow = ((dynamic)slideShowWindows).Item(1);
            view = ((dynamic)slideShowWindow).View;
            slide = ((dynamic)view).Slide;
            presentation = ((dynamic)slideShowWindow).Presentation;
            slides = ((dynamic)presentation).Slides;

            int slideNumber = (int)((dynamic)slide).SlideIndex;
            int totalSlides = (int)((dynamic)slides).Count;
            string presentationKey = GetPresentationKey(presentation);
            string slideKey = $"{presentationKey}\0{slideNumber}";

            if (slideKey == _lastSlideKey)
            {
                return Snapshot(
                    PowerPointState.Connected,
                    "PowerPoint 接続中",
                    slideNumber,
                    totalSlides);
            }

            SpeakerNotesResult speakerNotes;
            try
            {
                speakerNotes = ReadSpeakerNotes(slide);
            }
            catch (Exception exception)
            {
                AppLog.WriteThrottled(
                    "note-read-error",
                    "PowerPoint speaker notes could not be read.",
                    exception);
                _lastSlideKey = null;
                return Snapshot(
                    PowerPointState.NoteReadError,
                    "ノートを読み取れません",
                    slideNumber,
                    totalSlides);
            }

            _lastSlideKey = slideKey;
            string notes = speakerNotes.PlainText;
            if (string.IsNullOrWhiteSpace(notes))
            {
                notes = "このスライドにはノートがありません";
                speakerNotes = speakerNotes with { FormattedNotes = null };
            }

            return Snapshot(
                PowerPointState.Connected,
                "PowerPoint 接続中",
                slideNumber,
                totalSlides,
                NotesChanged: true,
                Notes: notes,
                FormattedNotes: speakerNotes.FormattedNotes);
        }
        catch (Exception exception)
        {
            AppLog.WriteThrottled(
                "powerpoint-read-error",
                "PowerPoint state could not be read.",
                exception);
            ResetCurrentSlide();
            return _hasConnected
                ? Snapshot(PowerPointState.Reconnecting, "PowerPointへ再接続しています")
                : Snapshot(PowerPointState.WaitingForPowerPoint, "PowerPointを待っています");
        }
        finally
        {
            ReleaseComObject(slides);
            ReleaseComObject(presentation);
            ReleaseComObject(slide);
            ReleaseComObject(view);
            ReleaseComObject(slideShowWindow);
            ReleaseComObject(slideShowWindows);
            ReleaseComObject(presentations);
            ReleaseComObject(application);
        }
    }

    private static PowerPointSnapshot Snapshot(
        PowerPointState state,
        string statusText,
        int? slideNumber = null,
        int? totalSlides = null,
        bool NotesChanged = false,
        string? Notes = null,
        FormattedNoteDocument? FormattedNotes = null) =>
        new(state, statusText, slideNumber, totalSlides, NotesChanged, Notes, FormattedNotes);

    private static bool TryGetRunningPowerPoint(out object? application)
    {
        application = null;
        int classIdResult = CLSIDFromProgID("PowerPoint.Application", out Guid classId);
        if (classIdResult < 0)
        {
            return false;
        }

        int activeObjectResult = GetActiveObject(ref classId, IntPtr.Zero, out application);
        return activeObjectResult >= 0 && application is not null;
    }

    private static string GetPresentationKey(object presentation)
    {
        try
        {
            return (string)((dynamic)presentation).FullName;
        }
        catch
        {
            return (string)((dynamic)presentation).Name;
        }
    }

    private static SpeakerNotesResult ReadSpeakerNotes(object slide)
    {
        object? notesPage = null;
        object? shapes = null;

        try
        {
            notesPage = ((dynamic)slide).NotesPage;
            shapes = ((dynamic)notesPage).Shapes;
            int shapeCount = (int)((dynamic)shapes).Count;

            for (int index = 1; index <= shapeCount; index++)
            {
                object? shape = null;
                object? placeholderFormat = null;
                object? textFrame = null;
                object? textRange = null;

                try
                {
                    shape = ((dynamic)shapes).Item(index);
                    if ((int)((dynamic)shape).Type != MsoShapeTypePlaceholder)
                    {
                        continue;
                    }

                    placeholderFormat = ((dynamic)shape).PlaceholderFormat;
                    int placeholderType = (int)((dynamic)placeholderFormat).Type;
                    if (placeholderType is not PpPlaceholderBody and not PpPlaceholderVerticalBody)
                    {
                        continue;
                    }

                    textFrame = ((dynamic)shape).TextFrame;
                    textRange = ((dynamic)textFrame).TextRange;
                    string text = (string)((dynamic)textRange).Text;
                    string plainText = NormalizeLineEndings(text);
                    FormattedNoteDocument? formattedNotes = null;

                    try
                    {
                        formattedNotes = ReadFormattedNotes(textRange);
                    }
                    catch (Exception exception)
                    {
                        AppLog.WriteThrottled(
                            "note-format-read-error",
                            "PowerPoint speaker note formatting could not be read. Plain text will be used.",
                            exception);
                    }

                    return new SpeakerNotesResult(plainText, formattedNotes);
                }
                finally
                {
                    ReleaseComObject(textRange);
                    ReleaseComObject(textFrame);
                    ReleaseComObject(placeholderFormat);
                    ReleaseComObject(shape);
                }
            }

            return new SpeakerNotesResult(string.Empty, null);
        }
        finally
        {
            ReleaseComObject(shapes);
            ReleaseComObject(notesPage);
        }
    }

    private static FormattedNoteDocument ReadFormattedNotes(object textRange)
    {
        object? allParagraphs = null;

        try
        {
            allParagraphs = ((dynamic)textRange).Paragraphs();
            int paragraphCount = (int)((dynamic)allParagraphs).Count;
            var paragraphs = new List<FormattedNoteParagraph>(paragraphCount);

            for (int paragraphIndex = 1; paragraphIndex <= paragraphCount; paragraphIndex++)
            {
                object? paragraphRange = null;
                object? paragraphFormat = null;
                object? bulletFormat = null;

                try
                {
                    paragraphRange = ((dynamic)textRange).Paragraphs(paragraphIndex, 1);
                    paragraphFormat = ((dynamic)paragraphRange).ParagraphFormat;
                    bulletFormat = ((dynamic)paragraphFormat).Bullet;

                    var runs = ReadFormattedRuns(paragraphRange);
                    int alignment = (int)((dynamic)paragraphFormat).Alignment;
                    int indentLevel = Math.Clamp(
                        (int)((dynamic)paragraphRange).IndentLevel,
                        1,
                        5);
                    string? bulletText = ReadBulletText(bulletFormat);

                    paragraphs.Add(
                        new FormattedNoteParagraph(
                            runs,
                            MapAlignment(alignment),
                            indentLevel,
                            bulletText));
                }
                finally
                {
                    ReleaseComObject(bulletFormat);
                    ReleaseComObject(paragraphFormat);
                    ReleaseComObject(paragraphRange);
                }
            }

            return new FormattedNoteDocument(paragraphs.ToArray());
        }
        finally
        {
            ReleaseComObject(allParagraphs);
        }
    }

    private static FormattedNoteRun[] ReadFormattedRuns(object paragraphRange)
    {
        object? allRuns = null;

        try
        {
            allRuns = ((dynamic)paragraphRange).Runs();
            int runCount = (int)((dynamic)allRuns).Count;
            var runs = new List<FormattedNoteRun>(runCount);

            for (int runIndex = 1; runIndex <= runCount; runIndex++)
            {
                object? runRange = null;
                object? font = null;
                object? color = null;

                try
                {
                    runRange = ((dynamic)paragraphRange).Runs(runIndex, 1);
                    font = ((dynamic)runRange).Font;
                    color = ((dynamic)font).Color;

                    string text = NormalizeRunText((string)((dynamic)runRange).Text);
                    if (text.Length == 0)
                    {
                        continue;
                    }

                    runs.Add(
                        new FormattedNoteRun(
                            text,
                            IsMsoTrue(((dynamic)font).Bold),
                            IsMsoTrue(((dynamic)font).Italic),
                            IsMsoTrue(((dynamic)font).Underline),
                            IsMsoTrue(((dynamic)font).Superscript),
                            IsMsoTrue(((dynamic)font).Subscript),
                            ReadRgbColor(color)));
                }
                finally
                {
                    ReleaseComObject(color);
                    ReleaseComObject(font);
                    ReleaseComObject(runRange);
                }
            }

            return runs.ToArray();
        }
        finally
        {
            ReleaseComObject(allRuns);
        }
    }

    private static string? ReadBulletText(object bulletFormat)
    {
        int bulletType = (int)((dynamic)bulletFormat).Type;
        return bulletType switch
        {
            PpBulletUnnumbered => "•",
            PpBulletNumbered => $"{(int)((dynamic)bulletFormat).Number}.",
            PpBulletPicture => "•",
            _ => null
        };
    }

    private static NoteTextAlignment MapAlignment(int alignment) =>
        alignment switch
        {
            PpAlignCenter => NoteTextAlignment.Center,
            PpAlignRight => NoteTextAlignment.Right,
            PpAlignJustify
                or PpAlignDistribute
                or PpAlignThaiDistribute
                or PpAlignJustifyLow => NoteTextAlignment.Justify,
            PpAlignLeft or _ => NoteTextAlignment.Left
        };

    private static bool IsMsoTrue(object value) =>
        Convert.ToInt32(value) == -1;

    private static int? ReadRgbColor(object color)
    {
        int rgb = (int)((dynamic)color).RGB;
        return rgb < 0 ? null : rgb & 0x00FFFFFF;
    }

    private static string NormalizeRunText(string text) =>
        text.Replace("\r\n", "\n", StringComparison.Ordinal)
            .TrimEnd('\r')
            .Replace('\r', '\n')
            .Replace('\v', '\n');

    private static string NormalizeLineEndings(string text) =>
        text.Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n')
            .Replace('\v', '\n')
            .Replace("\n", Environment.NewLine, StringComparison.Ordinal);

    private void ResetCurrentSlide()
    {
        _lastSlideKey = null;
    }

    private sealed record SpeakerNotesResult(
        string PlainText,
        FormattedNoteDocument? FormattedNotes);

    private static void ReleaseComObject(object? value)
    {
        if (value is null || !Marshal.IsComObject(value))
        {
            return;
        }

        try
        {
            Marshal.ReleaseComObject(value);
        }
        catch
        {
            // PowerPoint may have exited while its COM objects were being released.
        }
    }

    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int CLSIDFromProgID(string progId, out Guid classId);

    [DllImport("oleaut32.dll", PreserveSig = true)]
    private static extern int GetActiveObject(
        ref Guid classId,
        IntPtr reserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object? application);
}
