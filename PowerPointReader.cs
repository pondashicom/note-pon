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
    string? Notes = null);

internal sealed class PowerPointReader
{
    private const int MsoShapeTypePlaceholder = 14;
    private const int PpPlaceholderBody = 2;
    private const int PpPlaceholderVerticalBody = 6;

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

            string notes;
            try
            {
                notes = ReadSpeakerNotes(slide);
            }
            catch
            {
                _lastSlideKey = null;
                return Snapshot(
                    PowerPointState.NoteReadError,
                    "ノートを読み取れません",
                    slideNumber,
                    totalSlides);
            }

            _lastSlideKey = slideKey;
            if (string.IsNullOrWhiteSpace(notes))
            {
                notes = "このスライドにはノートがありません";
            }

            return Snapshot(
                PowerPointState.Connected,
                "PowerPoint 接続中",
                slideNumber,
                totalSlides,
                NotesChanged: true,
                Notes: notes);
        }
        catch
        {
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
        string? Notes = null) =>
        new(state, statusText, slideNumber, totalSlides, NotesChanged, Notes);

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

    private static string ReadSpeakerNotes(object slide)
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
                    return NormalizeLineEndings(text);
                }
                finally
                {
                    ReleaseComObject(textRange);
                    ReleaseComObject(textFrame);
                    ReleaseComObject(placeholderFormat);
                    ReleaseComObject(shape);
                }
            }

            return string.Empty;
        }
        finally
        {
            ReleaseComObject(shapes);
            ReleaseComObject(notesPage);
        }
    }

    private static string NormalizeLineEndings(string text) =>
        text.Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n')
            .Replace('\v', '\n')
            .Replace("\n", Environment.NewLine, StringComparison.Ordinal);

    private void ResetCurrentSlide()
    {
        _lastSlideKey = null;
    }

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
