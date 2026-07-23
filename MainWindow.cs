using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace NotePon;

internal sealed class MainWindow : Window
{
    private const int HotKeyScrollUpId = 1;
    private const int HotKeyScrollDownId = 2;
    private const int HotKeyVolumeUpId = 3;
    private const int HotKeyVolumeDownId = 4;
    private const int HotKeyVolumeMuteId = 5;
    private const int WmHotKey = 0x0312;
    private const int WmMouseActivate = 0x0021;
    private const int MaNoActivate = 3;
    private const int GwlExStyle = -20;
    private const long WsExNoActivate = 0x08000000L;
    private const uint ModAlt = 0x0001;
    private const uint ModControl = 0x0002;
    private const uint ModNoRepeat = 0x4000;
    private const uint VkF10 = 0x79;
    private const uint VkF11 = 0x7A;
    private const uint VkVolumeMute = 0xAD;
    private const uint VkVolumeDown = 0xAE;
    private const uint VkVolumeUp = 0xAF;
    private const double MinimumFontSize = 20;
    private const double MaximumFontSize = 96;
    private const double FontSizeStep = 4;
    private const double StandardScrollLineCount = 4;
    private const double VolumeScrollLineCount = 1;
    private const double ScrollAnimationMilliseconds = 160;

    private static readonly Brush WindowBackground = new SolidColorBrush(Color.FromRgb(0, 0, 0));
    private static readonly Brush StatusBackground = new SolidColorBrush(Color.FromRgb(25, 25, 25));
    private static readonly Brush ControlBackground = new SolidColorBrush(Color.FromRgb(50, 50, 50));
    private static readonly Brush ControlBorder = new SolidColorBrush(Color.FromRgb(105, 105, 105));

    private readonly PowerPointWorkerClient _powerPointWorker = new();
    private readonly CancellationTokenSource _shutdown = new();
    private readonly DispatcherTimer _pollTimer;
    private readonly DispatcherTimer _scrollTimer;
    private readonly TextBlock _statusText;
    private readonly TextBlock _slideText;
    private readonly StackPanel _notesPanel;
    private readonly ScrollViewer _notesScroller;
    private readonly TextBlock _moreIndicator;

    private HwndSource? _windowSource;
    private bool _scrollUpRegistered;
    private bool _scrollDownRegistered;
    private bool _volumeUpRegistered;
    private bool _volumeDownRegistered;
    private bool _volumeMuteRegistered;
    private bool _hotKeyRegistrationFailed;
    private PowerPointState? _lastState;
    private double _scrollStart;
    private double _scrollTarget;
    private long _scrollStartedAt;
    private bool _pollInProgress;
    private bool _hasDisplayedNotes;
    private double _noteFontSize = 44;
    private string _displayedBodyText = "PowerPointを待っています";
    private FormattedNoteDocument? _displayedFormattedNotes;

    public MainWindow()
    {
        Title = "NOTE-PON";
        Icon = BitmapFrame.Create(
            new Uri("pack://application:,,,/assets/note-pon.ico", UriKind.Absolute));
        ShowActivated = false;
        Focusable = false;
        Width = 1200;
        Height = 800;
        MinWidth = 640;
        MinHeight = 360;
        Background = WindowBackground;
        Foreground = Brushes.White;
        FontFamily = new FontFamily("Yu Gothic UI");
        WindowStartupLocation = WindowStartupLocation.CenterScreen;

        var root = new DockPanel();
        Content = root;

        var statusBar = new DockPanel
        {
            Background = StatusBackground,
            LastChildFill = true
        };
        DockPanel.SetDock(statusBar, Dock.Top);
        root.Children.Add(statusBar);

        var controls = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Margin = new Thickness(8)
        };
        DockPanel.SetDock(controls, Dock.Right);
        statusBar.Children.Add(controls);

        var smallerButton = CreateButton("文字を小さく");
        smallerButton.Click += (_, _) => ChangeFontSize(-FontSizeStep);
        controls.Children.Add(smallerButton);

        var largerButton = CreateButton("文字を大きく");
        largerButton.Click += (_, _) => ChangeFontSize(FontSizeStep);
        controls.Children.Add(largerButton);

        var topmostCheckBox = new CheckBox
        {
            Content = "常に手前",
            Focusable = false,
            Foreground = Brushes.White,
            VerticalAlignment = VerticalAlignment.Center,
            Padding = new Thickness(8, 5, 8, 5),
            Margin = new Thickness(4, 0, 0, 0)
        };
        topmostCheckBox.Checked += (_, _) => Topmost = true;
        topmostCheckBox.Unchecked += (_, _) => Topmost = false;
        controls.Children.Add(topmostCheckBox);

        _slideText = new TextBlock
        {
            Text = "- / -",
            FontSize = 16,
            FontWeight = FontWeights.SemiBold,
            Foreground = Brushes.White,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(12, 0, 16, 0)
        };
        DockPanel.SetDock(_slideText, Dock.Right);
        statusBar.Children.Add(_slideText);

        _statusText = new TextBlock
        {
            Text = "PowerPointを待っています",
            FontSize = 16,
            Foreground = Brushes.White,
            VerticalAlignment = VerticalAlignment.Center,
            TextTrimming = TextTrimming.CharacterEllipsis,
            Margin = new Thickness(16, 0, 8, 0)
        };
        statusBar.Children.Add(_statusText);

        _notesPanel = new StackPanel
        {
            Margin = new Thickness(32, 24, 32, 40)
        };

        _pollTimer = new DispatcherTimer(DispatcherPriority.Background)
        {
            Interval = TimeSpan.FromMilliseconds(350)
        };
        _pollTimer.Tick += OnPollTimerTick;

        _scrollTimer = new DispatcherTimer(DispatcherPriority.Render)
        {
            Interval = TimeSpan.FromMilliseconds(16)
        };
        _scrollTimer.Tick += (_, _) => AdvanceScrollAnimation();

        _notesScroller = new ScrollViewer
        {
            Content = _notesPanel,
            Focusable = false,
            IsHitTestVisible = false,
            Background = WindowBackground,
            VerticalScrollBarVisibility = ScrollBarVisibility.Hidden,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
            CanContentScroll = false,
            PanningMode = PanningMode.None
        };
        RenderCurrentBody();

        _moreIndicator = new TextBlock
        {
            Text = "▼",
            FontFamily = new FontFamily("Segoe UI Symbol"),
            FontSize = 30,
            Foreground = Brushes.White,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Bottom,
            Margin = new Thickness(0, 0, 0, 4),
            IsHitTestVisible = false,
            Focusable = false,
            Visibility = Visibility.Collapsed
        };

        _notesScroller.ScrollChanged += (_, _) =>
        {
            if (!_scrollTimer.IsEnabled)
            {
                _scrollTarget = _notesScroller.VerticalOffset;
            }

            UpdateMoreIndicator();
        };

        var notesLayer = new Grid();
        notesLayer.Children.Add(_notesScroller);
        notesLayer.Children.Add(_moreIndicator);
        root.Children.Add(notesLayer);

        SourceInitialized += OnSourceInitialized;
        PreviewKeyDown += (_, eventArgs) => eventArgs.Handled = true;
        Loaded += (_, _) =>
        {
            _pollTimer.Start();
            OnPollTimerTick(this, EventArgs.Empty);
        };
        Closed += OnClosed;
    }

    private static Button CreateButton(string text) =>
        new()
        {
            Content = text,
            Background = ControlBackground,
            Foreground = Brushes.White,
            BorderBrush = ControlBorder,
            Focusable = false,
            Padding = new Thickness(10, 5, 10, 5),
            Margin = new Thickness(4, 0, 0, 0),
            VerticalAlignment = VerticalAlignment.Center
        };

    private void OnSourceInitialized(object? sender, EventArgs eventArgs)
    {
        IntPtr handle = new WindowInteropHelper(this).Handle;
        MakeWindowNonActivating(handle);
        _windowSource = HwndSource.FromHwnd(handle);
        _windowSource?.AddHook(WindowProcedure);

        _scrollUpRegistered = RegisterHotKey(handle, HotKeyScrollUpId, ModControl | ModAlt, VkF10);
        _scrollDownRegistered = RegisterHotKey(handle, HotKeyScrollDownId, ModControl | ModAlt, VkF11);
        _volumeUpRegistered = RegisterHotKey(handle, HotKeyVolumeUpId, 0, VkVolumeUp);
        _volumeDownRegistered = RegisterHotKey(handle, HotKeyVolumeDownId, 0, VkVolumeDown);
        _volumeMuteRegistered = RegisterHotKey(handle, HotKeyVolumeMuteId, ModNoRepeat, VkVolumeMute);
        _hotKeyRegistrationFailed =
            !_scrollUpRegistered
            || !_scrollDownRegistered
            || !_volumeUpRegistered
            || !_volumeDownRegistered
            || !_volumeMuteRegistered;
        RefreshStatusText();
    }

    private static void MakeWindowNonActivating(IntPtr handle)
    {
        IntPtr currentStyle = GetWindowLongPtr(handle, GwlExStyle);
        IntPtr nonActivatingStyle = new(currentStyle.ToInt64() | WsExNoActivate);

        Marshal.SetLastPInvokeError(0);
        IntPtr previousStyle = SetWindowLongPtr(handle, GwlExStyle, nonActivatingStyle);
        int error = Marshal.GetLastPInvokeError();
        if (previousStyle == IntPtr.Zero && error != 0)
        {
            AppLog.Write(
                "The NOTE-PON window could not be marked as non-activating.",
                new System.ComponentModel.Win32Exception(error));
        }
    }

    private IntPtr WindowProcedure(IntPtr window, int message, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (message == WmMouseActivate)
        {
            handled = true;
            return new IntPtr(MaNoActivate);
        }

        if (message != WmHotKey)
        {
            return IntPtr.Zero;
        }

        int hotKeyId = wParam.ToInt32();
        if (hotKeyId == HotKeyScrollUpId)
        {
            BeginSmoothScroll(-1, StandardScrollLineCount);
            handled = true;
        }
        else if (hotKeyId == HotKeyScrollDownId)
        {
            BeginSmoothScroll(1, StandardScrollLineCount);
            handled = true;
        }
        else if (hotKeyId == HotKeyVolumeUpId)
        {
            BeginSmoothScroll(-1, VolumeScrollLineCount);
            handled = true;
        }
        else if (hotKeyId == HotKeyVolumeDownId)
        {
            BeginSmoothScroll(1, VolumeScrollLineCount);
            handled = true;
        }
        else if (hotKeyId == HotKeyVolumeMuteId)
        {
            BeginPageScroll();
            handled = true;
        }

        return IntPtr.Zero;
    }

    private async void OnPollTimerTick(object? sender, EventArgs eventArgs)
    {
        try
        {
            await UpdateFromPowerPointAsync();
        }
        catch (OperationCanceledException) when (_shutdown.IsCancellationRequested)
        {
            // The window is closing.
        }
        catch (Exception exception)
        {
            AppLog.Write(
                "The PowerPoint polling timer failed unexpectedly. NOTE-PON will shut down safely.",
                exception);
            Application.Current.Shutdown(1);
        }
    }

    private async Task UpdateFromPowerPointAsync()
    {
        if (_pollInProgress || _shutdown.IsCancellationRequested)
        {
            return;
        }

        _pollInProgress = true;
        try
        {
            PowerPointSnapshot snapshot = await _powerPointWorker.PollAsync(_shutdown.Token);
            if (_shutdown.IsCancellationRequested)
            {
                return;
            }

            _lastState = snapshot.State;
            RefreshStatusText(snapshot.StatusText);
            _slideText.Text = snapshot.SlideNumber.HasValue && snapshot.TotalSlides.HasValue
                ? $"{snapshot.SlideNumber} / {snapshot.TotalSlides}"
                : "- / -";

            if (snapshot.NotesChanged)
            {
                SetDisplayedBody(snapshot.Notes ?? string.Empty, snapshot.FormattedNotes);
                _hasDisplayedNotes = true;
                ResetScrollToTop();
                return;
            }

            if (!_hasDisplayedNotes
                && snapshot.State != PowerPointState.Connected
                && snapshot.State != _lastDisplayedBodyState)
            {
                SetDisplayedBody(snapshot.StatusText, null);
                ResetScrollToTop();
            }

            _lastDisplayedBodyState = snapshot.State;
        }
        finally
        {
            _pollInProgress = false;
        }
    }

    private PowerPointState? _lastDisplayedBodyState;

    private void RefreshStatusText(string? powerPointStatus = null)
    {
        string status = powerPointStatus ?? StateText(_lastState);
        _statusText.Text = _hotKeyRegistrationFailed
            ? $"{status} / ノートスクロールキーを登録できませんでした"
            : status;
    }

    private static string StateText(PowerPointState? state) => state switch
    {
        PowerPointState.WaitingForPresentation => "プレゼンテーションを待っています",
        PowerPointState.WaitingForSlideShow => "スライドショーを開始してください",
        PowerPointState.Connected => "PowerPoint 接続中",
        PowerPointState.Reconnecting => "PowerPointへ再接続しています",
        PowerPointState.NoteReadError => "ノートを読み取れません",
        _ => "PowerPointを待っています"
    };

    private void ChangeFontSize(double delta)
    {
        double currentOffset = _notesScroller.VerticalOffset;
        _noteFontSize = Math.Clamp(_noteFontSize + delta, MinimumFontSize, MaximumFontSize);
        RenderCurrentBody();

        Dispatcher.BeginInvoke(
            DispatcherPriority.Loaded,
            () =>
            {
                double restoredOffset = Math.Clamp(currentOffset, 0, _notesScroller.ScrollableHeight);
                _notesScroller.ScrollToVerticalOffset(restoredOffset);
                _scrollTarget = restoredOffset;
                UpdateMoreIndicator();
            });
    }

    private void BeginSmoothScroll(int direction, double lineCount)
    {
        BeginSmoothScrollBy(CurrentLineHeight * lineCount * direction);
    }

    private void BeginPageScroll()
    {
        double pageHeight = Math.Max(
            CurrentLineHeight,
            _notesScroller.ViewportHeight - CurrentLineHeight);
        BeginSmoothScrollBy(pageHeight);
    }

    private double CurrentLineHeight => Math.Round(_noteFontSize * 1.34);

    private void SetDisplayedBody(string text, FormattedNoteDocument? formattedNotes)
    {
        _displayedBodyText = text;
        _displayedFormattedNotes = formattedNotes;
        RenderCurrentBody();
    }

    private void RenderCurrentBody()
    {
        _notesPanel.Children.Clear();

        if (_displayedFormattedNotes?.Paragraphs is { Length: > 0 } paragraphs)
        {
            foreach (FormattedNoteParagraph paragraph in paragraphs)
            {
                _notesPanel.Children.Add(CreateFormattedParagraph(paragraph));
            }

            return;
        }

        var plainText = CreateNoteTextBlock(NoteTextAlignment.Left);
        plainText.Text = _displayedBodyText;
        _notesPanel.Children.Add(plainText);
    }

    private FrameworkElement CreateFormattedParagraph(FormattedNoteParagraph paragraph)
    {
        var paragraphText = CreateNoteTextBlock(paragraph.Alignment);
        foreach (FormattedNoteRun formattedRun in paragraph.Runs)
        {
            paragraphText.Inlines.Add(CreateFormattedRun(formattedRun));
        }

        if (paragraph.Runs.Length == 0)
        {
            paragraphText.Inlines.Add(new Run("\u00A0"));
        }

        double indent = Math.Max(0, paragraph.IndentLevel - 1) * _noteFontSize * 0.75;
        if (string.IsNullOrEmpty(paragraph.BulletText))
        {
            paragraphText.Margin = new Thickness(indent, 0, 0, 0);
            return paragraphText;
        }

        var bulletText = CreateNoteTextBlock(NoteTextAlignment.Left);
        bulletText.Text = paragraph.BulletText;
        bulletText.Margin = new Thickness(0, 0, _noteFontSize * 0.3, 0);

        var bulletRow = new Grid
        {
            Margin = new Thickness(indent, 0, 0, 0)
        };
        bulletRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        bulletRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        Grid.SetColumn(bulletText, 0);
        Grid.SetColumn(paragraphText, 1);
        bulletRow.Children.Add(bulletText);
        bulletRow.Children.Add(paragraphText);
        return bulletRow;
    }

    private TextBlock CreateNoteTextBlock(NoteTextAlignment alignment) =>
        new()
        {
            FontSize = _noteFontSize,
            LineHeight = CurrentLineHeight,
            LineStackingStrategy = LineStackingStrategy.BlockLineHeight,
            Foreground = Brushes.White,
            TextWrapping = TextWrapping.Wrap,
            TextAlignment = alignment switch
            {
                NoteTextAlignment.Center => TextAlignment.Center,
                NoteTextAlignment.Right => TextAlignment.Right,
                NoteTextAlignment.Justify => TextAlignment.Justify,
                _ => TextAlignment.Left
            },
            MinHeight = CurrentLineHeight
        };

    private Run CreateFormattedRun(FormattedNoteRun formattedRun)
    {
        var run = new Run(formattedRun.Text)
        {
            FontWeight = formattedRun.Bold ? FontWeights.Bold : FontWeights.Normal,
            FontStyle = formattedRun.Italic ? FontStyles.Italic : FontStyles.Normal,
            Foreground = CreateReadableForeground(formattedRun.RgbColor)
        };

        if (formattedRun.Underline)
        {
            run.TextDecorations = TextDecorations.Underline;
        }

        if (formattedRun.Superscript)
        {
            run.BaselineAlignment = BaselineAlignment.Superscript;
            run.FontSize = _noteFontSize * 0.72;
        }
        else if (formattedRun.Subscript)
        {
            run.BaselineAlignment = BaselineAlignment.Subscript;
            run.FontSize = _noteFontSize * 0.72;
        }

        return run;
    }

    private static Brush CreateReadableForeground(int? colorRef)
    {
        if (!colorRef.HasValue)
        {
            return Brushes.White;
        }

        int rgb = colorRef.Value;
        byte red = (byte)(rgb & 0xFF);
        byte green = (byte)((rgb >> 8) & 0xFF);
        byte blue = (byte)((rgb >> 16) & 0xFF);

        if (red == 0 && green == 0 && blue == 0)
        {
            return Brushes.White;
        }

        for (int step = 0; step <= 100; step++)
        {
            double blend = step / 100d;
            byte candidateRed = BlendTowardWhite(red, blend);
            byte candidateGreen = BlendTowardWhite(green, blend);
            byte candidateBlue = BlendTowardWhite(blue, blend);
            if (ContrastAgainstBlack(candidateRed, candidateGreen, candidateBlue) >= 4.5)
            {
                var brush = new SolidColorBrush(
                    Color.FromRgb(candidateRed, candidateGreen, candidateBlue));
                brush.Freeze();
                return brush;
            }
        }

        return Brushes.White;
    }

    private static byte BlendTowardWhite(byte component, double blend) =>
        (byte)Math.Round(component + ((255 - component) * blend));

    private static double ContrastAgainstBlack(byte red, byte green, byte blue) =>
        (RelativeLuminance(red, green, blue) + 0.05) / 0.05;

    private static double RelativeLuminance(byte red, byte green, byte blue) =>
        (0.2126 * LinearizeColor(red))
        + (0.7152 * LinearizeColor(green))
        + (0.0722 * LinearizeColor(blue));

    private static double LinearizeColor(byte component)
    {
        double channel = component / 255d;
        return channel <= 0.04045
            ? channel / 12.92
            : Math.Pow((channel + 0.055) / 1.055, 2.4);
    }

    private void BeginSmoothScrollBy(double amount)
    {
        double currentOffset = _notesScroller.VerticalOffset;

        _scrollStart = currentOffset;
        _scrollTarget = Math.Clamp(_scrollTarget + amount, 0, _notesScroller.ScrollableHeight);
        _scrollStartedAt = Stopwatch.GetTimestamp();
        _scrollTimer.Start();
    }

    private void AdvanceScrollAnimation()
    {
        double elapsedMilliseconds = Stopwatch.GetElapsedTime(_scrollStartedAt).TotalMilliseconds;
        double progress = Math.Clamp(elapsedMilliseconds / ScrollAnimationMilliseconds, 0, 1);
        double easedProgress = 1 - Math.Pow(1 - progress, 3);
        double offset = _scrollStart + ((_scrollTarget - _scrollStart) * easedProgress);
        _notesScroller.ScrollToVerticalOffset(offset);

        if (progress >= 1)
        {
            _notesScroller.ScrollToVerticalOffset(_scrollTarget);
            _scrollTimer.Stop();
        }
    }

    private void UpdateMoreIndicator()
    {
        bool hasMore =
            _notesScroller.ScrollableHeight > 0 &&
            _notesScroller.VerticalOffset < _notesScroller.ScrollableHeight - 0.5;
        _moreIndicator.Visibility = hasMore ? Visibility.Visible : Visibility.Collapsed;
    }

    private void ResetScrollToTop()
    {
        _scrollTimer.Stop();
        _scrollStart = 0;
        _scrollTarget = 0;
        _notesScroller.ScrollToTop();
    }

    private void OnClosed(object? sender, EventArgs eventArgs)
    {
        _shutdown.Cancel();
        _pollTimer.Stop();
        _scrollTimer.Stop();
        _powerPointWorker.Dispose();

        IntPtr handle = new WindowInteropHelper(this).Handle;
        if (_scrollUpRegistered)
        {
            UnregisterHotKey(handle, HotKeyScrollUpId);
        }

        if (_scrollDownRegistered)
        {
            UnregisterHotKey(handle, HotKeyScrollDownId);
        }

        if (_volumeUpRegistered)
        {
            UnregisterHotKey(handle, HotKeyVolumeUpId);
        }

        if (_volumeDownRegistered)
        {
            UnregisterHotKey(handle, HotKeyVolumeDownId);
        }

        if (_volumeMuteRegistered)
        {
            UnregisterHotKey(handle, HotKeyVolumeMuteId);
        }

        _windowSource?.RemoveHook(WindowProcedure);
        _windowSource = null;
    }

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool RegisterHotKey(IntPtr window, int id, uint modifiers, uint virtualKey);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnregisterHotKey(IntPtr window, int id);

    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW", SetLastError = true)]
    private static extern IntPtr GetWindowLongPtr(IntPtr window, int index);

    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW", SetLastError = true)]
    private static extern IntPtr SetWindowLongPtr(IntPtr window, int index, IntPtr newValue);

}
