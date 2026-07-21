using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;

namespace NotePon;

internal sealed class MainWindow : Window
{
    private const int HotKeyScrollUpId = 1;
    private const int HotKeyScrollDownId = 2;
    private const int WmHotKey = 0x0312;
    private const uint ModAlt = 0x0001;
    private const uint ModControl = 0x0002;
    private const uint VkF10 = 0x79;
    private const uint VkF11 = 0x7A;
    private const double MinimumFontSize = 20;
    private const double MaximumFontSize = 96;
    private const double FontSizeStep = 4;
    private const double ScrollLineCount = 4;
    private const double ScrollAnimationMilliseconds = 160;

    private static readonly Brush WindowBackground = new SolidColorBrush(Color.FromRgb(0, 0, 0));
    private static readonly Brush StatusBackground = new SolidColorBrush(Color.FromRgb(25, 25, 25));
    private static readonly Brush ControlBackground = new SolidColorBrush(Color.FromRgb(50, 50, 50));
    private static readonly Brush ControlBorder = new SolidColorBrush(Color.FromRgb(105, 105, 105));

    private readonly PowerPointReader _powerPointReader = new();
    private readonly DispatcherTimer _pollTimer;
    private readonly DispatcherTimer _scrollTimer;
    private readonly TextBlock _statusText;
    private readonly TextBlock _slideText;
    private readonly TextBlock _notesText;
    private readonly ScrollViewer _notesScroller;

    private HwndSource? _windowSource;
    private bool _scrollUpRegistered;
    private bool _scrollDownRegistered;
    private bool _hotKeyRegistrationFailed;
    private PowerPointState? _lastState;
    private double _scrollStart;
    private double _scrollTarget;
    private long _scrollStartedAt;

    public MainWindow()
    {
        Title = "NOTE-PON";
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

        _notesText = new TextBlock
        {
            Text = "PowerPointを待っています",
            FontSize = 44,
            LineHeight = 59,
            LineStackingStrategy = LineStackingStrategy.BlockLineHeight,
            Foreground = Brushes.White,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(32, 24, 32, 40)
        };

        _pollTimer = new DispatcherTimer(DispatcherPriority.Background)
        {
            Interval = TimeSpan.FromMilliseconds(350)
        };
        _pollTimer.Tick += (_, _) => UpdateFromPowerPoint();

        _scrollTimer = new DispatcherTimer(DispatcherPriority.Render)
        {
            Interval = TimeSpan.FromMilliseconds(16)
        };
        _scrollTimer.Tick += (_, _) => AdvanceScrollAnimation();

        _notesScroller = new ScrollViewer
        {
            Content = _notesText,
            Background = WindowBackground,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
            CanContentScroll = false,
            PanningMode = PanningMode.VerticalOnly
        };
        _notesScroller.ScrollChanged += (_, _) =>
        {
            if (!_scrollTimer.IsEnabled)
            {
                _scrollTarget = _notesScroller.VerticalOffset;
            }
        };
        root.Children.Add(_notesScroller);

        SourceInitialized += OnSourceInitialized;
        Loaded += (_, _) =>
        {
            UpdateFromPowerPoint();
            _pollTimer.Start();
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
            Padding = new Thickness(10, 5, 10, 5),
            Margin = new Thickness(4, 0, 0, 0),
            VerticalAlignment = VerticalAlignment.Center
        };

    private void OnSourceInitialized(object? sender, EventArgs eventArgs)
    {
        IntPtr handle = new WindowInteropHelper(this).Handle;
        _windowSource = HwndSource.FromHwnd(handle);
        _windowSource?.AddHook(WindowProcedure);

        _scrollUpRegistered = RegisterHotKey(handle, HotKeyScrollUpId, ModControl | ModAlt, VkF10);
        _scrollDownRegistered = RegisterHotKey(handle, HotKeyScrollDownId, ModControl | ModAlt, VkF11);
        _hotKeyRegistrationFailed = !_scrollUpRegistered || !_scrollDownRegistered;
        RefreshStatusText();
    }

    private IntPtr WindowProcedure(IntPtr window, int message, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (message != WmHotKey)
        {
            return IntPtr.Zero;
        }

        int hotKeyId = wParam.ToInt32();
        if (hotKeyId == HotKeyScrollUpId)
        {
            BeginSmoothScroll(-1);
            handled = true;
        }
        else if (hotKeyId == HotKeyScrollDownId)
        {
            BeginSmoothScroll(1);
            handled = true;
        }

        return IntPtr.Zero;
    }

    private void UpdateFromPowerPoint()
    {
        PowerPointSnapshot snapshot = _powerPointReader.Poll();
        _lastState = snapshot.State;
        RefreshStatusText(snapshot.StatusText);
        _slideText.Text = snapshot.SlideNumber.HasValue && snapshot.TotalSlides.HasValue
            ? $"{snapshot.SlideNumber} / {snapshot.TotalSlides}"
            : "- / -";

        if (snapshot.NotesChanged)
        {
            _notesText.Text = snapshot.Notes ?? string.Empty;
            ResetScrollToTop();
            return;
        }

        if (snapshot.State != PowerPointState.Connected && snapshot.State != _lastDisplayedBodyState)
        {
            _notesText.Text = snapshot.StatusText;
            ResetScrollToTop();
        }

        _lastDisplayedBodyState = snapshot.State;
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
        double newSize = Math.Clamp(_notesText.FontSize + delta, MinimumFontSize, MaximumFontSize);
        _notesText.FontSize = newSize;
        _notesText.LineHeight = Math.Round(newSize * 1.34);
    }

    private void BeginSmoothScroll(int direction)
    {
        double lineHeight = double.IsNaN(_notesText.LineHeight)
            ? _notesText.FontSize * 1.34
            : _notesText.LineHeight;
        double amount = lineHeight * ScrollLineCount * direction;
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

    private void ResetScrollToTop()
    {
        _scrollTimer.Stop();
        _scrollStart = 0;
        _scrollTarget = 0;
        _notesScroller.ScrollToTop();
    }

    private void OnClosed(object? sender, EventArgs eventArgs)
    {
        _pollTimer.Stop();
        _scrollTimer.Stop();

        IntPtr handle = new WindowInteropHelper(this).Handle;
        if (_scrollUpRegistered)
        {
            UnregisterHotKey(handle, HotKeyScrollUpId);
        }

        if (_scrollDownRegistered)
        {
            UnregisterHotKey(handle, HotKeyScrollDownId);
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
}
