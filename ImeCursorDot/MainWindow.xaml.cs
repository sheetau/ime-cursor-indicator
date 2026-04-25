using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Windows.Threading;
using Forms = System.Windows.Forms;
using Media = System.Windows.Media;

namespace ImeCursorDot
{
    public partial class MainWindow : Window
    {
        private readonly DispatcherTimer _imeTimer;
        private Thread? _posThread;
        private volatile bool _posThreadRunning = false;

        private readonly Window _dotWindow;
        private readonly Ellipse _dotEllipse;
        private IntPtr _dotHwnd = IntPtr.Zero;

        private readonly Forms.NotifyIcon _notifyIcon;

        private IntPtr _lastFgWnd = IntPtr.Zero;
        private bool _lastImeOpen = false;

        public MainWindow()
        {
            InitializeComponent();
            ShowInTaskbar = false;
            Hide();

            var ctx = new Forms.ContextMenuStrip();

            const int DOT_SIZE = 7;
            const int WIN_SIZE = DOT_SIZE + 2;

            _dotEllipse = new Ellipse
            {
                Width = DOT_SIZE,
                Height = DOT_SIZE,
                Fill = Media.Brushes.White,
                Stroke = Media.Brushes.Black,
                StrokeThickness = 1,
                Visibility = Visibility.Collapsed,
                SnapsToDevicePixels = true,
                UseLayoutRounding = true
            };
            RenderOptions.SetBitmapScalingMode(_dotEllipse, BitmapScalingMode.NearestNeighbor);

            AddColorItem(ctx, "White", Media.Brushes.White, Media.Brushes.Black, true);
            AddColorItem(ctx, "Neon Green",
                new Media.SolidColorBrush(Media.Color.FromRgb(57, 255, 20)),
                new Media.SolidColorBrush(Media.Color.FromRgb(0, 120, 0)),
                false);
            AddColorItem(ctx, "Red", Media.Brushes.Red,
                new Media.SolidColorBrush(Media.Color.FromRgb(120, 0, 0)), false);
            AddColorItem(ctx, "Orange", Media.Brushes.Orange,
                new Media.SolidColorBrush(Media.Color.FromRgb(140, 70, 0)), false);
            AddColorItem(ctx, "Magenta", Media.Brushes.Magenta,
                new Media.SolidColorBrush(Media.Color.FromRgb(120, 0, 120)), false);
            AddColorItem(ctx, "Cyan", Media.Brushes.Cyan,
                new Media.SolidColorBrush(Media.Color.FromRgb(0, 120, 120)), false);

            ctx.Items.Add(new Forms.ToolStripSeparator());

            var exitItem = new Forms.ToolStripMenuItem("Exit");
            exitItem.Click += (_, __) => Shutdown();
            ctx.Items.Add(exitItem);

            var stream = System.Windows.Application.GetResourceStream(
                new Uri("resources/favicon.ico", UriKind.Relative)
            ).Stream;

            _notifyIcon = new Forms.NotifyIcon
            {
                Icon = new System.Drawing.Icon(stream),
                Text = "IME Cursor Indicator",
                ContextMenuStrip = ctx,
                Visible = true
            };

            var container = new Grid
            {
                Background = Media.Brushes.Transparent,
                IsHitTestVisible = false,
                Focusable = false
            };
            container.Children.Add(_dotEllipse);

            _dotWindow = new Window
            {
                Width = WIN_SIZE,
                Height = WIN_SIZE,
                WindowStyle = WindowStyle.None,
                AllowsTransparency = true,
                Background = Media.Brushes.Transparent,
                ShowInTaskbar = false,
                Topmost = true,
                Content = container,
                ShowActivated = false,
                Focusable = false
            };

            _dotWindow.SourceInitialized += (s, e) =>
            {
                _dotHwnd = new WindowInteropHelper(_dotWindow).Handle;
                var ex = GetWindowLongPtr(_dotHwnd, GWL_EXSTYLE);
                ex = new IntPtr(ex.ToInt64() | WS_EX_TRANSPARENT | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW | WS_EX_LAYERED);
                SetWindowLongPtr(_dotHwnd, GWL_EXSTYLE, ex);
            };

            _dotWindow.Show();

            _imeTimer = new DispatcherTimer(DispatcherPriority.Background)
            {
                Interval = TimeSpan.FromMilliseconds(120)
            };
            _imeTimer.Tick += ImeTimer_Tick;
            _imeTimer.Start();

            _posThreadRunning = true;
            _posThread = new Thread(PositionThreadProc)
            {
                IsBackground = true,
                Priority = ThreadPriority.AboveNormal,
                Name = "DotPositionThread"
            };
            _posThread.Start();

            Closing += (s, e) => Shutdown();
        }

        private static readonly string _settingsPath = System.IO.Path.Combine(
    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
    "ImeCursorDot", "settings.json"
);

        private void AddColorItem(Forms.ContextMenuStrip ctx, string name, Media.Brush fill, Media.Brush stroke, bool isDefault)
        {
            string saved = LoadColorName();
            bool check = saved == name || (saved == "" && isDefault);

            var item = new Forms.ToolStripMenuItem(name) { Checked = check };

            if (check)
            {
                _dotEllipse.Fill = fill;
                _dotEllipse.Stroke = stroke;
            }

            item.Click += (_, __) =>
            {
                foreach (var obj in ctx.Items)
                    if (obj is Forms.ToolStripMenuItem mi) mi.Checked = false;

                item.Checked = true;
                _dotEllipse.Fill = fill;
                _dotEllipse.Stroke = stroke;
                SaveColorName(name);
            };

            ctx.Items.Add(item);
        }

        private static string LoadColorName()
        {
            try
            {
                if (System.IO.File.Exists(_settingsPath))
                {
                    string json = System.IO.File.ReadAllText(_settingsPath);
                    // {"color":"White"} の形式から値を取り出す
                    var match = System.Text.RegularExpressions.Regex.Match(json, "\"color\"\\s*:\\s*\"([^\"]+)\"");
                    if (match.Success) return match.Groups[1].Value;
                }
            }
            catch { }
            return "";
        }

        private static void SaveColorName(string name)
        {
            try
            {
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(_settingsPath)!);
                System.IO.File.WriteAllText(_settingsPath, $"{{\"color\":\"{name}\"}}");
            }
            catch { }
        }

        private void Shutdown()
        {
            _posThreadRunning = false;
            _imeTimer?.Stop();
            _dotWindow?.Close();
            _notifyIcon.Visible = false;
            _notifyIcon.Dispose();
            System.Windows.Application.Current.Shutdown();
        }

        private void PositionThreadProc()
        {
            timeBeginPeriod(1);
            try
            {
                int lastX = int.MinValue;
                int lastY = int.MinValue;

                while (_posThreadRunning)
                {
                    DwmFlush();

                    if (_dotHwnd == IntPtr.Zero) continue;
                    if (!GetCursorPos(out POINT pt)) continue;

                    int x = pt.X + 8;
                    int y = pt.Y + 8;

                    bool moved = (x != lastX || y != lastY);

                    IntPtr candidateHwnd = GetCandidateWindowForZOrder();

                    if (candidateHwnd != IntPtr.Zero)
                    {
                        SetWindowPos(_dotHwnd, HWND_TOPMOST, x, y, 0, 0, SWP_NOSIZE | SWP_NOACTIVATE | SWP_NOREDRAW | SWP_NOCOPYBITS);
                        SetWindowPos(candidateHwnd, _dotHwnd, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_NOREDRAW | SWP_NOCOPYBITS);
                    }
                    else if (moved)
                    {
                        SetWindowPos(_dotHwnd, HWND_TOPMOST, x, y, 0, 0, SWP_NOSIZE | SWP_NOACTIVATE | SWP_NOREDRAW | SWP_NOCOPYBITS);
                    }

                    if (moved)
                    {
                        lastX = x;
                        lastY = y;
                    }
                }
            }
            finally
            {
                timeEndPeriod(1);
            }
        }

        private IntPtr GetCandidateWindowForZOrder()
        {
            IntPtr hwnd = FindWindow("mscandui40.candidate", null);
            if (hwnd != IntPtr.Zero && IsWindowVisible(hwnd))
                return hwnd;

            hwnd = FindCoreWindowCandidate();
            if (hwnd != IntPtr.Zero)
                return hwnd;

            return IntPtr.Zero;
        }

        private IntPtr FindCoreWindowCandidate()
        {
            IntPtr found = IntPtr.Zero;
            EnumWindows((hwnd, _) =>
            {
                if (!IsWindowVisible(hwnd)) return true;

                var sb = new StringBuilder(256);
                GetClassName(hwnd, sb, sb.Capacity);
                if (sb.ToString() != "Windows.UI.Core.CoreWindow") return true;

                var title = new StringBuilder(256);
                GetWindowText(hwnd, title, title.Capacity);
                string titleStr = title.ToString();

                if (titleStr == "" || titleStr.Contains("Microsoft Input") || titleStr.Contains("候補"))
                {
                    found = hwnd;
                    return false;
                }
                return true;
            }, IntPtr.Zero);
            return found;
        }

        private void ImeTimer_Tick(object? sender, EventArgs e)
        {
            IntPtr fg = GetForegroundWindow();

            if (fg == IntPtr.Zero)
            {
                ApplyDotVisibility(false);
                _lastFgWnd = IntPtr.Zero;
                return;
            }

            const int WM_IME_CONTROL = 0x0283;
            const int IMC_GETOPENSTATUS = 0x0005;

            bool imeOpen = false;
            IntPtr imeWnd = ImmGetDefaultIMEWnd(fg);
            if (imeWnd != IntPtr.Zero)
            {
                IntPtr ok = SendMessageTimeout(
                    imeWnd,
                    (uint)WM_IME_CONTROL,
                    new IntPtr(IMC_GETOPENSTATUS),
                    IntPtr.Zero,
                    SendMessageTimeoutFlags.SMTO_ABORTIFHUNG,
                    50,
                    out IntPtr result
                );
                if (ok != IntPtr.Zero)
                    imeOpen = result != IntPtr.Zero;
            }

            _lastFgWnd = fg;

            if (!imeOpen && _lastImeOpen)
            {
                if (IsCandidateWindowVisible(fg))
                    return;
            }

            ApplyDotVisibility(imeOpen);
        }

        private bool IsCandidateWindowVisible(IntPtr fgWnd)
        {
            if (FindWindow("mscandui40.candidate", null) != IntPtr.Zero)
                return true;

            IntPtr imeUiWnd = ImmGetDefaultIMEWnd(fgWnd);
            if (imeUiWnd != IntPtr.Zero)
            {
                IntPtr child = GetWindow(imeUiWnd, GW_CHILD);
                if (child != IntPtr.Zero && IsWindowVisible(child))
                    return true;
            }

            return false;
        }

        private void ApplyDotVisibility(bool imeOpen)
        {
            if (imeOpen == _lastImeOpen) return;
            _lastImeOpen = imeOpen;
            _dotEllipse.Visibility = imeOpen ? Visibility.Visible : Visibility.Collapsed;
        }

        private const int GWL_EXSTYLE = -20;
        private const int GW_CHILD = 5;
        private const long WS_EX_TRANSPARENT = 0x00000020L;
        private const long WS_EX_NOACTIVATE = 0x08000000L;
        private const long WS_EX_TOOLWINDOW = 0x00000080L;
        private const long WS_EX_LAYERED = 0x00080000L;

        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOACTIVATE = 0x0010;
        private const uint SWP_NOREDRAW = 0x0008;
        private const uint SWP_NOCOPYBITS = 0x0100;
        private const uint SWP_NOMOVE = 0x0002;

        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT { public int X; public int Y; }

        [DllImport("user32.dll")] private static extern bool GetCursorPos(out POINT p);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("imm32.dll")] private static extern IntPtr ImmGetDefaultIMEWnd(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr FindWindow(string? lpClassName, string? lpWindowName);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindow(IntPtr hWnd, int uCmd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SendMessageTimeout(
            IntPtr hWnd,
            uint Msg,
            IntPtr wParam,
            IntPtr lParam,
            SendMessageTimeoutFlags flags,
            uint timeout,
            out IntPtr lpdwResult
        );

        [Flags]
        private enum SendMessageTimeoutFlags : uint
        {
            SMTO_ABORTIFHUNG = 0x0002
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
            int X, int Y, int cx, int cy, uint uFlags);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("winmm.dll")] private static extern uint timeBeginPeriod(uint uPeriod);
        [DllImport("winmm.dll")] private static extern uint timeEndPeriod(uint uPeriod);
        [DllImport("dwmapi.dll")] private static extern int DwmFlush();

        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")] private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);
        [DllImport("user32.dll", EntryPoint = "GetWindowLongW")] private static extern IntPtr GetWindowLong32(IntPtr hWnd, int nIndex);
        private static IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex)
            => IntPtr.Size == 8 ? GetWindowLongPtr64(hWnd, nIndex) : GetWindowLong32(hWnd, nIndex);

        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW")] private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr v);
        [DllImport("user32.dll", EntryPoint = "SetWindowLongW")] private static extern int SetWindowLong32(IntPtr hWnd, int nIndex, int v);
        private static IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr v)
        {
            if (IntPtr.Size == 8) return SetWindowLongPtr64(hWnd, nIndex, v);
            return new IntPtr(SetWindowLong32(hWnd, nIndex, v.ToInt32()));
        }
    }
}