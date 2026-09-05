using System;
using System.Runtime.InteropServices;
using Avalonia;
using Avalonia.Threading;

namespace PptConsole.Services;

/// <summary>
/// 轮询放映窗口（EnumWindows：可见性 + 进程名 + 类名 + 标题关键词组合判定），
/// 报告放映开始/结束与其所在显示器的物理边界。
/// MS PowerPoint（powerpnt.exe / screenClass）与 WPS 演示（wpp.exe / 标题含"幻灯片放映"）均覆盖。
/// </summary>
internal sealed class SlideshowWatcher
{
    /// <summary>
    /// 放映窗口类名候选。MS PowerPoint 与 WPS 放映窗多为 screenClass；
    /// wppslideshowwnd 为 WPS 候选。实际类名随版本而异，因此另配
    /// 进程名 + 标题关键词双兜底（见 IsSlideshowWindow）。
    /// </summary>
    private static readonly string[] SlideClasses =
    {
        "screenClass",          // MS PowerPoint / WPS 放映窗（常见）
        "wppslideshowwnd",      // WPS 演示放映窗（候选）
    };

    /// <summary>放映程序进程名（小写，含扩展名）：MS PowerPoint / WPS 演示。</summary>
    private static readonly string[] SlideProcessNames =
    {
        "powerpnt.exe",     // MS PowerPoint
        "wpp.exe",          // WPS 演示（放映进程，任务管理器显示为 "WPS Presentation"）
        "wps.exe",          // WPS 主程序（部分版本放映也挂在该进程）
    };

    /// <summary>放映窗口标题关键词（WPS 中文 / MS PowerPoint 中英文）。</summary>
    private static readonly string[] SlideTitleKeywords =
    {
        "幻灯片放映",
        "演示文稿放映",
        "slide show",
    };

    private const int PollMilliseconds = 1000;
    /// <summary>连续 N 次找不到才判定放映结束（防放映切换期/WPS 窗口短暂抖动误判收回）。</summary>
    private const int EndConfirmCount = 3;

    private DispatcherTimer? _timer;
    private bool _running;
    private int _missCount;

    /// <summary>放映开始（参数 = 放映所在显示器物理边界）。</summary>
    public event Action<PixelRect>? SlideshowStarted;

    /// <summary>放映结束。</summary>
    public event Action? SlideshowEnded;

    public void Start()
    {
        if (_timer is not null) return;

        _timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(PollMilliseconds) };
        _timer.Tick += (_, _) => Poll();
        _timer.Start();
    }

    private void Poll()
    {
        IntPtr hwnd = Find();
        bool found = hwnd != IntPtr.Zero;

        if (found)
        {
            _missCount = 0;
            if (_running)
                return; // 仍在放映

            _running = true;
            SlideshowStarted?.Invoke(GetMonitorBounds(hwnd));
        }
        else if (_running)
        {
            if (++_missCount >= EndConfirmCount)
            {
                _running = false;
                _missCount = 0;
                SlideshowEnded?.Invoke();
            }
        }
    }

    private static IntPtr Find()
    {
        IntPtr found = IntPtr.Zero;
        Win32Interop.EnumWindows((hwnd, _) =>
        {
            // 只认可见、非最小化的顶层窗口（隐藏/最小化的同名窗口不误报放映开始）
            if (!Win32Interop.IsWindowVisible(hwnd) || Win32Interop.IsIconic(hwnd))
                return true;

            string cls = GetClassName(hwnd);
            string title = GetWindowTitle(hwnd);
            string process = GetProcessName(hwnd);

            if (IsSlideshowWindow(hwnd, cls, title, process))
            {
                found = hwnd;
                return false;   // 停止枚举
            }
            return true;
        }, IntPtr.Zero);

        return found;
    }

    /// <summary>
    /// 判定逻辑（三重信号，防误报）：
    /// - 标题含"幻灯片放映/Slide Show"：极特异（用户实测 WPS 放映窗标题
    ///   "WPS演示幻灯片放映-[xxx.pptx]"），即使进程名不在候选也接受；
    /// - 进程名 ∈ {powerpnt,wpp,wps}.exe：作为已知放映程序的确认；
    /// - 类名命中但标题未命中时：要求放映窗口特征（WS_POPUP / 全屏），
    ///   防止不放映也存在的同名类窗口误报；
    /// 进程名获取失败（权限受限）时退化为"类名或标题"匹配。
    /// </summary>
    private static bool IsSlideshowWindow(IntPtr hwnd, string cls, string title, string process)
    {
        bool classMatch = Array.IndexOf(SlideClasses, cls) >= 0;
        bool titleMatch = title.Length > 0 &&
            Array.Exists(SlideTitleKeywords, k => title.Contains(k, StringComparison.OrdinalIgnoreCase));

        if (process.Length == 0)
            return classMatch || titleMatch;   // 拿不到进程名：标题"幻灯片放映"足够特异

        if (Array.IndexOf(SlideProcessNames, process) >= 0)
        {
            if (titleMatch)
                return true;
            // 仅类名命中：必须是放映窗口特征（WS_POPUP 无控制按钮 / 全屏）
            return classMatch && IsPopupOrFullscreen(hwnd);
        }

        // 进程不是已知放映程序：仅标题命中才接受（防其他软件误报）
        return titleMatch;
    }

    /// <summary>放映窗特征：WS_POPUP 样式，或窗口尺寸覆盖其所在显示器（全屏）。</summary>
    private static bool IsPopupOrFullscreen(IntPtr hwnd)
    {
        int style = (int)Win32Interop.GetWindowLongPtr(hwnd, Win32Interop.GWL_STYLE);
        if ((style & Win32Interop.WS_POPUP) != 0)
            return true;

        if (!Win32Interop.GetWindowRect(hwnd, out var rect))
            return false;

        var monitor = Win32Interop.MonitorFromWindow(hwnd, Win32Interop.MONITOR_DEFAULTTONEAREST);
        if (monitor == IntPtr.Zero)
            return false;

        var info = new Win32Interop.MONITORINFOEX
        {
            cbSize = Marshal.SizeOf<Win32Interop.MONITORINFOEX>(),
        };

        if (!Win32Interop.GetMonitorInfo(monitor, ref info))
            return false;

        int w = rect.Right - rect.Left;
        int h = rect.Bottom - rect.Top;
        int mw = info.rcMonitor.Right - info.rcMonitor.Left;
        int mh = info.rcMonitor.Bottom - info.rcMonitor.Top;

        // 容差 2px（部分窗口在边界有 1px 偏差）
        return w >= mw - 2 && h >= mh - 2;
    }

    private static string GetWindowTitle(IntPtr hwnd)
    {
        var sb = new System.Text.StringBuilder(256);
        Win32Interop.GetWindowText(hwnd, sb, sb.Capacity);
        return sb.ToString();
    }

    private static string GetProcessName(IntPtr hwnd)
    {
        Win32Interop.GetWindowThreadProcessId(hwnd, out uint pid);
        if (pid == 0)
            return string.Empty;

        var handle = Win32Interop.OpenProcess(Win32Interop.PROCESS_QUERY_LIMITED_INFORMATION, false, pid);
        if (handle == IntPtr.Zero)
            return string.Empty;

        try
        {
            var sb = new System.Text.StringBuilder(256);
            Win32Interop.GetModuleBaseName(handle, IntPtr.Zero, sb, (uint)sb.Capacity);
            return sb.ToString().ToLowerInvariant();
        }
        finally
        {
            Win32Interop.CloseHandle(handle);
        }
    }

    private static string GetClassName(IntPtr hwnd)
    {
        var sb = new System.Text.StringBuilder(256);
        Win32Interop.GetClassName(hwnd, sb, sb.Capacity);
        return sb.ToString();
    }

    private static PixelRect GetMonitorBounds(IntPtr hwnd)
    {
        var monitor = Win32Interop.MonitorFromWindow(hwnd, Win32Interop.MONITOR_DEFAULTTONEAREST);

        var info = new Win32Interop.MONITORINFOEX
        {
            cbSize = Marshal.SizeOf<Win32Interop.MONITORINFOEX>(),
        };

        if (monitor != IntPtr.Zero && Win32Interop.GetMonitorInfo(monitor, ref info))
        {
            return new PixelRect(
                info.rcMonitor.Left, info.rcMonitor.Top,
                info.rcMonitor.Right - info.rcMonitor.Left,
                info.rcMonitor.Bottom - info.rcMonitor.Top);
        }

        return default;
    }
}
