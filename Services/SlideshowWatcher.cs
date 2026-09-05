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
        "powerpnt.exe",
        "wpp.exe",
    };

    /// <summary>放映窗口标题关键词（WPS 中文 / MS PowerPoint 中英文）。</summary>
    private static readonly string[] SlideTitleKeywords =
    {
        "幻灯片放映",
        "演示文稿放映",
        "slide show",
    };

    private const int PollMilliseconds = 1000;
    /// <summary>连续 N 次找不到才判定放映结束（防放映切换期的窗口抖动）。</summary>
    private const int EndConfirmCount = 2;

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

            if (IsSlideshowWindow(cls, title, process))
            {
                found = hwnd;
                return false;   // 停止枚举
            }
            return true;
        }, IntPtr.Zero);

        return found;
    }

    /// <summary>
    /// 三重判定：进程名 ∈ {powerpnt,wpp}.exe 是硬条件（杜绝其他软件的误报），
    /// 类名或标题任一命中即认定为放映窗口。
    /// 进程名获取失败（权限受限，如放映程序以管理员运行）时退化为纯类名匹配。
    /// </summary>
    private static bool IsSlideshowWindow(string cls, string title, string process)
    {
        bool classMatch = Array.IndexOf(SlideClasses, cls) >= 0;
        bool titleMatch = title.Length > 0 &&
            Array.Exists(SlideTitleKeywords, k => title.Contains(k, StringComparison.OrdinalIgnoreCase));

        if (process.Length == 0)
            return classMatch;   // 拿不到进程名 → 退回类名匹配

        if (Array.IndexOf(SlideProcessNames, process) < 0)
            return false;

        return classMatch || titleMatch;
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
