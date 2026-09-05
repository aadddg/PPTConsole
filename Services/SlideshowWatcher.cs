using System;
using System.Runtime.InteropServices;
using Avalonia;
using Avalonia.Threading;

namespace PptConsole.Services;

/// <summary>
/// 轮询放映窗口（EnumWindows 按类名 + 可见性/非最小化匹配），
/// 报告放映开始/结束与其所在显示器的物理边界。
/// WPS 的放映窗口类名未逐项验证，后续在 SlideClasses 里追加即可。
/// </summary>
internal sealed class SlideshowWatcher
{
    /// <summary>
    /// 放映窗口类名候选（MS PowerPoint = screenClass；WPS 演示放映 = wppslideshowwnd）。
    /// 不再使用 WPS 通用主窗 KWMainFrame —— 它不放映也存在，命中即误报。
    /// 匹配时走 EnumWindows + 可见性/非最小化校验，排除隐藏或最小化的同名窗口。
    /// </summary>
    private static readonly string[] SlideClasses =
    {
        "screenClass",          // MS PowerPoint 放映窗
        "wppslideshowwnd",      // WPS 演示放映窗（候选，未逐项验证）
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
            foreach (var candidate in SlideClasses)
            {
                if (cls == candidate)
                {
                    found = hwnd;
                    return false;   // 停止枚举
                }
            }
            return true;
        }, IntPtr.Zero);

        return found;
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
