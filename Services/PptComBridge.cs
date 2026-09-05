using System;
using System.IO;
using System.Runtime.InteropServices;
using Avalonia.Threading;

namespace PptConsole.Services;

/// <summary>
/// PowerPoint COM 桥（晚期绑定，dynamic → IDispatch，不依赖 Office interop 程序集）。
///
/// 用于补齐两块能力：
///  1. 页码感知：轮询放映视图当前页，任意翻页方式（键盘/遥控/触控/超链）都能校准
///     墨迹层的"按页记忆"页码——取代原先的"仅内部计数"。
///  2. 页面列表：动画页数 / 当前页 / 跳页（GotoSlide）/ 缩略图导出。
///
/// 只读观测 + 一次性跳页，不接管、不修改放映内容。所有调用须在 UI(STA) 线程执行。
/// 连接失败（PowerPoint 未运行 / 未放映）时优雅降级：IsConnected=false，
/// App 回到内部计数兜底。ProgID 候选含 MS PowerPoint 与 WPS 演示
/// （WPS 对 PowerPoint COM 接口做兼容，可尝试挂接；未在真机逐项验证）。
/// </summary>
internal sealed class PptComBridge
{
    private const int PollIntervalMs = 400;

    /// <summary>COM ProgID 候选：MS PowerPoint 优先，WPS 演示（KWPP.Application）兜底。</summary>
    private static readonly string[] ComProgIds =
    {
        "PowerPoint.Application",
        "KWPP.Application",
    };

    // .NET 8 的 Marshal 类不再提供 GetActiveObject（仅 .NET Framework 有），
    // 这里直接 P/Invoke ole32 的 GetActiveObject / CLSIDFromProgID 实现等价逻辑。
    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int CLSIDFromProgID(string lpszProgID, out Guid lpclsid);

    [DllImport("ole32.dll")]
    private static extern int GetActiveObject(ref Guid rclsid, IntPtr pvReserved,
        [MarshalAs(UnmanagedType.Interface)] out object? ppunk);

    /// <summary>按 ProgID 获取正在运行的 COM 实例（ROT 查询）；未运行返回 null。</summary>
    private static object? GetRunningComObject(string progId)
    {
        if (CLSIDFromProgID(progId, out var clsid) != 0)
            return null;
        return GetActiveObject(ref clsid, IntPtr.Zero, out var obj) == 0 ? obj : null;
    }

    private readonly DispatcherTimer _poll = new() { Interval = TimeSpan.FromMilliseconds(PollIntervalMs) };
    private dynamic? _app;
    private int _slideCount = -1;
    private int _currentSlide = -1;

    /// <summary>当前页变化（1 基）。由轮询触发，UI 线程。</summary>
    public event Action<int>? CurrentSlideChanged;

    public bool IsConnected => _app is not null;
    /// <summary>页数；尚未得知为 -1。</summary>
    public int SlideCount => _slideCount;
    /// <summary>当前页（1 基）；未得知为 -1。</summary>
    public int CurrentSlide => _currentSlide;

    public PptComBridge()
    {
        _poll.Tick += (_, _) => Poll();
    }

    /// <summary>附着到正在运行的放映程序并开始轮询。返回是否成功。</summary>
    public bool Connect()
    {
        // 依次尝试 MS PowerPoint / WPS 演示的 COM ProgID
        object? app = null;
        foreach (var progId in ComProgIds)
        {
            app = GetRunningComObject(progId);
            if (app is not null)
                break;
        }

        _app = app;
        if (_app is null)
            return false;

        _slideCount = TryCount();
        _currentSlide = TryCurrentSlide();

        // 无活跃放映（程序在跑但没有打开演示/未进入放映）时不接管：
        // 页码恒为 -1，页面列表会只剩 1 页 —— 让 App 回到内部计数兜底。
        if (_slideCount <= 0 || _currentSlide <= 0)
        {
            Disconnect();
            return false;
        }

        _poll.Start();
        return true;
    }

    public void Disconnect()
    {
        _poll.Stop();
        _app = null;
        _slideCount = -1;
        _currentSlide = -1;
    }

    /// <summary>跳到指定页（1 基）。成功返回 true。</summary>
    public bool GotoSlide(int number)
    {
        if (_app is null)
            return false;

        try
        {
            dynamic view = _app.ActivePresentation.SlideShowWindow.View;
            view.GotoSlide(number);
            _currentSlide = number;
            CurrentSlideChanged?.Invoke(number);
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>把某页导出为临时 PNG，返回路径（调用方负责删除）；失败返回 null。</summary>
    public string? ExportSlideImage(int index, int width, int height)
    {
        if (_app is null)
            return null;

        try
        {
            var tmp = Path.Combine(Path.GetTempPath(), $"PptConsole_thumb_{index}_{Guid.NewGuid():N}.png");
            dynamic slide = _app.ActivePresentation.Slides[index];
            slide.Export(tmp, "PNG", width, height);
            return File.Exists(tmp) ? tmp : null;
        }
        catch
        {
            return null;
        }
    }

    private void Poll()
    {
        if (_app is null)
            return;

        int n = TryCurrentSlide();
        if (n > 0 && n != _currentSlide)
        {
            _currentSlide = n;
            var count = TryCount();
            if (count > 0)
                _slideCount = count;
            CurrentSlideChanged?.Invoke(n);
        }
    }

    /// <summary>读取当前放映页（1 基 SlideIndex）；未放映/异常返回 -1。</summary>
    private int TryCurrentSlide()
    {
        if (_app is null)
            return -1;

        try
        {
            dynamic slide = _app.ActivePresentation.SlideShowWindow.View.Slide;
            return (int)slide.SlideIndex;
        }
        catch
        {
            return -1;
        }
    }

    private int TryCount()
    {
        if (_app is null)
            return -1;

        try
        {
            dynamic slides = _app.ActivePresentation.Slides;
            return (int)slides.Count;
        }
        catch
        {
            return -1;
        }
    }
}