using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Shapes;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Platform;
using PptConsole.Animations;

namespace PptConsole.Views;

/// <summary>
/// 中胶囊窗口：笔 / 选择 / 橡皮 + 向上扩展的工具面板。
/// 面板打开时窗口高度升到 160（104 面板 + 56 胶囊）并重定位；
/// 收起回 56，整窗只覆盖内容 → 不占用无效空白。
/// </summary>
public partial class CenterCapsuleWindow : Window
{
    public enum Tool { Select, Pen, Eraser }

    private static readonly Color[] PenColors =
    {
        Color.Parse("#C6CA4C"),
        Color.Parse("#E03131"),
        Color.Parse("#38BDF8"),
        Color.Parse("#E9E7E4"),
    };
    private static readonly double[] PenThicknesses = { 2.0, 3.5, 5.0 };
    private static readonly double[] EraserRadii = { 16, 28, 44 };

    // 对外事件
    public event Action<Tool>? ToolChanged;
    public event Action<Color, double>? PenSettingsChanged;
    public event Action<double>? EraserSettingsChanged;
    public event Action? InkUndo;
    public event Action? InkCleared;

    // 状态
    private Tool _tool = Tool.Select;
    private bool _panelOpen;
    private int _penColorIndex;
    private int _penSizeIndex = 1;
    private int _eraserSizeIndex = 1;

    private Screen? _screen;
    private CancellationTokenSource? _panelCts;

    /// <summary>动画用胶囊内容。</summary>
    public Border Pod => Pill;

    public CenterCapsuleWindow()
    {
        CapsuleBehavior.Init(this);
        InitializeComponent();

        // 工具切换
        BindTap(PenZone, PenCanvas, () => OnToolTapped(Tool.Pen));
        BindTap(SelectZone, SelectCanvas, () => OnToolTapped(Tool.Select));
        BindTap(EraserZone, EraserCanvas, () => OnToolTapped(Tool.Eraser));

        // 面板选择项
        BindPanelDots();
        BindPanelButtons();

        UpdateToolVisuals();
    }

    /// <summary>保留当前屏幕引用，供面板开合时自适应重定位。</summary>
    public void RememberScreen(Screen screen) => _screen = screen;

    /// <summary>会话重置：工具回选择、面板瞬间收起（新一次吊起时调用）。</summary>
    public void ResetSession()
    {
        _panelCts?.Cancel();
        _panelCts = null;
        _tool = Tool.Select;
        UpdateToolVisuals();
        CollapsePanelInstant();
    }

    /// <summary>立即触发面板收起（忽略其动画协程，供出场时快速复位）。</summary>
    public void CollapsePanelFireForget()
    {
        if (_panelOpen)
            _ = CollapsePanelAsync();
    }

    // ---------------- 输入区 + 波纹 ----------------

    private void BindTap(Border zone, Canvas rippleHost, Action action)
    {
        BindTap(zone, () =>
        {
            PlayRipple(rippleHost, zone, action);
        });
    }

    private void BindTap(Border zone, Action action)
    {
        zone.Cursor = new Cursor(StandardCursorType.Hand);
        zone.PointerPressed += (_, _) => action();
    }

    private static void PlayRipple(Canvas host, Border zone, Action action)
    {
        var position = new Point(
            zone.Bounds.Width / 2,
            zone.Bounds.Height / 2);

        const double size = 44d;
        var ellipse = new Ellipse
        {
            Width = size,
            Height = size,
            Fill = new SolidColorBrush(Color.Parse("#FF656363")),
            Opacity = 0,
            RenderTransform = new ScaleTransform(0, 0),
            RenderTransformOrigin = new RelativePoint(0.5, 0.5, RelativeUnit.Relative),
            IsHitTestVisible = false,
        };
        Canvas.SetLeft(ellipse, position.X - size / 2);
        Canvas.SetTop(ellipse, position.Y - size / 2);
        host.Children.Add(ellipse);

        _ = RunRippleAsync(host, ellipse);
        action();
    }

    private static async Task RunRippleAsync(Canvas host, Ellipse ellipse)
    {
        try { await ConsoleAnimations.TapRipple(2.2, 0.45).RunAsync(ellipse); }
        catch { }
        host.Children.Remove(ellipse);
    }

    // ---------------- 工具切换 + 面板 ----------------

    private void OnToolTapped(Tool tool)
    {
        if (tool == _tool && tool != Tool.Select)
        {
            _ = _panelOpen ? CollapsePanelAsync() : ExpandPanelAsync(tool);
            return;
        }

        _tool = tool;
        UpdateToolVisuals();
        ToolChanged?.Invoke(tool);

        if (tool == Tool.Select)
        {
            _ = CollapsePanelAsync();
        }
        else if (_panelOpen)
        {
            _ = SwitchPanelContentAsync(tool);
        }
        else
        {
            _ = ExpandPanelAsync(tool);
        }
    }

    private async Task ExpandPanelAsync(Tool tool)
    {
        FreezePanelState();
        _panelCts?.Cancel();
        _panelCts?.Dispose();
        _panelCts = new CancellationTokenSource();
        var ct = _panelCts.Token;

        ShowPanelFor(tool, resetRows: true);
        PanelHost.IsVisible = true;
        _panelOpen = true;

        // 窗口先一步扩到"面板+胶囊+回弹余量"高度：几何原子变更（Place 单次 SetWindowPos），
        // 且此瞬间扩展区全透明、面板高度仍为冻结值 → 视觉无感，只为动画腾出过冲空间
        Reposition(open: true, headroom: true);

        var rows = RowsFor(tool);
        try
        {
            // 阶段一：胶囊上缘变方（28→18）——"先转变为方形"，从当前值出发可随时打断
            double r0 = Pill.CornerRadius.TopLeft;
            if (Math.Abs(r0 - 18) > 0.5)
                await ConsoleAnimations.PillCorner(r0, 18, 140).RunAsync(Pill, ct);
            ct.ThrowIfCancellationRequested();

            // 阶段二：面板向上撑开（BackOut 回弹生长）+ 行内容等结构落位后错峰显现
            await Task.WhenAll(
                ConsoleAnimations.PanelExpand(PanelHost.Height, ConsoleMetrics.PanelHeight, PanelHost.Opacity)
                    .RunAsync(PanelHost, ct),
                Task.WhenAll(rows.Select((r, i) =>
                    ConsoleAnimations.PanelRowIn(i, 300, r.Opacity).RunAsync(r, ct))));
        }
        catch (OperationCanceledException)
        {
            return;   // 被打断：状态已冻结在当前值，交给下一次开/合衔接
        }
        finally
        {
            // 动画落定后收掉回弹余量（纯透明区，视觉无感）
            if (!ct.IsCancellationRequested)
                Reposition(open: true, headroom: false);
        }
    }

    private async Task CollapsePanelAsync()
    {
        if (!_panelOpen)
            return;

        FreezePanelState();
        _panelCts?.Cancel();
        _panelCts?.Dispose();
        _panelCts = new CancellationTokenSource();
        var ct = _panelCts.Token;

        var rows = RowsFor(_tool);
        try
        {
            // 阶段一：面板向下收拢（行内容同步错峰退场）
            await Task.WhenAll(
                ConsoleAnimations.PanelCollapse(PanelHost.Height, 0, PanelHost.Opacity).RunAsync(PanelHost, ct),
                Task.WhenAll(rows.Select((r, i) =>
                    ConsoleAnimations.PanelRowOut(i, rows.Length, r.Opacity).RunAsync(r, ct))));
            ct.ThrowIfCancellationRequested();

            // 阶段二：胶囊上缘回圆（方形→胶囊）
            double r0 = Pill.CornerRadius.TopLeft;
            if (Math.Abs(r0 - ConsoleMetrics.ToolHeight / 2) > 0.5)
                await ConsoleAnimations.PillCorner(r0, ConsoleMetrics.ToolHeight / 2, 140).RunAsync(Pill, ct);
        }
        catch (OperationCanceledException)
        {
            return;
        }

        CollapsePanelInstant();
    }

    private async Task SwitchPanelContentAsync(Tool tool)
    {
        FreezePanelState();
        _panelCts?.Cancel();
        _panelCts?.Dispose();
        _panelCts = new CancellationTokenSource();
        var ct = _panelCts.Token;

        // 高度/圆角可能停在半途（快速连点打断开合）：切换内容时同步"补完"，
        // 避免面板卡在中间高度导致内容被裁剪
        Reposition(open: true, headroom: true);

        var other = tool == Tool.Pen ? Tool.Eraser : Tool.Pen;
        var oldRows = RowsFor(other);
        try
        {
            await Task.WhenAll(oldRows.Select((r, i) =>
                ConsoleAnimations.PanelRowOut(i, oldRows.Length, r.Opacity).RunAsync(r, ct)));
        }
        catch (OperationCanceledException)
        {
            return;
        }

        if (ct.IsCancellationRequested)
            return;

        ShowPanelFor(tool, resetRows: true);
        var rows = RowsFor(tool);
        try
        {
            double h0 = PanelHost.Height, op0 = PanelHost.Opacity;
            double r0 = Pill.CornerRadius.TopLeft;

            var grow = h0 < ConsoleMetrics.PanelHeight - 0.5
                ? ConsoleAnimations.PanelExpand(h0, ConsoleMetrics.PanelHeight, op0).RunAsync(PanelHost, ct)
                : Task.CompletedTask;
            var corner = r0 > 18.5
                ? ConsoleAnimations.PillCorner(r0, 18, 140).RunAsync(Pill, ct)
                : Task.CompletedTask;

            await Task.WhenAll(
                grow,
                corner,
                Task.WhenAll(rows.Select((r, i) =>
                    ConsoleAnimations.PanelRowIn(i, 120, r.Opacity).RunAsync(r, ct))));
        }
        catch (OperationCanceledException)
        {
            return;
        }
        finally
        {
            if (!ct.IsCancellationRequested)
                Reposition(open: true, headroom: false);
        }
    }

    /// <summary>
    /// 把当前（可能正处于动画中间态的）视觉值固化为本地值。
    /// Avalonia 取消动画后属性会回落到本地值——若不冻结，回落到的是过期本地值
    /// （如高度回落到 0），快速连点时面板会瞬移/爆开。冻结后回落值=当前画面，视觉连续。
    /// </summary>
    private void FreezePanelState()
    {
        PanelHost.Height = PanelHost.Height;
        PanelHost.Opacity = PanelHost.Opacity;
        var cr = Pill.CornerRadius;
        Pill.CornerRadius = new CornerRadius(cr.TopLeft, cr.TopRight, cr.BottomLeft, cr.BottomRight);
        foreach (var row in new[] { PenRowColor, PenRowSize, EraserRowSize })
            row.Opacity = row.Opacity;
    }

    private void CollapsePanelInstant()
    {
        PanelHost.Height = 0;
        PanelHost.Opacity = 0;
        PanelHost.IsVisible = false;
        Pill.CornerRadius = new CornerRadius(ConsoleMetrics.ToolHeight / 2);
        _panelOpen = false;
        Reposition(open: false);
    }

    private StackPanel[] RowsFor(Tool tool) => tool == Tool.Pen
        ? new[] { PenRowColor, PenRowSize }
        : new[] { EraserRowSize };

    private void ShowPanelFor(Tool tool, bool resetRows)
    {
        PenPanel.IsVisible = tool == Tool.Pen;
        EraserPanel.IsVisible = tool == Tool.Eraser;

        if (resetRows)
        {
            foreach (var row in RowsFor(tool))
                row.Opacity = 0;
        }
    }

    private void UpdateToolVisuals()
    {
        SetToolVisual(_tool == Tool.Pen, PenPlate, PenIcon);
        SetToolVisual(_tool == Tool.Select, SelectPlate, SelectIcon);
        SetToolVisual(_tool == Tool.Eraser, EraserPlate, EraserIcon);
    }

    private static void SetToolVisual(bool active, Border plate, PathIcon icon)
    {
        plate.IsVisible = active;   // 激活：白圆底 + 深色图标
        icon.IsVisible = !active;   // 未激活：白色描线图标
    }

    /// <summary>
    /// 按开关状态把窗口定位到屏幕水平居中、底边贴 28px 边距处。
    /// headroom：开合动画期间额外预留回弹余量（BackOut 过冲不被窗口顶边平切），
    /// 动画落定后以 headroom=false 收掉。
    /// </summary>
    private void Reposition(bool open, bool headroom = false)
    {
        if (_screen is null)
            return;

        double scaling = _screen.Scaling > 0 ? _screen.Scaling : 1d;
        double screenWidthDip = _screen.Bounds.Width / scaling;
        double leftDip = (screenWidthDip - ConsoleMetrics.ToolWidth) / 2;
        double heightDip = open
            ? ConsoleMetrics.ToolStackHeight + (headroom ? ConsoleMetrics.PanelOvershootHeadroom : 0)
            : ConsoleMetrics.ToolHeight;

        CapsuleBehavior.Place(this, _screen, leftDip, ConsoleMetrics.ToolWidth, heightDip, ConsoleMetrics.BottomMarginDip);
    }

    // ---------------- 面板控件 ----------------

    private void BindPanelDots()
    {
        var colors = new[] { PenColor0, PenColor1, PenColor2, PenColor3 };
        for (int i = 0; i < colors.Length; i++)
        {
            int idx = i;
            BindTap(colors[idx], () =>
            {
                _penColorIndex = idx;
                UpdateDotSelections();
                PenSettingsChanged?.Invoke(PenColors[_penColorIndex], PenThicknesses[_penSizeIndex]);
            });
        }

        var penSizes = new[] { PenSize0, PenSize1, PenSize2 };
        for (int i = 0; i < penSizes.Length; i++)
        {
            int idx = i;
            BindTap(penSizes[idx], () =>
            {
                _penSizeIndex = idx;
                UpdateDotSelections();
                PenSettingsChanged?.Invoke(PenColors[_penColorIndex], PenThicknesses[_penSizeIndex]);
            });
        }

        var eraserSizes = new[] { EraserSize0, EraserSize1, EraserSize2 };
        for (int i = 0; i < eraserSizes.Length; i++)
        {
            int idx = i;
            BindTap(eraserSizes[idx], () =>
            {
                _eraserSizeIndex = idx;
                UpdateDotSelections();
                EraserSettingsChanged?.Invoke(EraserRadii[_eraserSizeIndex]);
            });
        }
    }

    private void BindPanelButtons()
    {
        BindTap(PenUndoBtn, () => InkUndo?.Invoke());
        BindTap(PenClearBtn, () => InkCleared?.Invoke());
        BindTap(EraserUndoBtn, () => InkUndo?.Invoke());
        BindTap(EraserClearBtn, () => InkCleared?.Invoke());
    }

    private static readonly IBrush Selected = new SolidColorBrush(Colors.White);
    private static readonly IBrush Unselected = new SolidColorBrush(Color.FromArgb(0, 0, 0, 0));

    private void UpdateDotSelections()
    {
        var colors = new[] { PenColor0, PenColor1, PenColor2, PenColor3 };
        for (int i = 0; i < colors.Length; i++)
            colors[i].BorderBrush = i == _penColorIndex ? Selected : Unselected;

        var penSizes = new[] { PenSize0, PenSize1, PenSize2 };
        for (int i = 0; i < penSizes.Length; i++)
            penSizes[i].BorderBrush = i == _penSizeIndex ? Selected : Unselected;

        var eraserSizes = new[] { EraserSize0, EraserSize1, EraserSize2 };
        for (int i = 0; i < eraserSizes.Length; i++)
            eraserSizes[i].BorderBrush = i == _eraserSizeIndex ? Selected : Unselected;
    }
}