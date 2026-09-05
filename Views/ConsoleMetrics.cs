namespace PptConsole.Views;

/// <summary>控制台几何度量（DIP）。中窗口高度可变：仅胶囊 or 面板+胶囊。</summary>
internal static class ConsoleMetrics
{
    public const double BottomMarginDip = 28;    // 距放映屏底边的留白
    public const double SideMarginDip = 48;      // 左右胶囊相对屏左右边距
    public const double GapDip = 24;             // 左右胶囊与中胶囊在内容上无重叠所需的最小间隔（可调）

    public const double SidePillWidth = 150;
    public const double SidePillHeight = 56;

    public const double ToolWidth = 340;
    public const double ToolHeight = 56;
    public const double PanelHeight = 104;
    public const double ToolStackHeight = ToolHeight + PanelHeight; // 160

    /// <summary>
    /// 面板撑高用 BackOut(0.275) 回弹，峰值 ≈ PanelHeight×1.09 ≈ 9.4px；
    /// 若窗口只开到 160，过冲会被窗口顶边平切（出现一条直线切边）。
    /// 开合动画期间窗口额外预留此余量，动画落定后再收掉（收掉的是纯透明区，视觉无感）。
    /// </summary>
    public const double PanelOvershootHeadroom = 14;
}