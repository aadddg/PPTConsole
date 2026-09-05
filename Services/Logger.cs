using System;
using System.IO;
using System.Text;

namespace PptConsole.Services;

/// <summary>
/// 最简文件日志（追加写 %TEMP%\PptConsole\pptconsole.log）。
/// 用于闪退/异常的死后取证：全局异常钩子把堆栈落盘，方便真机排查。
/// </summary>
internal static class Logger
{
    private static readonly object Gate = new();
    private static readonly string LogPath = Path.Combine(
        Path.GetTempPath(), "PptConsole", "pptconsole.log");

    private static bool _startupLogged;

    public static void Info(string message) => Write("INFO ", message);

    public static void Error(string context, Exception ex) =>
        Write("ERROR", $"{context}: {ex.GetType().Name}: {ex.Message}\r\n{ex.StackTrace}");

    /// <summary>进程启动记录一次（版本定位：新日志段落的分隔行）。</summary>
    public static void LogStartup()
    {
        if (_startupLogged) return;
        _startupLogged = true;
        Write("START", $"PptConsole 启动，PID={Environment.ProcessId}");
    }

    private static void Write(string level, string message)
    {
        try
        {
            lock (Gate)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(LogPath)!);
                File.AppendAllText(LogPath,
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {level} {message}\r\n",
                    Encoding.UTF8);
            }
        }
        catch
        {
            // 日志失败绝不能影响主流程
        }
    }
}
