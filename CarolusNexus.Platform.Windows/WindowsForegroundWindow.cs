using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using CarolusNexus.Core;

namespace CarolusNexus.Platform.Windows;

public static class WindowsForegroundWindow
{
    public static ActiveWindowInfo GetActiveWindow()
    {
        if (!OperatingSystem.IsWindows())
        {
            return new ActiveWindowInfo
            {
                ProcessName = "unsupported",
                WindowTitle = "active window inspection is only available on Windows",
                AppKind = "unsupported",
                DesktopFramework = "unsupported"
            };
        }

        try
        {
            var handle = GetForegroundWindow();
            if (handle == IntPtr.Zero)
            {
                return new ActiveWindowInfo();
            }

            GetWindowThreadProcessId(handle, out var processId);
            var processName = "unknown app";
            if (processId != 0)
            {
                using var process = Process.GetProcessById((int)processId);
                processName = process.ProcessName;
            }

            var titleBuilder = new StringBuilder(512);
            var classBuilder = new StringBuilder(256);
            GetWindowText(handle, titleBuilder, titleBuilder.Capacity);
            GetClassName(handle, classBuilder, classBuilder.Capacity);
            GetWindowRect(handle, out var rect);
            var className = classBuilder.ToString().Trim();

            return new ActiveWindowInfo
            {
                ProcessName = processName,
                WindowTitle = titleBuilder.ToString().Trim(),
                WindowClassName = className,
                AppKind = AppKindDetector.FromProcessName(processName),
                DesktopFramework = DetectDesktopFramework(className, processName),
                Left = rect.Left,
                Top = rect.Top,
                Width = Math.Max(0, rect.Right - rect.Left),
                Height = Math.Max(0, rect.Bottom - rect.Top)
            };
        }
        catch
        {
            return new ActiveWindowInfo();
        }
    }

    private static string DetectDesktopFramework(string windowClassName, string processName)
    {
        var normalizedClass = (windowClassName ?? string.Empty).Trim();
        var normalizedProcess = (processName ?? string.Empty).Trim().ToLowerInvariant();

        if (normalizedClass.StartsWith("WindowsForms10", StringComparison.OrdinalIgnoreCase))
        {
            return "winforms";
        }

        if (normalizedClass.StartsWith("HwndWrapper", StringComparison.OrdinalIgnoreCase))
        {
            return "wpf";
        }

        if (normalizedClass.StartsWith("SunAwt", StringComparison.OrdinalIgnoreCase) || normalizedProcess.Contains("java", StringComparison.OrdinalIgnoreCase))
        {
            return "java";
        }

        if (normalizedClass.StartsWith("Qt", StringComparison.OrdinalIgnoreCase))
        {
            return "qt";
        }

        return "classic";
    }

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder text, int count);

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }
}
