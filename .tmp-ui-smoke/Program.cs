using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using ClippyRWAvalonia.Services;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        var processIds = System.Diagnostics.Process.GetProcessesByName("ClippyRW.Avalonia").Select(p => p.Id).ToHashSet();
        if (processIds.Count == 0)
        {
            Console.WriteLine("ui=fail no process");
            return;
        }

        var handle = FindTopLevelWindow(processIds);
        if (handle == IntPtr.Zero)
        {
            Console.WriteLine("ui=fail no top level window");
            return;
        }

        SetForegroundWindow(handle);
        Thread.Sleep(1200);
        var inspector = new DesktopInspectorService();
        string[] tabs = ["Dashboard", "Setup", "Knowledge", "Rituals", "History", "Diagnostics", "Live Context", "Console", "Ask"];
        var outDir = Path.Combine("C:\\tmp\\clippy_rw\\avalonia", "ui-smoke");
        Directory.CreateDirectory(outDir);
        foreach (var tab in tabs)
        {
            var click = inspector.Execute($"click_control:{tab}");
            Thread.Sleep(900);
            if (!GetWindowRect(handle, out var rect))
            {
                Console.WriteLine($"tab={tab} click={click} capture=fail");
                continue;
            }

            using var bmp = new Bitmap(Math.Max(1, rect.Right - rect.Left), Math.Max(1, rect.Bottom - rect.Top));
            using (var g = Graphics.FromImage(bmp))
            {
                g.CopyFromScreen(rect.Left, rect.Top, 0, 0, bmp.Size);
            }

            var path = Path.Combine(outDir, tab.Replace(" ", "-") + ".png");
            bmp.Save(path, ImageFormat.Png);
            Console.WriteLine($"tab={tab} click={click} screenshot={path}");
        }

        Console.WriteLine("ui=done");
    }

    private static IntPtr FindTopLevelWindow(HashSet<int> processIds)
    {
        IntPtr found = IntPtr.Zero;
        EnumWindows((hWnd, _) =>
        {
            GetWindowThreadProcessId(hWnd, out var pid);
            if (processIds.Contains(pid) && IsWindowVisible(hWnd))
            {
                found = hWnd;
                return false;
            }
            return true;
        }, IntPtr.Zero);
        return found;
    }

    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int processId);

    private struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }
}
