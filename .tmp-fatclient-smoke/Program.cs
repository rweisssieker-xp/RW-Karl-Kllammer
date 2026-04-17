using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using ClippyRWAvalonia.Services;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        var inspector = new DesktopInspectorService();
        var targets = new List<(string name, Process process)>();
        try
        {
            var notepad = Process.Start(new ProcessStartInfo("notepad.exe") { UseShellExecute = true });
            if (notepad != null) targets.Add(("notepad", notepad));
        }
        catch (Exception ex)
        {
            Console.WriteLine("notepad.start=fail " + ex.Message);
        }
        try
        {
            var explorer = Process.Start(new ProcessStartInfo("explorer.exe") { UseShellExecute = true });
            if (explorer != null) targets.Add(("explorer", explorer));
        }
        catch (Exception ex)
        {
            Console.WriteLine("explorer.start=fail " + ex.Message);
        }

        Thread.Sleep(1800);
        foreach (var (name, process) in targets)
        {
            var handle = FindTopLevelWindow(process.Id);
            if (handle == IntPtr.Zero)
            {
                Console.WriteLine($"{name}=fail no window");
                continue;
            }

            SetForegroundWindow(handle);
            Thread.Sleep(700);
            Console.WriteLine($"{name}.list=" + inspector.Execute("list_controls").Split(Environment.NewLine)[0]);
            Console.WriteLine($"{name}.form=" + inspector.Execute("read_form").Split(Environment.NewLine)[0]);
            Console.WriteLine($"{name}.focus=" + inspector.Execute("focus_window"));
        }
    }

    private static IntPtr FindTopLevelWindow(int processId)
    {
        IntPtr found = IntPtr.Zero;
        EnumWindows((hWnd, _) =>
        {
            GetWindowThreadProcessId(hWnd, out var pid);
            if (pid == processId && IsWindowVisible(hWnd))
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
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int processId);
}
