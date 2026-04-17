using System.Runtime.InteropServices;
using System.Text;

namespace ClippyRWAvalonia.Services;

public sealed class DesktopInspectorService
{
    private const int WmSetText = 0x000C;
    private const int BmClick = 0x00F5;

    public string Execute(string actionArgument)
    {
        if (!OperatingSystem.IsWindows())
        {
            return "desktop inspector is only available on Windows";
        }

        var activeWindow = GetForegroundWindow();
        if (activeWindow == IntPtr.Zero)
        {
            return "no active window detected";
        }

        var controls = EnumerateVisibleChildControls(activeWindow);
        var normalized = (actionArgument ?? string.Empty).Trim();
        if (normalized.Length == 0)
        {
            normalized = "list_controls";
        }

        if (normalized.Equals("list_controls", StringComparison.OrdinalIgnoreCase))
        {
            return BuildControlsSummary(activeWindow, controls);
        }

        if (normalized.Equals("read_form", StringComparison.OrdinalIgnoreCase))
        {
            return BuildFormSummary(activeWindow, controls);
        }

        if (normalized.Equals("read_table", StringComparison.OrdinalIgnoreCase))
        {
            return BuildTableSummary(activeWindow, controls);
        }

        if (normalized.Equals("read_dialog", StringComparison.OrdinalIgnoreCase))
        {
            return BuildDialogSummary(activeWindow, controls);
        }

        if (normalized.Equals("read_selected_row", StringComparison.OrdinalIgnoreCase))
        {
            return BuildSelectedRowSummary(activeWindow, controls);
        }

        if (normalized.Equals("focus_window", StringComparison.OrdinalIgnoreCase))
        {
            SetForegroundWindow(activeWindow);
            return $"focused window {DescribeWindow(activeWindow)}";
        }

        if (normalized.StartsWith("read_control:", StringComparison.OrdinalIgnoreCase))
        {
            var query = normalized["read_control:".Length..].Trim();
            var match = FindBestControl(controls, query);
            return match == null ? $"no matching control for '{query}'" : DescribeControl("control", match);
        }

        if (normalized.StartsWith("focus_control:", StringComparison.OrdinalIgnoreCase))
        {
            var query = normalized["focus_control:".Length..].Trim();
            var match = FindBestControl(controls, query);
            if (match == null)
            {
                return $"no matching control for '{query}'";
            }

            SetForegroundWindow(activeWindow);
            SetFocus(match.Handle);
            return $"focused {DescribeControl("control", match)}";
        }

        if (normalized.StartsWith("click_control:", StringComparison.OrdinalIgnoreCase))
        {
            var query = normalized["click_control:".Length..].Trim();
            var match = FindBestControl(controls, query);
            if (match == null)
            {
                return $"no matching control for '{query}'";
            }

            SetForegroundWindow(activeWindow);
            SendMessage(match.Handle, BmClick, IntPtr.Zero, IntPtr.Zero);
            return $"clicked {DescribeControl("control", match)}";
        }

        if (normalized.StartsWith("type_control:", StringComparison.OrdinalIgnoreCase))
        {
            var payload = normalized["type_control:".Length..].Trim();
            var separatorIndex = payload.IndexOf('=');
            if (separatorIndex <= 0)
            {
                return "type_control requires name=text";
            }

            var query = payload[..separatorIndex].Trim();
            var text = payload[(separatorIndex + 1)..];
            var match = FindBestControl(controls, query);
            if (match == null)
            {
                return $"no matching control for '{query}'";
            }

            SetForegroundWindow(activeWindow);
            SetFocus(match.Handle);
            SendMessage(match.Handle, WmSetText, IntPtr.Zero, text);
            return $"typed into {DescribeControl("control", match)}";
        }

        if (normalized.StartsWith("activate_tab:", StringComparison.OrdinalIgnoreCase))
        {
            var query = normalized["activate_tab:".Length..].Trim();
            var match = FindBestControl(controls, query);
            if (match == null)
            {
                return $"no matching tab-like control for '{query}'";
            }

            SetForegroundWindow(activeWindow);
            SetFocus(match.Handle);
            return $"focused tab candidate {DescribeControl("control", match)}";
        }

        return $"unsupported inspector action '{normalized}'";
    }

    private static string BuildControlsSummary(IntPtr activeWindow, IReadOnlyList<ControlSnapshot> controls)
    {
        var lines = new List<string>
        {
            $"window: {DescribeWindow(activeWindow)}",
            $"visible child controls: {controls.Count}"
        };

        foreach (var control in controls.Take(24))
        {
            lines.Add($"- {control.ClassName} | {TrimForDisplay(control.Text)} | 0x{control.Handle.ToInt64():X}");
        }

        if (controls.Count > 24)
        {
            lines.Add($"... {controls.Count - 24} more");
        }

        return string.Join(Environment.NewLine, lines);
    }

    private static string BuildFormSummary(IntPtr activeWindow, IReadOnlyList<ControlSnapshot> controls)
    {
        var lines = new List<string> { $"form summary for {DescribeWindow(activeWindow)}" };
        foreach (var control in controls
                     .Where(control => !string.IsNullOrWhiteSpace(control.Text))
                     .Take(18))
        {
            lines.Add($"- {control.ClassName}: {TrimForDisplay(control.Text)}");
        }

        return lines.Count == 1 ? lines[0] + Environment.NewLine + "no readable form controls found" : string.Join(Environment.NewLine, lines);
    }

    private static string BuildTableSummary(IntPtr activeWindow, IReadOnlyList<ControlSnapshot> controls)
    {
        var tableControls = controls
            .Where(control =>
                control.ClassName.Contains("List", StringComparison.OrdinalIgnoreCase) ||
                control.ClassName.Contains("Grid", StringComparison.OrdinalIgnoreCase) ||
                control.ClassName.Contains("Table", StringComparison.OrdinalIgnoreCase) ||
                control.ClassName.Contains("Data", StringComparison.OrdinalIgnoreCase))
            .Take(12)
            .ToList();

        if (tableControls.Count == 0)
        {
            return $"table summary for {DescribeWindow(activeWindow)}{Environment.NewLine}no table-like controls detected";
        }

        return string.Join(Environment.NewLine,
        [
            $"table summary for {DescribeWindow(activeWindow)}",
            ..tableControls.Select(control => $"- {control.ClassName}: {TrimForDisplay(control.Text)}")
        ]);
    }

    private static string BuildDialogSummary(IntPtr activeWindow, IReadOnlyList<ControlSnapshot> controls)
    {
        var dialogControls = controls
            .Where(control =>
                !string.IsNullOrWhiteSpace(control.Text) ||
                control.ClassName.Contains("Button", StringComparison.OrdinalIgnoreCase) ||
                control.ClassName.Contains("Edit", StringComparison.OrdinalIgnoreCase))
            .Take(16)
            .ToList();

        return string.Join(Environment.NewLine,
        [
            $"dialog summary for {DescribeWindow(activeWindow)}",
            ..dialogControls.Select(control => $"- {control.ClassName}: {TrimForDisplay(control.Text)}")
        ]);
    }

    private static string BuildSelectedRowSummary(IntPtr activeWindow, IReadOnlyList<ControlSnapshot> controls)
    {
        var candidate = controls.FirstOrDefault(control =>
            control.ClassName.Contains("List", StringComparison.OrdinalIgnoreCase) ||
            control.ClassName.Contains("Grid", StringComparison.OrdinalIgnoreCase) ||
            control.ClassName.Contains("Data", StringComparison.OrdinalIgnoreCase)) ??
                        controls.FirstOrDefault(control => !string.IsNullOrWhiteSpace(control.Text));

        return candidate == null
            ? $"selected row summary for {DescribeWindow(activeWindow)}{Environment.NewLine}no list-like control found"
            : $"selected row summary for {DescribeWindow(activeWindow)}{Environment.NewLine}{DescribeControl("candidate", candidate)}";
    }

    private static ControlSnapshot? FindBestControl(IReadOnlyList<ControlSnapshot> controls, string query)
    {
        if (string.IsNullOrWhiteSpace(query))
        {
            return null;
        }

        var normalized = query.Trim();
        return controls
            .Select(control => new { Control = control, Score = ScoreControl(control, normalized) })
            .Where(entry => entry.Score > 0)
            .OrderByDescending(entry => entry.Score)
            .Select(entry => entry.Control)
            .FirstOrDefault();
    }

    private static int ScoreControl(ControlSnapshot control, string query)
    {
        var score = 0;
        if (control.Text.Equals(query, StringComparison.OrdinalIgnoreCase))
        {
            score += 10;
        }
        else if (!string.IsNullOrWhiteSpace(control.Text) && control.Text.Contains(query, StringComparison.OrdinalIgnoreCase))
        {
            score += 6;
        }

        if (control.ClassName.Equals(query, StringComparison.OrdinalIgnoreCase))
        {
            score += 8;
        }
        else if (control.ClassName.Contains(query, StringComparison.OrdinalIgnoreCase))
        {
            score += 3;
        }

        if (query.Length >= 3 && control.Text.Replace("&", string.Empty).Contains(query, StringComparison.OrdinalIgnoreCase))
        {
            score += 2;
        }

        return score;
    }

    private static List<ControlSnapshot> EnumerateVisibleChildControls(IntPtr parentHandle)
    {
        var controls = new List<ControlSnapshot>();
        EnumChildWindows(parentHandle, (handle, _) =>
        {
            if (!IsWindowVisible(handle))
            {
                return true;
            }

            var className = GetWindowClassName(handle);
            var text = GetWindowTextValue(handle);
            if (string.IsNullOrWhiteSpace(className) && string.IsNullOrWhiteSpace(text))
            {
                return true;
            }

            controls.Add(new ControlSnapshot
            {
                Handle = handle,
                ClassName = className,
                Text = text
            });
            return true;
        }, IntPtr.Zero);
        return controls;
    }

    private static string DescribeWindow(IntPtr handle)
    {
        return $"{GetWindowClassName(handle)} | {TrimForDisplay(GetWindowTextValue(handle))} | 0x{handle.ToInt64():X}";
    }

    private static string DescribeControl(string prefix, ControlSnapshot control)
    {
        return $"{prefix}: {control.ClassName} | {TrimForDisplay(control.Text)} | 0x{control.Handle.ToInt64():X}";
    }

    private static string TrimForDisplay(string text)
    {
        var normalized = (text ?? string.Empty).Replace(Environment.NewLine, " ").Trim();
        if (normalized.Length == 0)
        {
            return "<empty>";
        }

        return normalized.Length <= 120 ? normalized : normalized[..120] + "...";
    }

    private static string GetWindowTextValue(IntPtr handle)
    {
        var length = GetWindowTextLength(handle);
        var builder = new StringBuilder(Math.Max(length + 1, 64));
        GetWindowText(handle, builder, builder.Capacity);
        return builder.ToString();
    }

    private static string GetWindowClassName(IntPtr handle)
    {
        var builder = new StringBuilder(256);
        GetClassName(handle, builder, builder.Capacity);
        return builder.ToString();
    }

    private sealed class ControlSnapshot
    {
        public IntPtr Handle { get; init; }
        public string ClassName { get; init; } = string.Empty;
        public string Text { get; init; } = string.Empty;
    }

    private delegate bool EnumChildProc(IntPtr windowHandle, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern IntPtr SetFocus(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr windowHandle, EnumChildProc callback, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);

    [DllImport("user32.dll")]
    private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
}
