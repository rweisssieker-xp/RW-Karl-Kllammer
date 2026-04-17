using System.Runtime.InteropServices;
using System.Text;
using ClippyRWAvalonia.Models;
using System.Windows.Forms;

namespace ClippyRWAvalonia.Services;

public sealed class AxClientAutomationService
{
    private const int WmSetText = 0x000C;
    private const int BmClick = 0x00F5;
    private const int BmSetCheck = 0x00F1;
    private const int BstChecked = 0x0001;
    private const int BstUnchecked = 0x0000;

    public AxContextSnapshot CaptureActiveContext()
    {
        if (!OperatingSystem.IsWindows())
        {
            return new AxContextSnapshot();
        }

        var handle = GetForegroundWindow();
        if (handle == IntPtr.Zero)
        {
            return new AxContextSnapshot();
        }

        return CaptureContext(handle);
    }

    public string Execute(string actionArgument)
    {
        var context = CaptureActiveContext();
        return Execute(actionArgument, context);
    }

    public string Execute(string actionArgument, AxContextSnapshot context)
    {
        if (!OperatingSystem.IsWindows())
        {
            return "ax automation is only available on Windows";
        }

        if (context == null || !context.IsAvailable || !context.IsAxClient)
        {
            return "active window is not an AX client";
        }

        var action = (actionArgument ?? string.Empty).Trim();
        if (action.Length == 0)
        {
            return context.Summary;
        }

        if (action.Equals("ax.focus_window", StringComparison.OrdinalIgnoreCase))
        {
            var handle = GetForegroundWindow();
            SetForegroundWindow(handle);
            return $"focused AX window {context.WindowTitle}";
        }

        if (action.Equals("ax.read_grid", StringComparison.OrdinalIgnoreCase))
        {
            return BuildGridSummary(context);
        }

        if (action.Equals("ax.read_context", StringComparison.OrdinalIgnoreCase))
        {
            return context.Summary + Environment.NewLine + Environment.NewLine + BuildFormSummary(context);
        }

        if (action.Equals("ax.read_form", StringComparison.OrdinalIgnoreCase))
        {
            return BuildFormSummary(context);
        }

        if (action.Equals("ax.read_selected_row", StringComparison.OrdinalIgnoreCase))
        {
            var grid = context.Grids.FirstOrDefault();
            if (grid == null)
            {
                return "no AX grid detected";
            }

            return string.IsNullOrWhiteSpace(grid.SelectedRow)
                ? $"grid '{grid.Title}' has no selected row"
                : $"selected row in '{grid.Title}': {grid.SelectedRow}";
        }

        if (action.Equals("ax.confirm_dialog", StringComparison.OrdinalIgnoreCase))
        {
            return ClickNamedAction(context, ["ok", "&ok", "yes", "&yes", "close", "&close"]);
        }

        if (action.Equals("ax.cancel_dialog", StringComparison.OrdinalIgnoreCase))
        {
            return ClickNamedAction(context, ["cancel", "&cancel", "no", "&no", "close", "&close"]);
        }

        if (action.StartsWith("ax.read_field:", StringComparison.OrdinalIgnoreCase))
        {
            var label = action["ax.read_field:".Length..].Trim();
            var field = FindBestField(context, label);
            return field == null
                ? $"field not found: {label}"
                : $"field {field.Label}: {field.Value}";
        }

        if (action.StartsWith("ax.set_field:", StringComparison.OrdinalIgnoreCase))
        {
            var payload = action["ax.set_field:".Length..].Trim();
            var separatorIndex = payload.IndexOf('=');
            if (separatorIndex <= 0)
            {
                return "ax.set_field requires label=value";
            }

            var label = payload[..separatorIndex].Trim();
            var value = payload[(separatorIndex + 1)..];
            var field = FindBestField(context, label);
            if (field == null)
            {
                return $"field not found: {label}";
            }

            if (!field.IsEditable)
            {
                return $"field is not editable: {field.Label}";
            }

            SetForegroundWindow(GetForegroundWindow());
            SetFocus((IntPtr)field.Handle);
            SendMessage((IntPtr)field.Handle, WmSetText, IntPtr.Zero, value);
            return $"set field {field.Label} = {value}";
        }

        if (action.StartsWith("ax.click_action:", StringComparison.OrdinalIgnoreCase))
        {
            var label = action["ax.click_action:".Length..].Trim();
            return ClickNamedAction(context, [label]);
        }

        if (action.StartsWith("ax.open_dialog:", StringComparison.OrdinalIgnoreCase))
        {
            var label = action["ax.open_dialog:".Length..].Trim();
            return ClickNamedAction(context, [label]);
        }

        if (action.StartsWith("ax.open_lookup:", StringComparison.OrdinalIgnoreCase))
        {
            var label = action["ax.open_lookup:".Length..].Trim();
            var field = FindBestField(context, label);
            if (field == null)
            {
                return $"lookup field not found: {label}";
            }

            SetForegroundWindow(GetForegroundWindow());
            SetFocus((IntPtr)field.Handle);
            SendKeys.SendWait("%{DOWN}");
            return $"opened lookup for {field.Label}; confirm lookup dialog before continuing";
        }

        if (action.StartsWith("ax.open_tab:", StringComparison.OrdinalIgnoreCase))
        {
            var label = action["ax.open_tab:".Length..].Trim();
            var tab = FindBestTab(context, label);
            if (string.IsNullOrWhiteSpace(tab))
            {
                return $"tab not found: {label}";
            }

            var control = FindControlByText(GetForegroundWindow(), tab, "tab");
            if (control == null)
            {
                return $"tab control not found: {tab}";
            }

            SetForegroundWindow(GetForegroundWindow());
            SetFocus(control.Handle);
            SendMessage(control.Handle, BmClick, IntPtr.Zero, IntPtr.Zero);
            return $"opened tab {tab}";
        }

        if (action.StartsWith("ax.select_grid_row:", StringComparison.OrdinalIgnoreCase))
        {
            var query = action["ax.select_grid_row:".Length..].Trim();
            var grid = context.Grids.FirstOrDefault();
            if (grid == null)
            {
                return "no AX grid detected";
            }

            var row = grid.VisibleRows
                .Select(entry => new
                {
                    Row = entry,
                    Score = ScoreText(entry, query)
                })
                .OrderByDescending(entry => entry.Score)
                .FirstOrDefault();
            if (row == null || row.Score <= 0)
            {
                return $"grid row not found: {query}";
            }

            SetForegroundWindow(GetForegroundWindow());
            SetFocus((IntPtr)grid.Handle);
            return $"selected grid row candidate in '{grid.Title}': {row.Row}";
        }

        if (action.StartsWith("ax.toggle_checkbox:", StringComparison.OrdinalIgnoreCase))
        {
            var payload = action["ax.toggle_checkbox:".Length..].Trim();
            var separatorIndex = payload.IndexOf('=');
            var label = separatorIndex > 0 ? payload[..separatorIndex].Trim() : payload;
            var rawValue = separatorIndex > 0 ? payload[(separatorIndex + 1)..].Trim() : "true";
            var desired = !rawValue.Equals("false", StringComparison.OrdinalIgnoreCase);
            var control = FindControlByText(GetForegroundWindow(), label, "button");
            if (control == null)
            {
                return $"checkbox not found: {label}";
            }

            SetForegroundWindow(GetForegroundWindow());
            SetFocus(control.Handle);
            SendMessage(control.Handle, BmSetCheck, (IntPtr)(desired ? BstChecked : BstUnchecked), IntPtr.Zero);
            return $"set checkbox {label} = {desired}";
        }

        if (action.StartsWith("ax.wait_for_form:", StringComparison.OrdinalIgnoreCase))
        {
            var query = action["ax.wait_for_form:".Length..].Trim();
            return WaitForCondition(snapshot =>
                snapshot.IsAxClient &&
                snapshot.FormName.Contains(query, StringComparison.OrdinalIgnoreCase),
                $"form {query}");
        }

        if (action.StartsWith("ax.wait_for_dialog:", StringComparison.OrdinalIgnoreCase))
        {
            var query = action["ax.wait_for_dialog:".Length..].Trim();
            return WaitForCondition(snapshot =>
                snapshot.IsAxClient &&
                snapshot.DialogTitle.Contains(query, StringComparison.OrdinalIgnoreCase),
                $"dialog {query}");
        }

        if (action.StartsWith("ax.wait_for_text:", StringComparison.OrdinalIgnoreCase))
        {
            var query = action["ax.wait_for_text:".Length..].Trim();
            return WaitForCondition(snapshot =>
                snapshot.IsAxClient &&
                (snapshot.WindowTitle.Contains(query, StringComparison.OrdinalIgnoreCase) ||
                 snapshot.ValidationSummary.Contains(query, StringComparison.OrdinalIgnoreCase) ||
                 snapshot.Fields.Any(field => field.Value.Contains(query, StringComparison.OrdinalIgnoreCase))),
                $"text {query}");
        }

        return $"unsupported AX action '{action}'";
    }

    private static AxContextSnapshot CaptureContext(IntPtr handle)
    {
        var processName = GetProcessName(handle);
        var title = GetWindowTextValue(handle);
        var className = GetWindowClassName(handle);
        var controls = EnumerateVisibleChildControls(handle);
        var context = new AxContextSnapshot
        {
            IsAvailable = true,
            ProcessName = processName,
            WindowTitle = title,
            WindowClassName = className,
            IsAxClient = IsAxProcess(processName, title, className)
        };

        if (!context.IsAxClient)
        {
            return context;
        }

        context.WindowType = InferWindowType(title, className, controls);
        context.FormName = InferFormName(title, controls);
        context.DialogTitle = context.WindowType.Contains("dialog", StringComparison.OrdinalIgnoreCase) ? title : string.Empty;
        context.Tabs = ExtractTabs(controls);
        context.ActiveTab = context.Tabs.FirstOrDefault() ?? string.Empty;
        context.Fields = ExtractFields(controls);
        context.Grids = ExtractGrids(controls);
        context.PrimaryGridTitle = context.Grids.FirstOrDefault()?.Title ?? string.Empty;
        context.PrimaryGridSummary = context.Grids.FirstOrDefault() == null ? "no grid detected" : BuildGridSummary(context);
        context.Actions = ExtractActions(controls);
        context.ValidationSummary = InferValidationSummary(title, controls);
        return context;
    }

    private static string WaitForCondition(Func<AxContextSnapshot, bool> predicate, string label)
    {
        for (var attempt = 0; attempt < 10; attempt++)
        {
            var snapshot = new AxClientAutomationService().CaptureActiveContext();
            if (predicate(snapshot))
            {
                return $"wait satisfied for {label}";
            }

            Thread.Sleep(200);
        }

        return $"wait failed for {label}";
    }

    private static string ClickNamedAction(AxContextSnapshot context, IEnumerable<string> names)
    {
        var handle = GetForegroundWindow();
        foreach (var name in names.Where(name => !string.IsNullOrWhiteSpace(name)))
        {
            var control = FindControlByText(handle, name, "button");
            if (control == null)
            {
                continue;
            }

            SetForegroundWindow(handle);
            SetFocus(control.Handle);
            SendMessage(control.Handle, BmClick, IntPtr.Zero, IntPtr.Zero);
            return $"clicked AX action {name}";
        }

        return $"action not found: {string.Join(" / ", names)}";
    }

    private static string BuildGridSummary(AxContextSnapshot context)
    {
        if (context.Grids.Count == 0)
        {
            return "no AX grid detected";
        }

        var lines = new List<string>();
        foreach (var grid in context.Grids.Take(3))
        {
            lines.Add($"grid: {grid.Title} ({grid.ClassName})");
            if (grid.Headers.Count > 0)
            {
                lines.Add("headers: " + string.Join(", ", grid.Headers.Take(8)));
            }

            if (grid.VisibleRows.Count > 0)
            {
                lines.AddRange(grid.VisibleRows.Take(5).Select(row => $"- {row}"));
            }

            if (!string.IsNullOrWhiteSpace(grid.SelectedRow))
            {
                lines.Add($"selected: {grid.SelectedRow}");
            }
        }

        return string.Join(Environment.NewLine, lines);
    }

    private static string BuildFormSummary(AxContextSnapshot context)
    {
        var lines = new List<string>
        {
            $"form: {(string.IsNullOrWhiteSpace(context.FormName) ? "<unknown>" : context.FormName)}"
        };

        if (!string.IsNullOrWhiteSpace(context.DialogTitle))
        {
            lines.Add($"dialog: {context.DialogTitle}");
        }

        if (!string.IsNullOrWhiteSpace(context.ActiveTab))
        {
            lines.Add($"tab: {context.ActiveTab}");
        }

        if (context.Fields.Count == 0)
        {
            lines.Add("no readable AX fields detected");
        }
        else
        {
            lines.AddRange(context.Fields.Take(18).Select(field =>
                $"- {field.Label}: {(string.IsNullOrWhiteSpace(field.Value) ? "<empty>" : field.Value)}"));
        }

        if (context.Actions.Count > 0)
        {
            lines.Add("actions: " + string.Join(", ", context.Actions.Take(10)));
        }

        return string.Join(Environment.NewLine, lines);
    }

    private static AxFieldSnapshot? FindBestField(AxContextSnapshot context, string label)
    {
        if (string.IsNullOrWhiteSpace(label))
        {
            return null;
        }

        return context.Fields
            .Select(field => new
            {
                Field = field,
                Score = ScoreText(field.Label, label) + (field.IsEditable ? 1 : 0)
            })
            .Where(entry => entry.Score > 0)
            .OrderByDescending(entry => entry.Score)
            .Select(entry => entry.Field)
            .FirstOrDefault();
    }

    private static string? FindBestTab(AxContextSnapshot context, string label)
    {
        if (string.IsNullOrWhiteSpace(label))
        {
            return null;
        }

        return context.Tabs
            .Select(tab => new { Tab = tab, Score = ScoreText(tab, label) })
            .Where(entry => entry.Score > 0)
            .OrderByDescending(entry => entry.Score)
            .Select(entry => entry.Tab)
            .FirstOrDefault();
    }

    private static string GetProcessName(IntPtr handle)
    {
        GetWindowThreadProcessId(handle, out var processId);
        if (processId == 0)
        {
            return string.Empty;
        }

        try
        {
            using var process = System.Diagnostics.Process.GetProcessById((int)processId);
            return process.ProcessName;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static bool IsAxProcess(string processName, string title, string className)
    {
        return processName.Contains("ax32", StringComparison.OrdinalIgnoreCase) ||
               title.Contains("Dynamics AX", StringComparison.OrdinalIgnoreCase) ||
               title.Contains("Microsoft Dynamics AX", StringComparison.OrdinalIgnoreCase) ||
               className.Contains("Ax", StringComparison.OrdinalIgnoreCase);
    }

    private static string InferWindowType(string title, string className, IReadOnlyList<ControlSnapshot> controls)
    {
        if (title.Contains("lookup", StringComparison.OrdinalIgnoreCase))
        {
            return "lookup dialog";
        }

        if (title.Contains("warning", StringComparison.OrdinalIgnoreCase) ||
            title.Contains("error", StringComparison.OrdinalIgnoreCase) ||
            title.Contains("validation", StringComparison.OrdinalIgnoreCase))
        {
            return "message / validation dialog";
        }

        if (controls.Any(control => control.ClassName.Contains("SysTabControl32", StringComparison.OrdinalIgnoreCase)))
        {
            return "workspace window";
        }

        if (controls.Any(control => control.ClassName.Contains("Grid", StringComparison.OrdinalIgnoreCase) ||
                                    control.ClassName.Contains("List", StringComparison.OrdinalIgnoreCase)))
        {
            return "grid/list page";
        }

        if (className.Contains("#32770", StringComparison.OrdinalIgnoreCase))
        {
            return "modal dialog";
        }

        return "workspace window";
    }

    private static string InferFormName(string title, IReadOnlyList<ControlSnapshot> controls)
    {
        if (!string.IsNullOrWhiteSpace(title))
        {
            var fragments = title.Split(['-', '|'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Where(fragment =>
                    !fragment.Contains("Dynamics AX", StringComparison.OrdinalIgnoreCase) &&
                    !fragment.Contains("Microsoft", StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (fragments.Count > 0)
            {
                return fragments[0];
            }
        }

        return controls
            .Where(control => control.ClassName.Contains("Static", StringComparison.OrdinalIgnoreCase))
            .Select(control => control.Text)
            .FirstOrDefault(text => text.Length > 2) ?? string.Empty;
    }

    private static List<string> ExtractTabs(IReadOnlyList<ControlSnapshot> controls)
    {
        return controls
            .Where(control =>
                control.Text.Length > 0 &&
                (control.ClassName.Contains("Tab", StringComparison.OrdinalIgnoreCase) ||
                 control.ClassName.Contains("Button", StringComparison.OrdinalIgnoreCase)))
            .Select(control => NormalizeText(control.Text))
            .Where(text => text.Length > 1)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(12)
            .ToList();
    }

    private static List<string> ExtractActions(IReadOnlyList<ControlSnapshot> controls)
    {
        return controls
            .Where(control => control.Text.Length > 0 &&
                              (control.ClassName.Contains("Button", StringComparison.OrdinalIgnoreCase) ||
                               control.ClassName.Contains("Toolbar", StringComparison.OrdinalIgnoreCase)))
            .Select(control => NormalizeText(control.Text))
            .Where(text => text.Length > 1)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(16)
            .ToList();
    }

    private static List<AxFieldSnapshot> ExtractFields(IReadOnlyList<ControlSnapshot> controls)
    {
        var labels = controls
            .Where(control =>
                control.Text.Length > 0 &&
                control.ClassName.Contains("Static", StringComparison.OrdinalIgnoreCase))
            .OrderBy(control => control.Top)
            .ThenBy(control => control.Left)
            .ToList();

        var inputs = controls
            .Where(control => IsEditableClass(control.ClassName))
            .OrderBy(control => control.Top)
            .ThenBy(control => control.Left)
            .ToList();

        var results = new List<AxFieldSnapshot>();
        foreach (var label in labels)
        {
            var match = inputs
                .Where(input => input.Left >= label.Left && Math.Abs(input.Top - label.Top) <= 28)
                .OrderBy(input => Math.Abs(input.Top - label.Top))
                .ThenBy(input => Math.Abs(input.Left - label.Left))
                .FirstOrDefault();
            if (match == null)
            {
                continue;
            }

            results.Add(new AxFieldSnapshot
            {
                Handle = match.Handle,
                Label = NormalizeText(label.Text),
                Value = NormalizeText(match.Text),
                ClassName = match.ClassName,
                IsEditable = true,
                Left = match.Left,
                Top = match.Top
            });
        }

        return results
            .GroupBy(field => field.Label, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .Take(24)
            .ToList();
    }

    private static List<AxGridSnapshot> ExtractGrids(IReadOnlyList<ControlSnapshot> controls)
    {
        return controls
            .Where(control =>
                control.ClassName.Contains("Grid", StringComparison.OrdinalIgnoreCase) ||
                control.ClassName.Contains("List", StringComparison.OrdinalIgnoreCase) ||
                control.ClassName.Contains("Data", StringComparison.OrdinalIgnoreCase))
            .Select(control => new AxGridSnapshot
            {
                Handle = control.Handle,
                Title = string.IsNullOrWhiteSpace(control.Text) ? control.ClassName : NormalizeText(control.Text),
                ClassName = control.ClassName,
                Headers = GuessHeaders(controls, control),
                VisibleRows = GuessRows(controls, control),
                SelectedRow = GuessRows(controls, control).FirstOrDefault() ?? string.Empty
            })
            .Take(6)
            .ToList();
    }

    private static List<string> GuessHeaders(IReadOnlyList<ControlSnapshot> controls, ControlSnapshot grid)
    {
        return controls
            .Where(control =>
                control.Text.Length > 0 &&
                Math.Abs(control.Top - grid.Top) <= 24 &&
                control.Left >= grid.Left - 16)
            .Select(control => NormalizeText(control.Text))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(8)
            .ToList();
    }

    private static List<string> GuessRows(IReadOnlyList<ControlSnapshot> controls, ControlSnapshot grid)
    {
        return controls
            .Where(control =>
                control.Text.Length > 0 &&
                control.Top > grid.Top &&
                control.Top <= grid.Top + 220 &&
                control.Left >= grid.Left - 12)
            .Select(control => NormalizeText(control.Text))
            .Where(text => text.Length > 1)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(10)
            .ToList();
    }

    private static string InferValidationSummary(string title, IReadOnlyList<ControlSnapshot> controls)
    {
        if (title.Contains("warning", StringComparison.OrdinalIgnoreCase) ||
            title.Contains("error", StringComparison.OrdinalIgnoreCase))
        {
            return NormalizeText(title);
        }

        var validationText = controls
            .Where(control =>
                control.Text.Length > 0 &&
                control.ClassName.Contains("Static", StringComparison.OrdinalIgnoreCase))
            .Select(control => NormalizeText(control.Text))
            .FirstOrDefault(text =>
                text.Contains("error", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("warning", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("must", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("required", StringComparison.OrdinalIgnoreCase));

        return validationText ?? string.Empty;
    }

    private static bool IsEditableClass(string className)
    {
        return className.Contains("Edit", StringComparison.OrdinalIgnoreCase) ||
               className.Contains("ComboBox", StringComparison.OrdinalIgnoreCase) ||
               className.Contains("RichEdit", StringComparison.OrdinalIgnoreCase) ||
               className.Contains("WindowsForms10.EDIT", StringComparison.OrdinalIgnoreCase);
    }

    private static int ScoreText(string source, string query)
    {
        if (string.IsNullOrWhiteSpace(source) || string.IsNullOrWhiteSpace(query))
        {
            return 0;
        }

        var normalizedSource = NormalizeText(source);
        var normalizedQuery = NormalizeText(query);
        if (normalizedSource.Equals(normalizedQuery, StringComparison.OrdinalIgnoreCase))
        {
            return 10;
        }

        if (normalizedSource.Contains(normalizedQuery, StringComparison.OrdinalIgnoreCase))
        {
            return 6;
        }

        return normalizedQuery.Length >= 3 && normalizedSource.Replace(" ", string.Empty).Contains(normalizedQuery.Replace(" ", string.Empty), StringComparison.OrdinalIgnoreCase)
            ? 3
            : 0;
    }

    private static ControlSnapshot? FindControlByText(IntPtr parentHandle, string text, string preferredKind)
    {
        var controls = EnumerateVisibleChildControls(parentHandle);
        return controls
            .Select(control => new
            {
                Control = control,
                Score = ScoreText(control.Text, text) +
                        (preferredKind == "button" && control.ClassName.Contains("Button", StringComparison.OrdinalIgnoreCase) ? 2 : 0) +
                        (preferredKind == "tab" && control.ClassName.Contains("Tab", StringComparison.OrdinalIgnoreCase) ? 2 : 0)
            })
            .Where(entry => entry.Score > 0)
            .OrderByDescending(entry => entry.Score)
            .Select(entry => entry.Control)
            .FirstOrDefault();
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
            var text = NormalizeText(GetWindowTextValue(handle));
            GetWindowRect(handle, out var rect);
            if (string.IsNullOrWhiteSpace(className) && string.IsNullOrWhiteSpace(text))
            {
                return true;
            }

            controls.Add(new ControlSnapshot
            {
                Handle = handle,
                ClassName = className,
                Text = text,
                Left = rect.Left,
                Top = rect.Top
            });
            return true;
        }, IntPtr.Zero);
        return controls;
    }

    private static string NormalizeText(string text)
    {
        return (text ?? string.Empty).Replace("&", string.Empty).Replace(Environment.NewLine, " ").Trim();
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
        public int Left { get; init; }
        public int Top { get; init; }
    }

    private delegate bool EnumChildProc(IntPtr windowHandle, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    private struct Rect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

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

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out Rect rect);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);

    [DllImport("user32.dll")]
    private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
}
