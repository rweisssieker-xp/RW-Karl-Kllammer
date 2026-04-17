namespace ClippyRWAvalonia.Models;

public sealed class AxContextSnapshot
{
    public bool IsAvailable { get; set; }
    public bool IsAxClient { get; set; }
    public string ProcessName { get; set; } = string.Empty;
    public string WindowTitle { get; set; } = string.Empty;
    public string WindowClassName { get; set; } = string.Empty;
    public string WindowType { get; set; } = "unknown";
    public string FormName { get; set; } = string.Empty;
    public string DialogTitle { get; set; } = string.Empty;
    public string ActiveTab { get; set; } = string.Empty;
    public string PrimaryGridTitle { get; set; } = string.Empty;
    public string PrimaryGridSummary { get; set; } = string.Empty;
    public string ValidationSummary { get; set; } = string.Empty;
    public List<string> Tabs { get; set; } = [];
    public List<string> Actions { get; set; } = [];
    public List<AxFieldSnapshot> Fields { get; set; } = [];
    public List<AxGridSnapshot> Grids { get; set; } = [];

    public string Summary
    {
        get
        {
            if (!IsAvailable)
            {
                return "no active AX context";
            }

            var lines = new List<string>
            {
                $"window: {WindowTitle}",
                $"type: {WindowType}",
                $"form: {(string.IsNullOrWhiteSpace(FormName) ? "<unknown>" : FormName)}"
            };

            if (!string.IsNullOrWhiteSpace(DialogTitle))
            {
                lines.Add($"dialog: {DialogTitle}");
            }

            if (!string.IsNullOrWhiteSpace(ActiveTab))
            {
                lines.Add($"tab: {ActiveTab}");
            }

            if (Fields.Count > 0)
            {
                lines.Add($"fields: {Fields.Count}");
            }

            if (Grids.Count > 0)
            {
                lines.Add($"grids: {Grids.Count}");
            }

            if (Actions.Count > 0)
            {
                lines.Add("actions: " + string.Join(", ", Actions.Take(8)));
            }

            if (!string.IsNullOrWhiteSpace(ValidationSummary))
            {
                lines.Add($"validation: {ValidationSummary}");
            }

            return string.Join(Environment.NewLine, lines);
        }
    }
}

public sealed class AxFieldSnapshot
{
    public nint Handle { get; set; }
    public string Label { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
    public string ClassName { get; set; } = string.Empty;
    public bool IsEditable { get; set; }
    public int Left { get; set; }
    public int Top { get; set; }
}

public sealed class AxGridSnapshot
{
    public nint Handle { get; set; }
    public string Title { get; set; } = string.Empty;
    public string ClassName { get; set; } = string.Empty;
    public List<string> Headers { get; set; } = [];
    public List<string> VisibleRows { get; set; } = [];
    public string SelectedRow { get; set; } = string.Empty;
}
