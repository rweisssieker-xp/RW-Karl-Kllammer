namespace ClippyRWAvalonia.Models;

public sealed class AutomationRecipe
{
    public string Id { get; set; } = Guid.NewGuid().ToString("n");
    public string Name { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string Prompt { get; set; } = string.Empty;
    public string CompanionMode { get; set; } = "automation";
    public string Category { get; set; } = "general";
    public string GuardApp { get; set; } = string.Empty;
    public string GuardForm { get; set; } = string.Empty;
    public string GuardDialog { get; set; } = string.Empty;
    public string GuardTab { get; set; } = string.Empty;
    public string SourceType { get; set; } = "manual";
    public string RiskLevel { get; set; } = "low";
    public bool Enabled { get; set; } = true;
    public bool Archived { get; set; }
    public string CreatedAtUtc { get; set; } = string.Empty;
    public string LastRunUtc { get; set; } = string.Empty;
    public string LastSuccessUtc { get; set; } = string.Empty;
    public int RunCount { get; set; }
    public int FailureCount { get; set; }
    public string LastActiveApp { get; set; } = string.Empty;
    public string LastActiveForm { get; set; } = string.Empty;
    public int ConfidenceScore { get; set; }
    public int EstimatedMinutesSaved { get; set; }
    public List<string> KnowledgeSources { get; set; } = [];
    public List<string> Tags { get; set; } = [];
    public List<RitualParameter> Parameters { get; set; } = [];
    public List<RitualStep> Steps { get; set; } = [];

    public override string ToString() => Name;
}

public sealed class RitualStep
{
    public string ActionType { get; set; } = "app";
    public string ActionArgument { get; set; } = string.Empty;
    public int WaitMs { get; set; }
    public int RetryCount { get; set; }
    public string IfApp { get; set; } = string.Empty;
    public string IfForm { get; set; } = string.Empty;
    public string IfDialog { get; set; } = string.Empty;
    public string IfTab { get; set; } = string.Empty;
    public string OnFail { get; set; } = "stop";
    public string RiskLevel { get; set; } = "low";
    public string TargetLabel { get; set; } = string.Empty;
    public Dictionary<string, string> ParameterBindings { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    public override string ToString()
    {
        var parts = new List<string> { $"{ActionType}|{ActionArgument}" };
        if (WaitMs > 0)
        {
            parts.Add($"wait={WaitMs}");
        }

        if (RetryCount > 0)
        {
            parts.Add($"retry={RetryCount}");
        }

        if (!string.IsNullOrWhiteSpace(IfApp))
        {
            parts.Add($"ifapp={IfApp}");
        }

        if (!string.IsNullOrWhiteSpace(IfForm))
        {
            parts.Add($"if_form={IfForm}");
        }

        if (!string.IsNullOrWhiteSpace(IfDialog))
        {
            parts.Add($"if_dialog={IfDialog}");
        }

        if (!string.IsNullOrWhiteSpace(IfTab))
        {
            parts.Add($"if_tab={IfTab}");
        }

        if (!string.Equals(OnFail, "stop", StringComparison.OrdinalIgnoreCase))
        {
            parts.Add($"on_fail={OnFail}");
        }

        return string.Join(" | ", parts);
    }
}

public sealed class RitualParameter
{
    public string Name { get; set; } = string.Empty;
    public string Label { get; set; } = string.Empty;
    public string DefaultValue { get; set; } = string.Empty;
    public bool Required { get; set; } = true;
    public string Kind { get; set; } = "text";

    public override string ToString() => string.IsNullOrWhiteSpace(Label) ? Name : $"{Label} ({Name})";
}
