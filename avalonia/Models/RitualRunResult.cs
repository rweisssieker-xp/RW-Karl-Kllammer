namespace ClippyRWAvalonia.Models;

public sealed class RitualRunResult
{
    public string Status { get; set; } = "idle";
    public string Summary { get; set; } = string.Empty;
    public List<string> LogLines { get; set; } = [];
    public int NextStepIndex { get; set; }
    public string LastResult { get; set; } = string.Empty;
}
