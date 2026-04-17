namespace ClippyRWAvalonia.Models;

public sealed class ActionHistoryEntry
{
    public string TimestampUtc { get; set; } = string.Empty;
    public string ActionName { get; set; } = string.Empty;
    public string ActionArgument { get; set; } = string.Empty;
    public string TargetLabel { get; set; } = string.Empty;
    public string SpokenText { get; set; } = string.Empty;
    public string ActiveApp { get; set; } = string.Empty;
    public string Category { get; set; } = "general";
    public string Result { get; set; } = string.Empty;

    public override string ToString()
    {
        return $"{TimestampUtc} | {ActionName} | {ActiveApp}";
    }
}
