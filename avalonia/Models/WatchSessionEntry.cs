namespace ClippyRWAvalonia.Models;

public sealed class WatchSessionEntry
{
    public string TimestampUtc { get; set; } = string.Empty;
    public string Prompt { get; set; } = string.Empty;
    public string AssistantResponse { get; set; } = string.Empty;
    public string Provider { get; set; } = string.Empty;
    public string Model { get; set; } = string.Empty;
    public string ScreenSummary { get; set; } = string.Empty;
    public string ActiveApp { get; set; } = string.Empty;
}
