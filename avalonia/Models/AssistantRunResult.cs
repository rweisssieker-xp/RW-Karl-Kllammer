namespace ClippyRWAvalonia.Models;

public sealed class AssistantRunResult
{
    public string ResponseText { get; set; } = string.Empty;
    public string CleanResponseText { get; set; } = string.Empty;
    public string Provider { get; set; } = string.Empty;
    public string Model { get; set; } = string.Empty;
    public List<KnowledgeChunk> KnowledgeChunks { get; set; } = [];
    public List<ScreenCapturePayload> Screens { get; set; } = [];
    public AssistantActionPlan ActionPlan { get; set; } = new();
}
