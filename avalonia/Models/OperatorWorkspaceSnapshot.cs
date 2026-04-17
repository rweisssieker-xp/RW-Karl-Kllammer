namespace ClippyRWAvalonia.Models;

public sealed class OperatorWorkspaceSnapshot
{
    public string RepoRoot { get; init; } = string.Empty;
    public string WindowsRoot { get; init; } = string.Empty;
    public string DataRoot { get; init; } = string.Empty;
    public string EnvPath { get; init; } = string.Empty;
    public bool EnvExists { get; init; }
    public string Provider { get; init; } = "anthropic";
    public string Model { get; init; } = string.Empty;
    public string Mode { get; init; } = "companion";
    public bool SpeakResponses { get; init; }
    public bool UseLocalKnowledge { get; init; }
    public bool SuggestAutomations { get; init; }
    public string RuntimeSummary { get; init; } = string.Empty;
    public string KnowledgeStatus { get; init; } = "knowledge: unknown";
    public IReadOnlyList<KnowledgeDocumentSummary> KnowledgeDocuments { get; init; } = Array.Empty<KnowledgeDocumentSummary>();
    public IReadOnlyList<AutomationRecipe> Recipes { get; init; } = Array.Empty<AutomationRecipe>();
    public IReadOnlyList<WatchSessionEntry> WatchSessions { get; init; } = Array.Empty<WatchSessionEntry>();
    public IReadOnlyList<ActionHistoryEntry> ActionHistory { get; init; } = Array.Empty<ActionHistoryEntry>();
    public IReadOnlyList<DiagnosticEntry> Diagnostics { get; init; } = Array.Empty<DiagnosticEntry>();
    public Dictionary<string, string> Settings { get; init; } = new(StringComparer.OrdinalIgnoreCase);
    public Dictionary<string, string> EnvValues { get; init; } = new(StringComparer.OrdinalIgnoreCase);
}
