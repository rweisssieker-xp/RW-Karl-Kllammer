namespace ClippyRWAvalonia.Models;

public sealed class AgentRunResult
{
    public string Agent { get; init; } = string.Empty;
    public string Prompt { get; init; } = string.Empty;
    public string WorkingDirectory { get; init; } = string.Empty;
    public string OutputFilePath { get; init; } = string.Empty;
    public int ExitCode { get; init; }
    public string StandardOutput { get; init; } = string.Empty;
    public string StandardError { get; init; } = string.Empty;
    public string ResponseText { get; init; } = string.Empty;
}
