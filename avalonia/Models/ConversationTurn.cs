namespace ClippyRWAvalonia.Models;

public sealed class ConversationTurn
{
    public string UserTranscript { get; set; } = string.Empty;
    public string AssistantResponse { get; set; } = string.Empty;
}
