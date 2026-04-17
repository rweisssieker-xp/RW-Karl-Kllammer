namespace ClippyRWAvalonia.Models;

public sealed class KnowledgeIndex
{
    public List<KnowledgeChunk> Chunks { get; set; } = new();
    public string IndexedAtUtc { get; set; } = string.Empty;
}
