namespace ClippyRWAvalonia.Models;

public sealed class KnowledgeDocumentSummary
{
    public string Title { get; set; } = string.Empty;
    public string RelativePath { get; set; } = string.Empty;
    public int ChunkCount { get; set; }
    public string LastWriteUtc { get; set; } = string.Empty;
    public string Extension { get; set; } = string.Empty;

    public override string ToString()
    {
        return $"{Title} ({ChunkCount} chunks)";
    }
}
