namespace ClippyRWAvalonia.Models;

public sealed class DiagnosticEntry
{
    public string SourceFile { get; set; } = string.Empty;
    public string Line { get; set; } = string.Empty;
    public string TimestampHint { get; set; } = string.Empty;

    public override string ToString()
    {
        return string.IsNullOrWhiteSpace(TimestampHint) ? Line : $"{TimestampHint} | {Line}";
    }
}
