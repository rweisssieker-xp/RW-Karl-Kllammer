namespace ClippyRWAvalonia.Models;

public sealed class ScreenCapturePayload
{
    public int ScreenIndex { get; set; }
    public string Label { get; set; } = string.Empty;
    public string ImageBase64 { get; set; } = string.Empty;
    public int Left { get; set; }
    public int Top { get; set; }
    public int Width { get; set; }
    public int Height { get; set; }
}
