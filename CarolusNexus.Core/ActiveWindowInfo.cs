namespace CarolusNexus.Core;

public sealed class ActiveWindowInfo
{
    public string ProcessName { get; set; } = "unknown app";
    public string WindowTitle { get; set; } = string.Empty;
    public string WindowClassName { get; set; } = string.Empty;
    public string AppKind { get; set; } = "unknown";
    public string DesktopFramework { get; set; } = "unknown";
    public int Left { get; set; }
    public int Top { get; set; }
    public int Width { get; set; }
    public int Height { get; set; }

    public bool HasBounds => Width > 0 && Height > 0;
    public int CenterX => Left + (Width / 2);
    public int CenterY => Top + (Height / 2);

    public string DisplayName =>
        string.IsNullOrWhiteSpace(WindowTitle)
            ? ProcessName
            : $"{ProcessName} - {WindowTitle}";
}
