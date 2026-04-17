using Avalonia;
using Avalonia.Controls;
namespace ClippyRWAvalonia.Views;

public sealed class CompanionAnchorWindow : Window
{
    private readonly Border _ring;
    private readonly TextBlock _label;

    public CompanionAnchorWindow()
    {
        Width = 92;
        Height = 48;
        CanResize = false;
        ShowInTaskbar = false;
        Topmost = true;
        WindowDecorations = WindowDecorations.None;
        Background = Avalonia.Media.Brushes.Transparent;
        TransparencyLevelHint = [WindowTransparencyLevel.Transparent];
        ShowActivated = false;

        _ring = new Border
        {
            Width = 18,
            Height = 18,
            CornerRadius = new CornerRadius(999),
            BorderBrush = Avalonia.Media.Brush.Parse("#F6F7FB"),
            BorderThickness = new Thickness(2),
            Background = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#55A7E6FF")),
            BoxShadow = new Avalonia.Media.BoxShadows(new Avalonia.Media.BoxShadow
            {
                Blur = 16,
                Color = Avalonia.Media.Color.Parse("#66A7E6FF")
            })
        };

        _label = new TextBlock
        {
            Foreground = Avalonia.Media.Brush.Parse("#EAF2FD"),
            FontSize = 11,
            FontWeight = Avalonia.Media.FontWeight.Medium
        };

        Content = new Border
        {
            CornerRadius = new CornerRadius(16),
            Background = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#CC132132")),
            BorderBrush = Avalonia.Media.Brush.Parse("#3E6189"),
            BorderThickness = new Thickness(1),
            Padding = new Thickness(10, 8),
            Child = new StackPanel
            {
                Orientation = Avalonia.Layout.Orientation.Horizontal,
                Spacing = 8,
                Children = { _ring, _label }
            }
        };
    }

    public void Apply(string label, string colorHex)
    {
        var safe = string.IsNullOrWhiteSpace(colorHex) ? "#A7E6FF" : colorHex;
        _ring.Background = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#55" + safe.TrimStart('#')));
        _label.Text = string.IsNullOrWhiteSpace(label) ? "target" : label;
    }
}
