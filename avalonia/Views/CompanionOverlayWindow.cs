using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;

namespace ClippyRWAvalonia.Views;

public sealed class CompanionOverlayWindow : Window
{
    private readonly Border _tail;
    private readonly Border _orb;
    private readonly Border _orbHalo;
    private readonly Border[] _activityBars;
    private readonly ScaleTransform _haloScaleTransform;
    private readonly ScaleTransform _orbScaleTransform;
    private readonly ScaleTransform _tailScaleTransform;
    private readonly RotateTransform _haloRotateTransform;
    private readonly RotateTransform _orbRotateTransform;
    private readonly RotateTransform _tailRotateTransform;
    private readonly Border _stateBadge;
    private readonly Border _riskBadge;
    private readonly TextBlock _titleBlock;
    private readonly TextBlock _messageBlock;
    private readonly Border _chrome;
    private readonly TextBlock _stateHintBlock;
    private readonly TextBlock _riskTextBlock;
    private readonly DispatcherTimer _animationTimer;
    private string _currentState = "ready";
    private double _phase;

    public CompanionOverlayWindow()
    {
        Width = 332;
        Height = 128;
        CanResize = false;
        ShowInTaskbar = false;
        Topmost = true;
        WindowDecorations = WindowDecorations.None;
        Background = Avalonia.Media.Brushes.Transparent;
        TransparencyLevelHint = [WindowTransparencyLevel.Transparent];
        ShowActivated = false;

        _tail = new Border
        {
            Width = 30,
            Height = 12,
            CornerRadius = new CornerRadius(999),
            Background = new SolidColorBrush(Avalonia.Media.Color.Parse("#338CF0A5")),
            BorderBrush = Avalonia.Media.Brush.Parse("#38FFFFFF"),
            BorderThickness = new Thickness(1),
            HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Center,
            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
            Margin = new Thickness(0, 42, 0, 0),
            Opacity = 0.78
        };
        _tail.RenderTransformOrigin = new RelativePoint(0.2, 0.5, RelativeUnit.Relative);
        _tailScaleTransform = new ScaleTransform(1, 1);
        _tailRotateTransform = new RotateTransform(18);
        _tail.RenderTransform = new TransformGroup
        {
            Children =
            {
                _tailScaleTransform,
                _tailRotateTransform
            }
        };

        _orbHalo = new Border
        {
            Width = 58,
            Height = 58,
            CornerRadius = new CornerRadius(999),
            Background = new SolidColorBrush(Avalonia.Media.Color.Parse("#228CF0A5")),
            BorderBrush = Avalonia.Media.Brush.Parse("#2EFFFFFF"),
            BorderThickness = new Thickness(1),
            BoxShadow = new BoxShadows(new BoxShadow
            {
                Blur = 34,
                OffsetX = 0,
                OffsetY = 0,
                Spread = 0,
                Color = Avalonia.Media.Color.Parse("#448CF0A5")
            }),
            HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Center,
            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center
        };
        _orbHalo.RenderTransformOrigin = new RelativePoint(0.5, 0.5, RelativeUnit.Relative);
        _haloScaleTransform = new ScaleTransform(1, 1);
        _haloRotateTransform = new RotateTransform(0);
        _orbHalo.RenderTransform = new TransformGroup
        {
            Children =
            {
                _haloScaleTransform,
                _haloRotateTransform
            }
        };

        _orb = new Border
        {
            Width = 28,
            Height = 28,
            CornerRadius = new CornerRadius(999),
            Background = new LinearGradientBrush
            {
                StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
                EndPoint = new RelativePoint(1, 1, RelativeUnit.Relative),
                GradientStops =
                {
                    new GradientStop(Avalonia.Media.Color.Parse("#F6FFFE"), 0),
                    new GradientStop(Avalonia.Media.Color.Parse("#8CF0A5"), 0.55),
                    new GradientStop(Avalonia.Media.Color.Parse("#56C976"), 1)
                }
            },
            BorderBrush = Avalonia.Media.Brush.Parse("#E8FFF1"),
            BorderThickness = new Thickness(1.6),
            BoxShadow = new BoxShadows(new BoxShadow
            {
                Blur = 18,
                OffsetX = 0,
                OffsetY = 0,
                Spread = 0,
                Color = Avalonia.Media.Color.Parse("#8898FFC2")
            }),
            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center
        };
        _orb.RenderTransformOrigin = new RelativePoint(0.5, 0.5, RelativeUnit.Relative);
        _orbScaleTransform = new ScaleTransform(1, 1);
        _orbRotateTransform = new RotateTransform(0);
        _orb.RenderTransform = new TransformGroup
        {
            Children =
            {
                _orbScaleTransform,
                _orbRotateTransform
            }
        };

        _titleBlock = new TextBlock
        {
            Text = "ready",
            FontSize = 14,
            FontWeight = FontWeight.SemiBold,
            Foreground = Avalonia.Media.Brushes.White,
            LetterSpacing = 0.2
        };

        _stateHintBlock = new TextBlock
        {
            Text = "Karl Klammer",
            FontSize = 11,
            FontWeight = FontWeight.Medium,
            Foreground = Avalonia.Media.Brush.Parse("#9DB3D2")
        };

        _messageBlock = new TextBlock
        {
            Text = "operator surface ready",
            FontSize = 12,
            Foreground = Avalonia.Media.Brush.Parse("#D6DEEB"),
            TextWrapping = TextWrapping.Wrap
        };

        _stateBadge = new Border
        {
            CornerRadius = new CornerRadius(999),
            Padding = new Thickness(10, 4),
            HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Left,
            Background = Avalonia.Media.Brush.Parse("#1F2F46"),
            BorderBrush = Avalonia.Media.Brush.Parse("#38506F"),
            BorderThickness = new Thickness(1),
            Child = _stateHintBlock
        };

        _riskTextBlock = new TextBlock
        {
            Text = "safe",
            FontSize = 11,
            FontWeight = FontWeight.SemiBold,
            Foreground = Avalonia.Media.Brush.Parse("#DFFFE7")
        };

        _riskBadge = new Border
        {
            CornerRadius = new CornerRadius(999),
            Padding = new Thickness(10, 4),
            HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Right,
            Background = Avalonia.Media.Brush.Parse("#173524"),
            BorderBrush = Avalonia.Media.Brush.Parse("#2B7051"),
            BorderThickness = new Thickness(1),
            Child = _riskTextBlock
        };

        _activityBars =
        [
            CreateActivityBar(),
            CreateActivityBar(),
            CreateActivityBar()
        ];

        _chrome = new Border
        {
            CornerRadius = new CornerRadius(24),
            BorderThickness = new Thickness(1),
            BorderBrush = Avalonia.Media.Brush.Parse("#33507A"),
            Background = new LinearGradientBrush
            {
                StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
                EndPoint = new RelativePoint(1, 1, RelativeUnit.Relative),
                GradientStops =
                {
                    new GradientStop(Avalonia.Media.Color.Parse("#F0151E2B"), 0),
                    new GradientStop(Avalonia.Media.Color.Parse("#F01C2736"), 0.58),
                    new GradientStop(Avalonia.Media.Color.Parse("#F0121A25"), 1)
                }
            },
            Padding = new Thickness(18, 16),
            Child = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("Auto,*"),
                RowDefinitions = new RowDefinitions("Auto,*"),
                ColumnSpacing = 14,
                RowSpacing = 10,
                Children =
                {
                    new Grid
                    {
                        [Grid.ColumnSpanProperty] = 2,
                        ColumnDefinitions = new ColumnDefinitions("*,Auto"),
                        Children =
                        {
                            _stateBadge,
                            new ContentControl
                            {
                                [Grid.ColumnProperty] = 1,
                                Content = _riskBadge
                            }
                        }
                    },
                    new Grid
                    {
                        [Grid.RowProperty] = 1,
                        Width = 64,
                        Height = 64,
                        VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                        Children =
                        {
                            _tail,
                            _orbHalo,
                            _orb
                        }
                    },
                    new StackPanel
                    {
                        [Grid.RowProperty] = 1,
                        [Grid.ColumnProperty] = 1,
                        Spacing = 6,
                        VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                        Children =
                        {
                            _titleBlock,
                            _messageBlock,
                            new StackPanel
                            {
                                Orientation = Avalonia.Layout.Orientation.Horizontal,
                                Spacing = 5,
                                Children =
                                {
                                    _activityBars[0],
                                    _activityBars[1],
                                    _activityBars[2]
                                }
                            }
                        }
                    }
                }
            }
        };

        Content = new Grid
        {
            Margin = new Thickness(8),
            Children = { _chrome }
        };

        _animationTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(32)
        };
        _animationTimer.Tick += (_, _) => AnimateFrame();
        _animationTimer.Start();
    }

    public void ApplyState(string state, string message, string riskLevel = "safe")
    {
        var normalized = string.IsNullOrWhiteSpace(state) ? "ready" : state.Trim().ToLowerInvariant();
        _currentState = normalized;
        var colors = normalized switch
        {
            "listening" => ("#8CF0A5", "#5C8F6C", "#1D3421", "#E6FFF0", "#D6FFE2"),
            "transcribing" => ("#7FDBFF", "#43758D", "#173245", "#E5F8FF", "#CFEFFF"),
            "thinking" => ("#FFD36A", "#8B6D2B", "#3A2E13", "#FFF3D1", "#FFE6AB"),
            "speaking" => ("#F6A2FF", "#7D4B86", "#31163A", "#FFE8FF", "#F6D0FF"),
            "error" => ("#FF8C8C", "#8F4646", "#3A1717", "#FFEAEA", "#FFD1D1"),
            _ => ("#8CF0A5", "#33507A", "#182332", "#F5F7FA", "#D9FDE4")
        };

        _titleBlock.Text = normalized;
        _titleBlock.Foreground = Avalonia.Media.Brush.Parse(colors.Item4);
        _stateHintBlock.Text = normalized switch
        {
            "listening" => "Live audio",
            "transcribing" => "Speech to text",
            "thinking" => "Reasoning",
            "speaking" => "Voice output",
            "error" => "Needs attention",
            _ => "Karl Klammer"
        };
        _stateHintBlock.Foreground = Avalonia.Media.Brush.Parse(colors.Item5);
        _messageBlock.Text = string.IsNullOrWhiteSpace(message) ? "ready for the next task" : message.Trim();
        var isIdle = normalized == "ready";
        Width = isIdle ? 268 : 332;
        Height = isIdle ? 104 : 128;
        _messageBlock.MaxWidth = isIdle ? 158 : 210;
        ApplyOrbGeometry(normalized);
        _orb.Background = new LinearGradientBrush
        {
            StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
            EndPoint = new RelativePoint(1, 1, RelativeUnit.Relative),
            GradientStops =
            {
                new GradientStop(Avalonia.Media.Color.Parse("#F9FFFF"), 0),
                new GradientStop(Avalonia.Media.Color.Parse(colors.Item1), 0.58),
                new GradientStop(Avalonia.Media.Color.Parse(colors.Item1), 1)
            }
        };
        _orbHalo.Background = new SolidColorBrush(Avalonia.Media.Color.Parse("#22" + colors.Item1.TrimStart('#')));
        _orb.BorderBrush = Avalonia.Media.Brush.Parse(colors.Item4);
        _chrome.BorderBrush = Avalonia.Media.Brush.Parse(colors.Item2);
        _chrome.Background = new SolidColorBrush(Avalonia.Media.Color.Parse(colors.Item3));
        _stateBadge.BorderBrush = Avalonia.Media.Brush.Parse(colors.Item2);
        _stateBadge.Background = new SolidColorBrush(Avalonia.Media.Color.Parse("#33" + colors.Item1.TrimStart('#')));
        ApplyRisk(riskLevel);
        Opacity = isIdle ? 0.9 : 0.98;
    }

    private void ApplyRisk(string riskLevel)
    {
        var normalized = string.IsNullOrWhiteSpace(riskLevel) ? "safe" : riskLevel.Trim().ToLowerInvariant();
        var tuple = normalized switch
        {
            "high" => ("high risk", "#3D1717", "#8C4B4B", "#FFE1E1"),
            "medium" => ("medium risk", "#3B2A12", "#8E6E33", "#FFE7BA"),
            "low" => ("low risk", "#16301E", "#326847", "#D9FFE2"),
            _ => ("safe", "#173524", "#2B7051", "#DFFFE7")
        };

        _riskTextBlock.Text = tuple.Item1;
        _riskTextBlock.Foreground = Avalonia.Media.Brush.Parse(tuple.Item4);
        _riskBadge.Background = Avalonia.Media.Brush.Parse(tuple.Item2);
        _riskBadge.BorderBrush = Avalonia.Media.Brush.Parse(tuple.Item3);
    }

    private void ApplyOrbGeometry(string state)
    {
        switch (state)
        {
            case "listening":
                ApplyShape(_orbHalo, 70, 42, 21, 0);
                ApplyShape(_orb, 38, 24, 12, 0);
                ApplyTail(34, 12, 24, 18);
                break;
            case "transcribing":
                ApplyShape(_orbHalo, 46, 70, 23, 0);
                ApplyShape(_orb, 24, 38, 12, 0);
                ApplyTail(18, 22, 16, 58);
                break;
            case "thinking":
                ApplyShape(_orbHalo, 58, 58, 16, 45);
                ApplyShape(_orb, 28, 28, 8, 45);
                ApplyTail(26, 11, 999, -12);
                break;
            case "speaking":
                ApplyShape(_orbHalo, 76, 36, 18, 0);
                ApplyShape(_orb, 44, 18, 9, 0);
                ApplyTail(36, 10, 999, 10);
                break;
            case "error":
                ApplyShape(_orbHalo, 56, 56, 14, 0);
                ApplyShape(_orb, 28, 28, 7, 0);
                ApplyTail(24, 12, 6, 28);
                break;
            default:
                ApplyShape(_orbHalo, 58, 58, 29, 0);
                ApplyShape(_orb, 28, 28, 14, 0);
                ApplyTail(24, 10, 999, 12);
                break;
        }
    }

    private void ApplyShape(Border border, double width, double height, double radius, double rotation)
    {
        border.Width = width;
        border.Height = height;
        border.CornerRadius = new CornerRadius(radius);
        if (ReferenceEquals(border, _orbHalo))
        {
            _haloRotateTransform.Angle = rotation;
            return;
        }

        _orbRotateTransform.Angle = rotation;
    }

    private void ApplyTail(double width, double height, double radius, double rotation)
    {
        _tail.Width = width;
        _tail.Height = height;
        _tail.CornerRadius = new CornerRadius(radius);
        _tailRotateTransform.Angle = rotation;
    }

    private void AnimateFrame()
    {
        _phase += 0.095;
        var haloScale = 1.0;
        var orbScale = 1.0;
        var haloOpacity = 0.65;

        switch (_currentState)
        {
            case "listening":
                haloScale = 1.02 + Math.Sin(_phase * 1.6) * 0.16;
                orbScale = 1.0 + Math.Sin(_phase * 1.6 + 0.5) * 0.08;
                haloOpacity = 0.58 + (Math.Sin(_phase * 1.6) + 1) * 0.16;
                AnimateBars(0.4, 1.1, 1.5);
                _tailScaleTransform.ScaleX = 1.05 + Math.Abs(Math.Sin(_phase * 1.4)) * 0.22;
                _tailScaleTransform.ScaleY = 1.0 + Math.Abs(Math.Cos(_phase * 1.1)) * 0.08;
                break;
            case "thinking":
                haloScale = 1.02 + Math.Sin(_phase * 0.8) * 0.1;
                orbScale = 1.0 + Math.Cos(_phase * 1.1) * 0.05;
                haloOpacity = 0.52 + (Math.Sin(_phase) + 1) * 0.08;
                AnimateBars(0.45, 0.75, 1.05);
                _tailScaleTransform.ScaleX = 0.92 + Math.Abs(Math.Sin(_phase * 0.7)) * 0.15;
                _tailScaleTransform.ScaleY = 0.95 + Math.Abs(Math.Cos(_phase * 0.7)) * 0.1;
                break;
            case "transcribing":
                haloScale = 1.0 + Math.Sin(_phase * 1.2) * 0.09;
                orbScale = 0.98 + Math.Abs(Math.Sin(_phase * 1.4)) * 0.08;
                haloOpacity = 0.5 + (Math.Sin(_phase * 1.2) + 1) * 0.1;
                AnimateBars(0.55, 0.95, 1.25);
                _tailScaleTransform.ScaleX = 0.9 + Math.Abs(Math.Sin(_phase * 1.1)) * 0.12;
                _tailScaleTransform.ScaleY = 1.0 + Math.Abs(Math.Sin(_phase * 1.6)) * 0.18;
                break;
            case "speaking":
                haloScale = 1.0 + Math.Abs(Math.Sin(_phase * 1.8)) * 0.14;
                orbScale = 0.98 + Math.Abs(Math.Sin(_phase * 2.1)) * 0.12;
                haloOpacity = 0.55 + Math.Abs(Math.Sin(_phase * 1.8)) * 0.18;
                AnimateBars(0.75, 1.25, 1.8);
                _tailScaleTransform.ScaleX = 1.0 + Math.Abs(Math.Sin(_phase * 1.8)) * 0.2;
                _tailScaleTransform.ScaleY = 0.92 + Math.Abs(Math.Sin(_phase * 2.1)) * 0.12;
                break;
            case "error":
                haloScale = 1.0 + Math.Sin(_phase * 2.6) * 0.06;
                orbScale = 1.0 + Math.Sin(_phase * 2.6) * 0.03;
                haloOpacity = 0.68;
                AnimateBars(0.2, 0.2, 0.2);
                _tailScaleTransform.ScaleX = 1.0;
                _tailScaleTransform.ScaleY = 1.0;
                break;
            default:
                haloScale = 1.0 + Math.Sin(_phase * 0.6) * 0.04;
                orbScale = 1.0 + Math.Sin(_phase * 0.6) * 0.02;
                haloOpacity = 0.38 + (Math.Sin(_phase * 0.6) + 1) * 0.03;
                AnimateBars(0.16, 0.22, 0.18);
                _tailScaleTransform.ScaleX = 0.96 + Math.Abs(Math.Sin(_phase * 0.6)) * 0.06;
                _tailScaleTransform.ScaleY = 0.96 + Math.Abs(Math.Cos(_phase * 0.6)) * 0.04;
                break;
        }

        _haloScaleTransform.ScaleX = haloScale;
        _haloScaleTransform.ScaleY = haloScale;
        _orbScaleTransform.ScaleX = orbScale;
        _orbScaleTransform.ScaleY = orbScale;

        _orbHalo.Opacity = Math.Clamp(haloOpacity, 0.2, 0.95);
    }

    private void AnimateBars(double left, double middle, double right)
    {
        AnimateBar(_activityBars[0], left);
        AnimateBar(_activityBars[1], middle);
        AnimateBar(_activityBars[2], right);
    }

    private void AnimateBar(Border bar, double intensity)
    {
        var value = 0.35 + Math.Abs(Math.Sin(_phase * intensity)) * 0.65;
        bar.Height = 5 + (value * 10);
        bar.Opacity = 0.35 + (value * 0.65);
    }

    private static Border CreateActivityBar()
    {
        return new Border
        {
            Width = 5,
            Height = 7,
            CornerRadius = new CornerRadius(999),
            Background = Avalonia.Media.Brush.Parse("#9EDBFF"),
            Opacity = 0.6,
            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center
        };
    }
}
