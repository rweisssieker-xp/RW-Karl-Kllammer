using Avalonia;
using Avalonia.Threading;
using ClippyRWAvalonia.Models;
using ClippyRWAvalonia.Views;
using Forms = System.Windows.Forms;

namespace ClippyRWAvalonia.Services;

public sealed class CompanionOverlayService : IDisposable
{
    private readonly DispatcherTimer _followTimer;
    private readonly DispatcherTimer _messageResetTimer;
    private CompanionOverlayWindow? _window;
    private CompanionAnchorWindow? _anchorWindow;
    private string _state = "ready";
    private string _message = "operator surface ready";
    private string _riskLevel = "safe";
    private double _currentX;
    private double _currentY;
    private bool _hasPosition;
    private DockMode _dockMode = DockMode.Cursor;
    private PixelPoint? _targetPoint;
    private string _targetLabel = string.Empty;

    public CompanionOverlayService()
    {
        _followTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(33)
        };
        _followTimer.Tick += (_, _) => MoveNearCursor();

        _messageResetTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(5)
        };
        _messageResetTimer.Tick += (_, _) =>
        {
            _messageResetTimer.Stop();
            SetState(_state, string.Empty);
        };
    }

    public void Initialize()
    {
        if (_window != null)
        {
            return;
        }

        _window = new CompanionOverlayWindow();
        _window.Closed += (_, _) =>
        {
            _followTimer.Stop();
            _window = null;
        };
        _anchorWindow = new CompanionAnchorWindow();
        _anchorWindow.Show();
        _window.ApplyState(_state, _message, _riskLevel);
        _window.Show();
        MoveNearCursor();
        _followTimer.Start();
    }

    public void SetState(string state, string message, string riskLevel = "safe")
    {
        _state = string.IsNullOrWhiteSpace(state) ? "ready" : state.Trim().ToLowerInvariant();
        _message = message ?? string.Empty;
        _riskLevel = string.IsNullOrWhiteSpace(riskLevel) ? "safe" : riskLevel.Trim().ToLowerInvariant();
        _window?.ApplyState(_state, _message, _riskLevel);
    }

    public void ShowTransient(string state, string message, int milliseconds = 4200, string riskLevel = "safe")
    {
        SetState(state, message, riskLevel);
        _messageResetTimer.Stop();
        _messageResetTimer.Interval = TimeSpan.FromMilliseconds(Math.Max(800, milliseconds));
        _messageResetTimer.Start();
    }

    public void DockToActiveWindow(ActiveWindowInfo activeWindow)
    {
        if (activeWindow == null || !activeWindow.HasBounds)
        {
            _dockMode = DockMode.Cursor;
            return;
        }

        _dockMode = DockMode.ActiveWindow;
        _targetPoint = new PixelPoint(activeWindow.Left + activeWindow.Width - 24, activeWindow.Top + 24);
        _targetLabel = activeWindow.AppKind;
        UpdateAnchor("#A7E6FF");
    }

    public void ShowTargetAnchor(int x, int y, string label, string riskLevel = "low")
    {
        _dockMode = DockMode.Target;
        _targetPoint = new PixelPoint(x, y);
        _targetLabel = label;
        _riskLevel = riskLevel;
        UpdateAnchor(riskLevel == "high" ? "#FF8C8C" : riskLevel == "medium" ? "#FFD36A" : "#A7E6FF");
    }

    public void ClearTargetAnchor()
    {
        _dockMode = DockMode.Cursor;
        _targetPoint = null;
        _targetLabel = string.Empty;
        _anchorWindow?.Hide();
    }

    private void MoveNearCursor()
    {
        if (_window == null)
        {
            return;
        }

        var cursor = Forms.Cursor.Position;
        var screen = Forms.Screen.FromPoint(cursor);
        var workArea = screen.WorkingArea;
        var desiredX = cursor.X + 24;
        var desiredY = cursor.Y + 24;
        if (_dockMode != DockMode.Cursor && _targetPoint.HasValue)
        {
            desiredX = _targetPoint.Value.X + 18;
            desiredY = _targetPoint.Value.Y - 12;
        }
        var clampedX = Math.Min(Math.Max(workArea.Left + 8, desiredX), workArea.Right - (int)_window.Width - 8);
        var clampedY = Math.Min(Math.Max(workArea.Top + 8, desiredY), workArea.Bottom - (int)_window.Height - 8);
        if (!_hasPosition)
        {
            _currentX = clampedX;
            _currentY = clampedY;
            _hasPosition = true;
        }
        else
        {
            var lerp = _state == "ready" ? 0.18 : 0.27;
            _currentX += (clampedX - _currentX) * lerp;
            _currentY += (clampedY - _currentY) * lerp;
        }

        _window.Position = new PixelPoint((int)Math.Round(_currentX), (int)Math.Round(_currentY));
        if (_dockMode != DockMode.Cursor && _targetPoint.HasValue)
        {
            UpdateAnchor(_riskLevel == "high" ? "#FF8C8C" : _riskLevel == "medium" ? "#FFD36A" : "#A7E6FF");
        }
    }

    public void ShowActionPreview(string riskLevel, string actionText, ActiveWindowInfo? activeWindow = null)
    {
        SetState("thinking", actionText, riskLevel);
        if (activeWindow != null && activeWindow.HasBounds)
        {
            DockToActiveWindow(activeWindow);
            ShowTargetAnchor(activeWindow.CenterX, activeWindow.Top + 18, activeWindow.AppKind, riskLevel);
        }
    }

    private void UpdateAnchor(string colorHex)
    {
        if (_anchorWindow == null)
        {
            return;
        }

        if (!_targetPoint.HasValue)
        {
            _anchorWindow.Hide();
            return;
        }

        _anchorWindow.Apply(string.IsNullOrWhiteSpace(_targetLabel) ? "target" : _targetLabel, colorHex);
        _anchorWindow.Position = new PixelPoint(_targetPoint.Value.X - 18, _targetPoint.Value.Y - 18);
        _anchorWindow.Show();
    }

    public void Dispose()
    {
        _messageResetTimer.Stop();
        _followTimer.Stop();
        _window?.Close();
        _anchorWindow?.Close();
        _window = null;
        _anchorWindow = null;
    }

    private enum DockMode
    {
        Cursor,
        ActiveWindow,
        Target
    }
}
