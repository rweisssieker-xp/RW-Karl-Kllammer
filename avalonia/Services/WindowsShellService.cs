using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ClippyRWAvalonia.Services;

public sealed class WindowsShellService : IDisposable
{
    private sealed class PushToTalkHotKeyListener : IDisposable
    {
        private const int WhKeyboardLl = 13;
        private const int WmKeyDown = 0x0100;
        private const int WmKeyUp = 0x0101;
        private const int WmSysKeyDown = 0x0104;
        private const int WmSysKeyUp = 0x0105;

        private delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        public event EventHandler? HotKeyPressed;
        public event EventHandler? HotKeyReleased;

        private readonly HookProc _hookProc;
        private readonly Keys _hotKey;
        private IntPtr _hookHandle;
        private bool _isPressed;

        public PushToTalkHotKeyListener(Keys hotKey)
        {
            _hotKey = hotKey;
            _hookProc = HookCallback;

            using var currentProcess = Process.GetCurrentProcess();
            using var currentModule = currentProcess.MainModule;
            var moduleHandle = currentModule == null ? IntPtr.Zero : GetModuleHandle(currentModule.ModuleName);
            _hookHandle = SetWindowsHookEx(WhKeyboardLl, _hookProc, moduleHandle, 0);
        }

        public bool IsRegistered => _hookHandle != IntPtr.Zero;

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                var message = wParam.ToInt32();
                var virtualKeyCode = Marshal.ReadInt32(lParam);

                if (virtualKeyCode == (int)_hotKey)
                {
                    if ((message == WmKeyDown || message == WmSysKeyDown) && !_isPressed)
                    {
                        _isPressed = true;
                        HotKeyPressed?.Invoke(this, EventArgs.Empty);
                    }
                    else if ((message == WmKeyUp || message == WmSysKeyUp) && _isPressed)
                    {
                        _isPressed = false;
                        HotKeyReleased?.Invoke(this, EventArgs.Empty);
                    }
                }
            }

            return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
        }

        public void Dispose()
        {
            if (_hookHandle != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookHandle);
                _hookHandle = IntPtr.Zero;
            }
        }
    }

    private readonly Dictionary<string, string> _env;
    private NotifyIcon? _notifyIcon;
    private ContextMenuStrip? _trayMenu;
    private PushToTalkHotKeyListener? _listener;

    public WindowsShellService(Dictionary<string, string> env)
    {
        _env = env;
    }

    public event EventHandler? OpenRequested;
    public event EventHandler? AskRequested;
    public event EventHandler? HotKeyPressed;
    public event EventHandler? HotKeyReleased;
    public event EventHandler? ExitRequested;

    public void Initialize()
    {
        if (!OperatingSystem.IsWindows())
        {
            return;
        }

        _notifyIcon = new NotifyIcon
        {
            Icon = System.Drawing.SystemIcons.Information,
            Text = "Carolus Nexus",
            Visible = true
        };
        _notifyIcon.DoubleClick += (_, _) => OpenRequested?.Invoke(this, EventArgs.Empty);

        _trayMenu = new ContextMenuStrip();
        var openMenuItem = _trayMenu.Items.Add("Open Carolus Nexus");
        var askMenuItem = _trayMenu.Items.Add("Ask from current prompt");
        _trayMenu.Items.Add("-");
        var quitMenuItem = _trayMenu.Items.Add("Quit");

        openMenuItem.Click += (_, _) => OpenRequested?.Invoke(this, EventArgs.Empty);
        askMenuItem.Click += (_, _) => AskRequested?.Invoke(this, EventArgs.Empty);
        quitMenuItem.Click += (_, _) => ExitRequested?.Invoke(this, EventArgs.Empty);
        _notifyIcon.ContextMenuStrip = _trayMenu;

        if (TryParsePushToTalkKey(out var hotKey))
        {
            _listener = new PushToTalkHotKeyListener(hotKey);
            if (_listener.IsRegistered)
            {
                _listener.HotKeyPressed += (_, _) => HotKeyPressed?.Invoke(this, EventArgs.Empty);
                _listener.HotKeyReleased += (_, _) => HotKeyReleased?.Invoke(this, EventArgs.Empty);
            }
        }
    }

    public void UpdateTooltip(string text)
    {
        if (_notifyIcon == null)
        {
            return;
        }

        var compact = (text ?? string.Empty).Replace(Environment.NewLine, " ").Trim();
        if (compact.Length > 60)
        {
            compact = compact[..60];
        }

        _notifyIcon.Text = string.IsNullOrWhiteSpace(compact) ? "Carolus Nexus" : compact;
    }

    private bool TryParsePushToTalkKey(out Keys key)
    {
        var raw = _env.TryGetValue("PUSH_TO_TALK_KEY", out var configured) && !string.IsNullOrWhiteSpace(configured)
            ? configured
            : "F8";
        return Enum.TryParse(raw, true, out key);
    }

    public void Dispose()
    {
        if (_listener != null)
        {
            _listener.Dispose();
            _listener = null;
        }

        if (_notifyIcon != null)
        {
            _notifyIcon.Visible = false;
            _notifyIcon.Dispose();
            _notifyIcon = null;
        }

        if (_trayMenu != null)
        {
            _trayMenu.Dispose();
            _trayMenu = null;
        }
    }
}
