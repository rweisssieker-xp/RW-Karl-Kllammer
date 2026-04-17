using System.Runtime.InteropServices;
using System.Text;

namespace ClippyRWAvalonia.Services;

public sealed class MicrophoneRecorderService : IDisposable
{
    private const string RecordingAlias = "clippyrwrec";
    private readonly string _dataRoot;
    private bool _isRecording;
    private string? _recordingPath;

    [DllImport("winmm.dll", CharSet = CharSet.Auto)]
    private static extern int mciSendString(string command, StringBuilder? returnValue, int returnLength, IntPtr callbackHandle);

    [DllImport("winmm.dll", CharSet = CharSet.Auto)]
    private static extern bool mciGetErrorString(int errorCode, StringBuilder errorText, int errorTextSize);

    public MicrophoneRecorderService(string dataRoot)
    {
        _dataRoot = dataRoot;
    }

    public bool IsRecording => _isRecording;

    public void Start()
    {
        if (_isRecording)
        {
            return;
        }

        Directory.CreateDirectory(_dataRoot);
        _recordingPath = Path.Combine(_dataRoot, $"avalonia-recording-{DateTime.Now:yyyyMMdd-HHmmss}.wav");
        CloseAliasQuietly();
        SendCommand($"open new type waveaudio alias {RecordingAlias}");

        try
        {
            SendCommand($"set {RecordingAlias} time format ms");
            SendCommand($"record {RecordingAlias}");
            _isRecording = true;
        }
        catch
        {
            CloseAliasQuietly();
            _recordingPath = null;
            throw;
        }
    }

    public string Stop()
    {
        if (!_isRecording || string.IsNullOrWhiteSpace(_recordingPath))
        {
            throw new InvalidOperationException("No microphone recording is currently running.");
        }

        try
        {
            SendCommand($"stop {RecordingAlias}");
            SendCommand($"save {RecordingAlias} {QuotePath(_recordingPath)}");
            return _recordingPath;
        }
        finally
        {
            CloseAliasQuietly();
            _isRecording = false;
        }
    }

    public void Cancel()
    {
        if (!_isRecording)
        {
            return;
        }

        CloseAliasQuietly();
        _isRecording = false;
        _recordingPath = null;
    }

    public void Dispose()
    {
        Cancel();
    }

    private static void SendCommand(string command)
    {
        var errorCode = mciSendString(command, null, 0, IntPtr.Zero);
        if (errorCode == 0)
        {
            return;
        }

        var errorText = new StringBuilder(256);
        if (!mciGetErrorString(errorCode, errorText, errorText.Capacity))
        {
            errorText.Append("unknown MCI error");
        }

        throw new InvalidOperationException("Microphone capture failed: " + errorText);
    }

    private static void CloseAliasQuietly()
    {
        mciSendString("close " + RecordingAlias, null, 0, IntPtr.Zero);
    }

    private static string QuotePath(string path)
    {
        return "\"" + path.Replace("\"", string.Empty) + "\"";
    }
}
