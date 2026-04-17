namespace CarolusNexus.Platform.Windows;

/// <summary>
/// Uses Windows SAPI (<c>SAPI.SpVoice</c>) when cloud TTS is unavailable — parity with legacy WinForms fallback.
/// </summary>
public static class WindowsSpeechFallback
{
    public static bool TrySpeak(string text)
    {
        if (string.IsNullOrWhiteSpace(text) || !OperatingSystem.IsWindows())
        {
            return false;
        }

        var compact = text.Trim();
        if (compact.Length > 8000)
        {
            compact = compact[..8000];
        }

        try
        {
            var voiceType = Type.GetTypeFromProgID("SAPI.SpVoice");
            if (voiceType == null)
            {
                return false;
            }

            dynamic voice = Activator.CreateInstance(voiceType)!;
            voice.Speak(compact, 1);
            return true;
        }
        catch
        {
            return false;
        }
    }
}
