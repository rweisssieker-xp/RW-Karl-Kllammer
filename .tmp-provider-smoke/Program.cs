using ClippyRWAvalonia.Services;

var workspace = new OperatorWorkspaceService();
var runtime = new AssistantRuntimeService(workspace);
var wavPath = @"C:\tmp\clippy_rw\.tmp-provider-smoke\elevenlabs-input.wav";

async Task ProbeAnthropicAsync()
{
    var candidates = new [] { "claude-3-7-sonnet-latest", "claude-sonnet-4-20250514", "claude-3-5-sonnet-latest" };
    foreach (var candidate in candidates)
    {
        try
        {
            var text = await runtime.SmokeTestAsync("anthropic", candidate);
            Console.WriteLine($"anthropic=ok model={candidate} reply={text}");
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"anthropic.try.fail model={candidate} type={ex.GetType().Name} msg={ex.Message}");
        }
    }
    Console.WriteLine("anthropic=fail");
}

async Task ProbeOpenAiAsync()
{
    try
    {
        var text = await runtime.SmokeTestAsync("openai", "");
        Console.WriteLine($"openai=ok reply={text}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"openai=fail type={ex.GetType().Name} msg={ex.Message}");
    }
}

async Task ProbeElevenLabsAsync()
{
    try
    {
        var ttsPath = await runtime.SynthesizeSpeechToFileAsync("Dies ist ein kurzer ElevenLabs Test.");
        Console.WriteLine($"elevenlabs.tts=ok file={ttsPath}");
        var transcript = await runtime.TranscribeAudioFileAsync(wavPath);
        Console.WriteLine($"elevenlabs.stt=ok transcript={transcript}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"elevenlabs=fail type={ex.GetType().Name} msg={ex.Message}");
    }
}

await ProbeAnthropicAsync();
await ProbeOpenAiAsync();
await ProbeElevenLabsAsync();
