using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Windows.Forms;
using ClippyRWAvalonia.Models;

namespace ClippyRWAvalonia.Services;

public sealed class AssistantRuntimeService
{
    private static readonly HttpClient HttpClient = new();
    private readonly OperatorWorkspaceService _workspaceService;

    public AssistantRuntimeService(OperatorWorkspaceService workspaceService)
    {
        _workspaceService = workspaceService;
    }

    public async Task<string> SmokeTestAsync(string provider, string model)
    {
        var env = _workspaceService.ReadEnvFile();
        var normalizedProvider = NormalizeProvider(provider);
        var requestModel = ResolveModel(normalizedProvider, model, env);

        if (normalizedProvider is "openai" or "openai-compatible")
        {
            var body = new JsonObject
            {
                ["model"] = requestModel,
                ["temperature"] = 0.0,
                ["messages"] = new JsonArray
                {
                    new JsonObject { ["role"] = "system", ["content"] = "reply with only the word ready" },
                    new JsonObject { ["role"] = "user", ["content"] = "say ready" }
                }
            };

            var response = await PostOpenAiCompatibleAsync(body, normalizedProvider, env).ConfigureAwait(false);
            return ExtractOpenAiText(response);
        }

        var anthropicBody = new JsonObject
        {
            ["model"] = requestModel,
            ["max_tokens"] = 24,
            ["stream"] = false,
            ["system"] = "reply with only the word ready",
            ["messages"] = new JsonArray
            {
                new JsonObject
                {
                    ["role"] = "user",
                    ["content"] = new JsonArray
                    {
                        new JsonObject
                        {
                            ["type"] = "text",
                            ["text"] = "say ready"
                        }
                    }
                }
            }
        };

        var anthropicResponse = await PostAnthropicAsync(anthropicBody, env).ConfigureAwait(false);
        return ExtractAnthropicText(anthropicResponse);
    }

    public async Task<AssistantRunResult> AskAsync(
        string provider,
        string model,
        string mode,
        bool suggestAutomations,
        string prompt,
        bool includeScreens,
        IReadOnlyList<ConversationTurn> conversationHistory,
        IReadOnlyList<KnowledgeChunk> knowledgeChunks)
    {
        var env = _workspaceService.ReadEnvFile();
        var normalizedProvider = NormalizeProvider(provider);
        var requestModel = ResolveModel(normalizedProvider, model, env);
        var screens = includeScreens ? CaptureAllScreens() : [];
        var recipes = _workspaceService.Load().Recipes;

        if (normalizedProvider is "openai" or "openai-compatible")
        {
            var messages = new JsonArray
            {
                new JsonObject
                {
                    ["role"] = "system",
                    ["content"] = BuildCompanionPrompt(mode, suggestAutomations, recipes)
                }
            };

            foreach (var turn in conversationHistory)
            {
                messages.Add(new JsonObject
                {
                    ["role"] = "user",
                    ["content"] = turn.UserTranscript
                });
                messages.Add(new JsonObject
                {
                    ["role"] = "assistant",
                    ["content"] = turn.AssistantResponse
                });
            }

            var content = new JsonArray();
            foreach (var screen in screens)
            {
                content.Add(new JsonObject
                {
                    ["type"] = "text",
                    ["text"] = screen.Label
                });
                content.Add(new JsonObject
                {
                    ["type"] = "image_url",
                    ["image_url"] = new JsonObject
                    {
                        ["url"] = $"data:image/jpeg;base64,{screen.ImageBase64}"
                    }
                });
            }

            content.Add(new JsonObject
            {
                ["type"] = "text",
                ["text"] = BuildUserPrompt(prompt, knowledgeChunks)
            });

            messages.Add(new JsonObject
            {
                ["role"] = "user",
                ["content"] = content
            });

            var openAiBody = new JsonObject
            {
                ["model"] = requestModel,
                ["temperature"] = 0.6,
                ["messages"] = messages
            };

            var response = await PostOpenAiCompatibleAsync(openAiBody, normalizedProvider, env).ConfigureAwait(false);
            var responseText = ExtractOpenAiText(response);
            var actionPlan = AssistantActionPlan.Parse(responseText);
            return new AssistantRunResult
            {
                Provider = normalizedProvider,
                Model = requestModel,
                KnowledgeChunks = knowledgeChunks.ToList(),
                Screens = screens.ToList(),
                ResponseText = responseText,
                CleanResponseText = string.IsNullOrWhiteSpace(actionPlan.CleanText) ? responseText : actionPlan.CleanText,
                ActionPlan = actionPlan
            };
        }

        var anthropicMessages = new JsonArray();
        foreach (var turn in conversationHistory)
        {
            anthropicMessages.Add(new JsonObject
            {
                ["role"] = "user",
                ["content"] = turn.UserTranscript
            });
            anthropicMessages.Add(new JsonObject
            {
                ["role"] = "assistant",
                ["content"] = turn.AssistantResponse
            });
        }

        var anthropicContent = new JsonArray();
        foreach (var screen in screens)
        {
            anthropicContent.Add(new JsonObject
            {
                ["type"] = "image",
                ["source"] = new JsonObject
                {
                    ["type"] = "base64",
                    ["media_type"] = "image/jpeg",
                    ["data"] = screen.ImageBase64
                }
            });
            anthropicContent.Add(new JsonObject
            {
                ["type"] = "text",
                ["text"] = screen.Label
            });
        }

        anthropicContent.Add(new JsonObject
        {
            ["type"] = "text",
            ["text"] = BuildUserPrompt(prompt, knowledgeChunks)
        });

        anthropicMessages.Add(new JsonObject
        {
            ["role"] = "user",
            ["content"] = anthropicContent
        });

        var anthropicBody = new JsonObject
        {
            ["model"] = requestModel,
            ["max_tokens"] = 1024,
            ["stream"] = false,
            ["system"] = BuildCompanionPrompt(mode, suggestAutomations, recipes),
            ["messages"] = anthropicMessages
        };

        var anthropicResponse = await PostAnthropicAsync(anthropicBody, env).ConfigureAwait(false);
        var anthropicText = ExtractAnthropicText(anthropicResponse);
        var anthropicPlan = AssistantActionPlan.Parse(anthropicText);
        return new AssistantRunResult
        {
            Provider = normalizedProvider,
            Model = requestModel,
            KnowledgeChunks = knowledgeChunks.ToList(),
            Screens = screens.ToList(),
            ResponseText = anthropicText,
            CleanResponseText = string.IsNullOrWhiteSpace(anthropicPlan.CleanText) ? anthropicText : anthropicPlan.CleanText,
            ActionPlan = anthropicPlan
        };
    }

    public async Task<string> TranscribeAudioFileAsync(string audioFilePath)
    {
        var env = _workspaceService.ReadEnvFile();
        var provider = GetValueOrDefault(env, "STT_PROVIDER", "whisper").Trim().ToLowerInvariant();
        if (provider == "elevenlabs")
        {
            return await TranscribeWithElevenLabsAsync(env, audioFilePath).ConfigureAwait(false);
        }

        return await TranscribeWithWhisperAsync(env, audioFilePath).ConfigureAwait(false);
    }

    public async Task<string> SynthesizeSpeechToFileAsync(string text)
    {
        var env = _workspaceService.ReadEnvFile();
        var apiKey = GetValueOrDefault(env, "ELEVENLABS_API_KEY", string.Empty);
        var voiceId = GetValueOrDefault(env, "ELEVENLABS_VOICE_ID", string.Empty);
        if (string.IsNullOrWhiteSpace(apiKey) || string.IsNullOrWhiteSpace(voiceId))
        {
            throw new InvalidOperationException("ELEVENLABS_API_KEY or ELEVENLABS_VOICE_ID is missing.");
        }

        var requestBody = new JsonObject
        {
            ["text"] = text,
            ["model_id"] = "eleven_flash_v2_5",
            ["voice_settings"] = new JsonObject
            {
                ["stability"] = 0.5,
                ["similarity_boost"] = 0.75
            }
        };

        using var request = new HttpRequestMessage(HttpMethod.Post, $"https://api.elevenlabs.io/v1/text-to-speech/{voiceId}");
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("audio/mpeg"));
        request.Headers.Add("xi-api-key", apiKey);
        request.Content = new StringContent(requestBody.ToJsonString(), Encoding.UTF8, "application/json");

        using var response = await HttpClient.SendAsync(request).ConfigureAwait(false);
        var bytes = await response.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"ElevenLabs TTS error ({(int)response.StatusCode}): {Encoding.UTF8.GetString(bytes)}");
        }

        var outputPath = Path.Combine(_workspaceService.DataRoot, $"avalonia-response-{DateTime.Now:yyyyMMdd-HHmmss}.mp3");
        Directory.CreateDirectory(_workspaceService.DataRoot);
        await File.WriteAllBytesAsync(outputPath, bytes).ConfigureAwait(false);
        return outputPath;
    }

    public void OpenFileWithShell(string path)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            throw new FileNotFoundException("Could not open generated file.", path);
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = path,
            UseShellExecute = true
        });
    }

    public bool IsElevenLabsVoiceConfigured()
    {
        var env = _workspaceService.ReadEnvFile();
        return !string.IsNullOrWhiteSpace(GetValueOrDefault(env, "ELEVENLABS_API_KEY", string.Empty))
            && !string.IsNullOrWhiteSpace(GetValueOrDefault(env, "ELEVENLABS_VOICE_ID", string.Empty));
    }

    /// <summary>JPEG files for Codex <c>-i</c>, same folder pattern as legacy WinForms.</summary>
    public IReadOnlyList<string> SaveScreenCapturesToCodexHandoffFolder()
    {
        var repoRoot = _workspaceService.RepoRoot;
        var directory = Path.Combine(repoRoot, "codex output", "screen captures");
        Directory.CreateDirectory(directory);
        var timestamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
        var paths = new List<string>();
        foreach (var cap in CaptureAllScreens())
        {
            var bytes = Convert.FromBase64String(cap.ImageBase64);
            var path = Path.Combine(directory, $"karl-klammer-codex-screen-{timestamp}-screen{cap.ScreenIndex}.jpg");
            File.WriteAllBytes(path, bytes);
            paths.Add(path);
        }

        return paths;
    }

    private static string NormalizeProvider(string provider)
    {
        var normalized = (provider ?? string.Empty).Trim().ToLowerInvariant();
        return normalized switch
        {
            "openai" => "openai",
            "openai-compatible" => "openai-compatible",
            _ => "anthropic"
        };
    }

    private static string ResolveModel(string provider, string model, Dictionary<string, string> env)
    {
        if (!string.IsNullOrWhiteSpace(model))
        {
            return model.Trim();
        }

        return provider switch
        {
            "openai" or "openai-compatible" => "gpt-4.1-mini",
            _ => GetValueOrDefault(env, "ANTHROPIC_MODEL", "claude-sonnet-4-20250514")
        };
    }

    private async Task<string> TranscribeWithElevenLabsAsync(Dictionary<string, string> env, string audioFilePath)
    {
        using var content = new MultipartFormDataContent();
        var bytes = await File.ReadAllBytesAsync(audioFilePath).ConfigureAwait(false);
        var audioContent = new ByteArrayContent(bytes);
        audioContent.Headers.ContentType = new MediaTypeHeaderValue("audio/wav");
        content.Add(audioContent, "file", Path.GetFileName(audioFilePath));
        content.Add(new StringContent("scribe_v2"), "model_id");

        var language = GetValueOrDefault(env, "WHISPER_LANGUAGE", string.Empty);
        if (!string.IsNullOrWhiteSpace(language))
        {
            content.Add(new StringContent(language), "language_code");
        }

        using var request = new HttpRequestMessage(HttpMethod.Post, "https://api.elevenlabs.io/v1/speech-to-text");
        request.Headers.Add("xi-api-key", GetValueOrDefault(env, "ELEVENLABS_API_KEY", string.Empty));
        request.Content = content;

        using var response = await HttpClient.SendAsync(request).ConfigureAwait(false);
        var responseText = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"ElevenLabs STT error ({(int)response.StatusCode}): {responseText}");
        }

        var transcript = ExtractSpeechToTextTranscript(responseText);
        if (string.IsNullOrWhiteSpace(transcript))
        {
            throw new InvalidOperationException("ElevenLabs returned an empty transcript.");
        }

        return transcript.Trim();
    }

    private async Task<string> TranscribeWithWhisperAsync(Dictionary<string, string> env, string audioFilePath)
    {
        var pythonCommand = GetValueOrDefault(env, "WHISPER_PYTHON", "python");
        var model = GetValueOrDefault(env, "WHISPER_MODEL", "base");
        var language = GetValueOrDefault(env, "WHISPER_LANGUAGE", "de");
        var outputDirectory = Path.Combine(_workspaceService.DataRoot, "whisper-output");
        Directory.CreateDirectory(outputDirectory);

        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = pythonCommand,
                Arguments = $"-m whisper \"{audioFilePath}\" --model \"{model}\" --language \"{language}\" --fp16 False --output_format txt --output_dir \"{outputDirectory}\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            }
        };

        if (!process.Start())
        {
            throw new InvalidOperationException("Failed to start local Whisper.");
        }

        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();
        await process.WaitForExitAsync().ConfigureAwait(false);
        var stdout = await stdoutTask.ConfigureAwait(false);
        var stderr = await stderrTask.ConfigureAwait(false);

        var transcriptPath = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(audioFilePath) + ".txt");
        if (process.ExitCode != 0)
        {
            throw new InvalidOperationException(string.IsNullOrWhiteSpace(stderr) ? stdout : stderr);
        }

        if (!File.Exists(transcriptPath))
        {
            throw new InvalidOperationException("Whisper finished but did not create a transcript file.");
        }

        return (await File.ReadAllTextAsync(transcriptPath).ConfigureAwait(false)).Trim();
    }

    private static string BuildCompanionPrompt(string mode, bool suggestAutomations, IReadOnlyList<AutomationRecipe> recipes)
    {
        var builder = new StringBuilder();
        builder.AppendLine(LoadSoulPrompt());
        builder.AppendLine();
        builder.AppendLine("You are Karl Klammer, the companion persona of Carolus Nexus, a direct desktop operator. Be concise, specific and useful.");
        builder.AppendLine("If screenshots are attached, reason from what is visible on screen.");
        builder.AppendLine($"Current mode: {(string.IsNullOrWhiteSpace(mode) ? "companion" : mode)}.");

        if (suggestAutomations)
        {
            builder.AppendLine("When you notice a repeatable workflow, you may suggest one reusable recipe.");
            builder.AppendLine("Append exactly one machine-readable tag only when useful: [AUTOMATION:name|short imperative prompt]");
        }
        builder.AppendLine("When you can identify a concrete desktop target, you may append one point tag: [POINT:screenIndex|xPercent|yPercent|short label].");
        builder.AppendLine("When a short desktop workflow is helpful, you may append one action chain: [ACTIONS:app|focus_window|ifapp=browser;app|focus_control:Search|retry=2|wait=250].");
        builder.AppendLine("Only use app actions that are short, reversible, and obvious from the current desktop state.");
        builder.AppendLine("For Microsoft Dynamics AX / AX 2012 style fat clients, prefer ax.* actions such as ax.read_field, ax.set_field, ax.click_action, ax.open_tab, ax.open_lookup, ax.confirm_lookup, ax.read_grid, ax.select_grid_row, ax.post, ax.confirm_dialog, ax.cancel_dialog, ax.wait_for_form, ax.wait_for_dialog and ax.wait_for_text.");
        builder.AppendLine("When local SOP or work-instruction snippets clearly describe an AX workflow, convert them into a cautious ax.* action plan instead of only summarizing the text.");
        builder.AppendLine("For PTC Creo, Babtec (B4/BCT), CATIA, Siemens NX, and other Win32 fat clients that are not Dynamics AX: use app| inspector-style actions (list_controls, read_form, focus_control, etc.) and |ifapp=creo|ifapp=babtec|ifapp=catia|ifapp=nx guards that match the active window; do not emit ax.* for those products.");

        if (recipes.Count > 0)
        {
            builder.AppendLine();
            builder.AppendLine("Known saved recipes:");
            foreach (var recipe in recipes.Take(5))
            {
                builder.AppendLine($"- {recipe.Name}: {recipe.Prompt}");
            }
        }

        return builder.ToString().Trim();
    }

    private static string BuildUserPrompt(string prompt, IReadOnlyList<KnowledgeChunk> knowledgeChunks)
    {
        var builder = new StringBuilder();
        builder.AppendLine((prompt ?? string.Empty).Trim());
        builder.AppendLine();
        builder.AppendLine("If you reference a target on screen, use [POINT:screenIndex|xPercent|yPercent|label] with percentages from 0 to 100.");
        builder.AppendLine("If you suggest a short desktop plan, use [ACTIONS:...] with optional |wait=250, |ifapp=word and |retry=2 directives on each step.");
        builder.AppendLine("For AX fat-client tasks, use ax.* step arguments inside [ACTIONS:...] and you may add |if_form=..., |if_dialog=..., |if_tab=... and |on_fail=skip.");
        builder.AppendLine("If local knowledge describes a Dynamics AX procedure, prefer emitting an ax.* plan that follows the SOP and keep the visible answer short.");
        builder.AppendLine("For Creo, Babtec, CATIA, NX, or other non-AX fat clients, use app| steps with appropriate |ifapp=... guards instead of ax.*.");

        if (knowledgeChunks.Count > 0)
        {
            builder.AppendLine();
            builder.AppendLine("local knowledge snippets:");
            foreach (var chunk in knowledgeChunks.Take(3))
            {
                builder.AppendLine($"[source: {chunk.Title}]");
                builder.AppendLine(chunk.Text);
                builder.AppendLine();
            }

            builder.AppendLine("Use the local knowledge when relevant and prefer it over guessing.");
        }

        return builder.ToString().Trim();
    }

    private static string LoadSoulPrompt()
    {
        try
        {
            var soulPath = Path.Combine(Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..")), "SOUL.md");
            if (File.Exists(soulPath))
            {
                var text = File.ReadAllText(soulPath, Encoding.UTF8).Trim();
                if (!string.IsNullOrWhiteSpace(text))
                {
                    return text;
                }
            }
        }
        catch
        {
        }

        return "Be genuinely helpful, skip filler, and speak like a sharp desktop operator.";
    }

    private static async Task<string> PostAnthropicAsync(JsonObject body, Dictionary<string, string> env)
    {
        var apiKey = GetValueOrDefault(env, "ANTHROPIC_API_KEY", string.Empty);
        if (string.IsNullOrWhiteSpace(apiKey))
        {
            throw new InvalidOperationException("ANTHROPIC_API_KEY is missing.");
        }

        using var request = new HttpRequestMessage(HttpMethod.Post, "https://api.anthropic.com/v1/messages");
        request.Headers.Add("x-api-key", apiKey);
        request.Headers.Add("anthropic-version", "2023-06-01");
        request.Content = new StringContent(body.ToJsonString(), Encoding.UTF8, "application/json");

        using var response = await HttpClient.SendAsync(request).ConfigureAwait(false);
        var responseText = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"Anthropic error ({(int)response.StatusCode}): {responseText}");
        }

        return responseText;
    }

    private static async Task<string> PostOpenAiCompatibleAsync(JsonObject body, string provider, Dictionary<string, string> env)
    {
        var apiKey = GetValueOrDefault(env, "OPENAI_API_KEY", string.Empty);
        if (string.IsNullOrWhiteSpace(apiKey))
        {
            throw new InvalidOperationException("OPENAI_API_KEY is missing.");
        }

        var baseUrl = GetValueOrDefault(env, "OPENAI_BASE_URL", "https://api.openai.com/v1").TrimEnd('/');
        using var request = new HttpRequestMessage(HttpMethod.Post, $"{baseUrl}/chat/completions");
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
        if (provider == "openai-compatible")
        {
            request.Headers.Add("HTTP-Referer", "https://clippyrw.local");
        request.Headers.Add("X-Title", "Carolus Nexus");
        }

        request.Content = new StringContent(body.ToJsonString(), Encoding.UTF8, "application/json");
        using var response = await HttpClient.SendAsync(request).ConfigureAwait(false);
        var responseText = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"OpenAI-compatible error ({(int)response.StatusCode}): {responseText}");
        }

        return responseText;
    }

    private static string ExtractAnthropicText(string responseBody)
    {
        var root = JsonNode.Parse(responseBody)?.AsObject() ?? throw new InvalidOperationException("Anthropic returned an invalid response.");
        var content = root["content"]?.AsArray() ?? throw new InvalidOperationException("Anthropic response did not contain content.");
        var parts = content
            .OfType<JsonObject>()
            .Where(item => string.Equals(item["type"]?.GetValue<string>(), "text", StringComparison.Ordinal))
            .Select(item => item["text"]?.GetValue<string>() ?? string.Empty)
            .Where(text => !string.IsNullOrWhiteSpace(text))
            .ToList();
        if (parts.Count == 0)
        {
            throw new InvalidOperationException("Anthropic response did not contain text blocks.");
        }

        return string.Join(Environment.NewLine, parts).Trim();
    }

    private static string ExtractOpenAiText(string responseBody)
    {
        var root = JsonNode.Parse(responseBody)?.AsObject() ?? throw new InvalidOperationException("OpenAI returned an invalid response.");
        var choices = root["choices"]?.AsArray() ?? throw new InvalidOperationException("OpenAI response did not contain choices.");
        var message = choices[0]?["message"]?.AsObject() ?? throw new InvalidOperationException("OpenAI response did not contain a message.");
        var contentNode = message["content"];
        if (contentNode == null)
        {
            throw new InvalidOperationException("OpenAI response message did not contain content.");
        }

        if (contentNode is JsonValue value)
        {
            return value.GetValue<string>().Trim();
        }

        if (contentNode is JsonArray array)
        {
            var parts = array
                .OfType<JsonObject>()
                .Where(item => string.Equals(item["type"]?.GetValue<string>(), "text", StringComparison.Ordinal))
                .Select(item => item["text"]?.GetValue<string>() ?? string.Empty)
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToList();
            if (parts.Count > 0)
            {
                return string.Join(Environment.NewLine, parts).Trim();
            }
        }

        throw new InvalidOperationException("OpenAI response content format was unsupported.");
    }

    private static string ExtractSpeechToTextTranscript(string responseText)
    {
        var root = JsonNode.Parse(responseText)?.AsObject();
        if (root == null)
        {
            return string.Empty;
        }

        if (root["text"] is JsonValue textValue)
        {
            return textValue.GetValue<string>();
        }

        if (root["words"] is not JsonArray words || words.Count == 0)
        {
            return string.Empty;
        }

        var parts = words
            .OfType<JsonObject>()
            .Select(word => word["text"]?.GetValue<string>() ?? string.Empty)
            .Where(text => !string.IsNullOrWhiteSpace(text));
        return string.Join(" ", parts).Trim();
    }

    private static IReadOnlyList<ScreenCapturePayload> CaptureAllScreens()
    {
        if (!OperatingSystem.IsWindows())
        {
            return [];
        }

        var cursor = Cursor.Position;
        var orderedScreens = Screen.AllScreens
            .OrderBy(screen => screen.Bounds.Contains(cursor) ? 0 : 1)
            .ToList();

        var captures = new List<ScreenCapturePayload>();
        for (var index = 0; index < orderedScreens.Count; index++)
        {
            var screen = orderedScreens[index];
            var bounds = screen.Bounds;
            using var bitmap = new Bitmap(bounds.Width, bounds.Height);
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.CopyFromScreen(bounds.Left, bounds.Top, 0, 0, bounds.Size);
            }

            using var stream = new MemoryStream();
            bitmap.Save(stream, ImageFormat.Jpeg);
            captures.Add(new ScreenCapturePayload
            {
                ScreenIndex = index + 1,
                Label = $"screen {index + 1} {(bounds.Contains(cursor) ? "(cursor screen)" : string.Empty)} {bounds.Width}x{bounds.Height}",
                ImageBase64 = Convert.ToBase64String(stream.ToArray()),
                Left = bounds.Left,
                Top = bounds.Top,
                Width = bounds.Width,
                Height = bounds.Height
            });
        }

        return captures;
    }

    private static string GetValueOrDefault(Dictionary<string, string> values, string key, string defaultValue)
    {
        return values.TryGetValue(key, out var value) && !string.IsNullOrWhiteSpace(value) ? value : defaultValue;
    }
}
