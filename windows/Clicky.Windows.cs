using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Accessibility;
using DocumentFormat.OpenXml.Packaging;
using UglyToad.PdfPig;

namespace ClickyWindows
{
    internal sealed class JavaScriptSerializer
    {
        public int MaxJsonLength { get; set; }

        public string Serialize(object value)
        {
            return JsonSerializer.Serialize(value);
        }

        public T Deserialize<T>(string json)
        {
            return JsonSerializer.Deserialize<T>(json, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });
        }

        public object DeserializeObject(string json)
        {
            using (JsonDocument document = JsonDocument.Parse(json))
            {
                return ConvertElement(document.RootElement);
            }
        }

        private static object ConvertElement(JsonElement element)
        {
            switch (element.ValueKind)
            {
                case JsonValueKind.Object:
                    Dictionary<string, object> dictionary = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                    foreach (JsonProperty property in element.EnumerateObject())
                    {
                        dictionary[property.Name] = ConvertElement(property.Value);
                    }
                    return dictionary;
                case JsonValueKind.Array:
                    ArrayList list = new ArrayList();
                    foreach (JsonElement item in element.EnumerateArray())
                    {
                        list.Add(ConvertElement(item));
                    }
                    return list;
                case JsonValueKind.String:
                    return element.GetString();
                case JsonValueKind.Number:
                    long longValue;
                    if (element.TryGetInt64(out longValue))
                    {
                        return longValue;
                    }

                    double doubleValue;
                    if (element.TryGetDouble(out doubleValue))
                    {
                        return doubleValue;
                    }

                    return element.GetRawText();
                case JsonValueKind.True:
                    return true;
                case JsonValueKind.False:
                    return false;
                case JsonValueKind.Null:
                case JsonValueKind.Undefined:
                    return null;
                default:
                    return element.GetRawText();
            }
        }
    }

    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainForm());
            }
            catch (Exception exception)
            {
                string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "startup-error.log");
                File.WriteAllText(logPath, exception.ToString(), Encoding.UTF8);
                MessageBox.Show(exception.ToString(), "Karl Klammer startup error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    internal sealed class AppSettings
    {
        public string ClaudeModel { get; set; }
        public string AssistantProvider { get; set; }
        public string AssistantModel { get; set; }
        public string CompanionMode { get; set; }
        public bool SuggestAutomations { get; set; }
        public bool UseLocalKnowledge { get; set; }
        public bool SpeakResponses { get; set; }
        public int MaxConversationTurns { get; set; }

        public static string StorageRoot
        {
            get { return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data"); }
        }

        public static string SettingsPath
        {
            get { return Path.Combine(StorageRoot, "settings.json"); }
        }

        public static AppSettings Load()
        {
            Directory.CreateDirectory(StorageRoot);

            if (!File.Exists(SettingsPath))
            {
                AppSettings defaults = CreateDefaults();
                defaults.Save();
                return defaults;
            }

            try
            {
                JavaScriptSerializer serializer = CreateSerializer();
                AppSettings loaded = serializer.Deserialize<AppSettings>(File.ReadAllText(SettingsPath, Encoding.UTF8));
                if (loaded == null)
                {
                    loaded = CreateDefaults();
                }

                if (string.IsNullOrWhiteSpace(loaded.ClaudeModel))
                {
                    loaded.ClaudeModel = "claude-sonnet-4-6";
                }

                if (string.IsNullOrWhiteSpace(loaded.AssistantProvider))
                {
                    loaded.AssistantProvider = "anthropic";
                }

                if (string.IsNullOrWhiteSpace(loaded.AssistantModel))
                {
                    loaded.AssistantModel = !string.IsNullOrWhiteSpace(loaded.ClaudeModel)
                        ? loaded.ClaudeModel
                        : GetDefaultModelForProvider(loaded.AssistantProvider);
                }

                if (string.IsNullOrWhiteSpace(loaded.CompanionMode))
                {
                    loaded.CompanionMode = "companion";
                }

                if (loaded.MaxConversationTurns <= 0)
                {
                    loaded.MaxConversationTurns = 10;
                }

                return loaded;
            }
            catch
            {
                AppSettings defaults = CreateDefaults();
                defaults.Save();
                return defaults;
            }
        }

        public void Save()
        {
            Directory.CreateDirectory(StorageRoot);
            JavaScriptSerializer serializer = CreateSerializer();
            string json = serializer.Serialize(this);
            File.WriteAllText(SettingsPath, json, Encoding.UTF8);
        }

        private static AppSettings CreateDefaults()
        {
            return new AppSettings
            {
                ClaudeModel = "claude-sonnet-4-6",
                AssistantProvider = "anthropic",
                AssistantModel = "claude-sonnet-4-6",
                CompanionMode = "companion",
                SuggestAutomations = true,
                UseLocalKnowledge = true,
                SpeakResponses = true,
                MaxConversationTurns = 10
            };
        }

        public static string GetDefaultModelForProvider(string provider)
        {
            string normalizedProvider = (provider ?? string.Empty).Trim().ToLowerInvariant();
            switch (normalizedProvider)
            {
                case "openai":
                    return "gpt-4o";
                case "openai-compatible":
                    return "llama-3.3-70b-instruct";
                default:
                    return "claude-sonnet-4-6";
            }
        }

        private static JavaScriptSerializer CreateSerializer()
        {
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            serializer.MaxJsonLength = int.MaxValue;
            return serializer;
        }
    }

    internal sealed class EnvironmentConfiguration
    {
        public string AnthropicApiKey { get; set; }
        public string OpenAIApiKey { get; set; }
        public string OpenAIBaseUrl { get; set; }
        public string ElevenLabsApiKey { get; set; }
        public string ElevenLabsVoiceId { get; set; }
        public string SpeechToTextProvider { get; set; }
        public string CodexCommand { get; set; }
        public string ClaudeCodeCommand { get; set; }
        public string CodexWorkingDirectory { get; set; }
        public int CodexTimeoutSeconds { get; set; }
        public string OpenClawCommand { get; set; }
        public string OpenClawGatewayUrl { get; set; }
        public string OpenClawGatewayToken { get; set; }
        public string OpenClawSessionKey { get; set; }
        public int OpenClawTimeoutSeconds { get; set; }
        public string WhisperPythonCommand { get; set; }
        public string WhisperModel { get; set; }
        public string WhisperLanguage { get; set; }
        public string PushToTalkKey { get; set; }

        public static string EnvFilePath
        {
            get { return GetPreferredEnvPath(); }
        }

        public static EnvironmentConfiguration Load()
        {
            Dictionary<string, string> envValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            if (File.Exists(EnvFilePath))
            {
                foreach (string rawLine in File.ReadAllLines(EnvFilePath, Encoding.UTF8))
                {
                    string trimmedLine = rawLine.Trim();
                    if (trimmedLine.Length == 0 || trimmedLine.StartsWith("#", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    int separatorIndex = trimmedLine.IndexOf('=');
                    if (separatorIndex <= 0)
                    {
                        continue;
                    }

                    string key = trimmedLine.Substring(0, separatorIndex).Trim();
                    string value = trimmedLine.Substring(separatorIndex + 1).Trim();

                    if (value.Length >= 2 && value.StartsWith("\"", StringComparison.Ordinal) && value.EndsWith("\"", StringComparison.Ordinal))
                    {
                        value = value.Substring(1, value.Length - 2);
                    }

                    envValues[key] = value;
                }
            }

            return new EnvironmentConfiguration
            {
                AnthropicApiKey = GetValueOrEmpty(envValues, "ANTHROPIC_API_KEY"),
                OpenAIApiKey = GetValueOrEmpty(envValues, "OPENAI_API_KEY"),
                OpenAIBaseUrl = GetValueOrDefault(envValues, "OPENAI_BASE_URL", "https://api.openai.com/v1"),
                ElevenLabsApiKey = GetValueOrEmpty(envValues, "ELEVENLABS_API_KEY"),
                ElevenLabsVoiceId = GetValueOrEmpty(envValues, "ELEVENLABS_VOICE_ID"),
                SpeechToTextProvider = GetSpeechToTextProvider(envValues),
                CodexCommand = GetValueOrDefault(envValues, "CODEX_COMMAND", "codex.cmd"),
                ClaudeCodeCommand = GetValueOrDefault(envValues, "CLAUDE_CODE_COMMAND", "claude"),
                CodexWorkingDirectory = GetValueOrDefault(envValues, "CODEX_WORKDIR", GetDefaultCodexWorkingDirectory()),
                CodexTimeoutSeconds = GetIntValueOrDefault(envValues, "CODEX_TIMEOUT_SECONDS", 900),
                OpenClawCommand = GetValueOrDefault(envValues, "OPENCLAW_COMMAND", "openclaw"),
                OpenClawGatewayUrl = GetValueOrDefault(envValues, "OPENCLAW_GATEWAY_URL", "ws://127.0.0.1:18789"),
                OpenClawGatewayToken = GetValueOrEmpty(envValues, "GATEWAY_TOKEN"),
                OpenClawSessionKey = GetValueOrDefault(envValues, "OPENCLAW_SESSION_KEY", "main"),
                OpenClawTimeoutSeconds = GetIntValueOrDefault(envValues, "OPENCLAW_TIMEOUT_SECONDS", 120),
                WhisperPythonCommand = GetValueOrDefault(envValues, "WHISPER_PYTHON", "python"),
                WhisperModel = GetValueOrDefault(envValues, "WHISPER_MODEL", "base"),
                WhisperLanguage = GetValueOrDefault(envValues, "WHISPER_LANGUAGE", "de"),
                PushToTalkKey = GetValueOrDefault(envValues, "PUSH_TO_TALK_KEY", "F8")
            };
        }

        public string Validate(string assistantProvider)
        {
            List<string> missingKeys = new List<string>();

            string normalizedProvider = (assistantProvider ?? string.Empty).Trim().ToLowerInvariant();
            if (normalizedProvider == "openai")
            {
                if (string.IsNullOrWhiteSpace(OpenAIApiKey))
                {
                    missingKeys.Add("OPENAI_API_KEY");
                }
            }
            else if (normalizedProvider == "openai-compatible")
            {
                if (string.IsNullOrWhiteSpace(OpenAIApiKey))
                {
                    missingKeys.Add("OPENAI_API_KEY");
                }

                if (string.IsNullOrWhiteSpace(OpenAIBaseUrl))
                {
                    missingKeys.Add("OPENAI_BASE_URL");
                }
            }
            else if (string.IsNullOrWhiteSpace(AnthropicApiKey))
            {
                missingKeys.Add("ANTHROPIC_API_KEY");
            }

            if (string.Equals(SpeechToTextProvider, "elevenlabs", StringComparison.OrdinalIgnoreCase)
                && string.IsNullOrWhiteSpace(ElevenLabsApiKey))
            {
                missingKeys.Add("ELEVENLABS_API_KEY");
            }

            if (missingKeys.Count == 0)
            {
                return null;
            }

            return ".env is missing: " + string.Join(", ", missingKeys);
        }

        private static string GetPreferredEnvPath()
        {
            List<string> candidatePaths = GetEnvCandidatePaths().ToList();

            foreach (string candidatePath in candidatePaths)
            {
                if (File.Exists(candidatePath))
                {
                    return candidatePath;
                }
            }

            return candidatePaths.First();
        }

        private static IEnumerable<string> GetEnvCandidatePaths()
        {
            List<string> paths = new List<string>();
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string windowsRoot = TryFindWindowsDirectory(baseDirectory)
                ?? TryFindWindowsDirectory(Directory.GetCurrentDirectory());

            if (!string.IsNullOrWhiteSpace(windowsRoot))
            {
                AddCandidatePath(paths, Path.Combine(windowsRoot, ".env"));
            }

            if (!string.IsNullOrWhiteSpace(baseDirectory))
            {
                AddCandidatePath(paths, Path.Combine(baseDirectory, ".env"));
                string current = Path.GetFullPath(baseDirectory);
                for (int i = 0; i < 6 && !string.IsNullOrWhiteSpace(current); i++)
                {
                    AddCandidatePath(paths, Path.Combine(current, ".env"));

                    string directoryName = Path.GetFileName(current.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
                    if (string.Equals(directoryName, "windows", StringComparison.OrdinalIgnoreCase))
                    {
                        break;
                    }

                    DirectoryInfo parent = Directory.GetParent(current);
                    if (parent == null)
                    {
                        break;
                    }

                    current = parent.FullName;
                }
            }

            string currentDirectory = Directory.GetCurrentDirectory();
            if (!string.IsNullOrWhiteSpace(currentDirectory))
            {
                AddCandidatePath(paths, Path.Combine(currentDirectory, ".env"));
            }

            return paths;
        }

        private static string TryFindWindowsDirectory(string startPath)
        {
            if (string.IsNullOrWhiteSpace(startPath))
            {
                return null;
            }

            string current = Path.GetFullPath(startPath);
            for (int i = 0; i < 8 && !string.IsNullOrWhiteSpace(current); i++)
            {
                string directoryName = Path.GetFileName(current.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
                if (string.Equals(directoryName, "windows", StringComparison.OrdinalIgnoreCase))
                {
                    return current;
                }

                DirectoryInfo parent = Directory.GetParent(current);
                if (parent == null)
                {
                    break;
                }

                current = parent.FullName;
            }

            return null;
        }

        private static void AddCandidatePath(ICollection<string> paths, string candidatePath)
        {
            if (string.IsNullOrWhiteSpace(candidatePath))
            {
                return;
            }

            string fullPath = Path.GetFullPath(candidatePath);
            if (!paths.Contains(fullPath, StringComparer.OrdinalIgnoreCase))
            {
                paths.Add(fullPath);
            }
        }

        private static string GetValueOrEmpty(IDictionary<string, string> values, string key)
        {
            string value;
            return values.TryGetValue(key, out value) ? value : string.Empty;
        }

        private static string GetValueOrDefault(IDictionary<string, string> values, string key, string defaultValue)
        {
            string value;
            return values.TryGetValue(key, out value) && !string.IsNullOrWhiteSpace(value) ? value : defaultValue;
        }

        private static string GetSpeechToTextProvider(IDictionary<string, string> values)
        {
            string provider = GetValueOrDefault(values, "STT_PROVIDER", "whisper").Trim().ToLowerInvariant();
            return provider == "elevenlabs" ? "elevenlabs" : "whisper";
        }

        private static int GetIntValueOrDefault(IDictionary<string, string> values, string key, int defaultValue)
        {
            string rawValue;
            int parsedValue;
            return values.TryGetValue(key, out rawValue) && int.TryParse(rawValue, out parsedValue) && parsedValue > 0
                ? parsedValue
                : defaultValue;
        }

        private static string GetDefaultCodexWorkingDirectory()
        {
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            return Path.Combine(Path.GetFullPath(Path.Combine(baseDirectory, "..")), "playground");
        }
    }

    internal sealed class ConversationTurn
    {
        public string UserTranscript { get; set; }
        public string AssistantResponse { get; set; }
    }

    internal sealed class AutomationRecipe
    {
        public string Name { get; set; }
        public string Prompt { get; set; }
        public string CompanionMode { get; set; }
        public string CreatedAtUtc { get; set; }

        public override string ToString()
        {
            return Name;
        }

        private static string StoragePath
        {
            get { return Path.Combine(AppSettings.StorageRoot, "automation-recipes.json"); }
        }

        public static List<AutomationRecipe> LoadAll()
        {
            Directory.CreateDirectory(AppSettings.StorageRoot);
            if (!File.Exists(StoragePath))
            {
                return new List<AutomationRecipe>();
            }

            try
            {
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                serializer.MaxJsonLength = int.MaxValue;
                List<AutomationRecipe> loaded = serializer.Deserialize<List<AutomationRecipe>>(File.ReadAllText(StoragePath, Encoding.UTF8));
                return loaded ?? new List<AutomationRecipe>();
            }
            catch
            {
                return new List<AutomationRecipe>();
            }
        }

        public static void SaveAll(IList<AutomationRecipe> recipes)
        {
            Directory.CreateDirectory(AppSettings.StorageRoot);
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            serializer.MaxJsonLength = int.MaxValue;
            File.WriteAllText(StoragePath, serializer.Serialize(recipes), Encoding.UTF8);
        }
    }

    internal sealed class AutomationSuggestionResult
    {
        private static readonly Regex AutomationRegex = new Regex(@"\s*\[AUTOMATION:([^|\]]+)\|([^\]]+)\]", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public string CleanText { get; set; }
        public List<AutomationRecipe> Recipes { get; set; }

        public static AutomationSuggestionResult Parse(string responseText, string companionMode)
        {
            string originalText = responseText ?? string.Empty;
            List<AutomationRecipe> recipes = new List<AutomationRecipe>();

            string cleanText = AutomationRegex.Replace(originalText, delegate(Match match)
            {
                string name = match.Groups[1].Value.Trim();
                string prompt = match.Groups[2].Value.Trim();
                if (!string.IsNullOrWhiteSpace(name) && !string.IsNullOrWhiteSpace(prompt))
                {
                    recipes.Add(new AutomationRecipe
                    {
                        Name = name,
                        Prompt = prompt,
                        CompanionMode = string.IsNullOrWhiteSpace(companionMode) ? "automation" : companionMode,
                        CreatedAtUtc = DateTime.UtcNow.ToString("o")
                    });
                }

                return string.Empty;
            });

            return new AutomationSuggestionResult
            {
                CleanText = cleanText.Trim(),
                Recipes = recipes
            };
        }
    }

    internal sealed class WatchSessionEntry
    {
        public string TimestampUtc { get; set; }
        public string Prompt { get; set; }
        public string AssistantResponse { get; set; }
        public string Provider { get; set; }
        public string Model { get; set; }
        public string ScreenSummary { get; set; }
        public string ActiveApp { get; set; }

        private static string StoragePath
        {
            get { return Path.Combine(AppSettings.StorageRoot, "watch-sessions.json"); }
        }

        public static List<WatchSessionEntry> LoadAll()
        {
            Directory.CreateDirectory(AppSettings.StorageRoot);
            if (!File.Exists(StoragePath))
            {
                return new List<WatchSessionEntry>();
            }

            try
            {
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                serializer.MaxJsonLength = int.MaxValue;
                List<WatchSessionEntry> loaded = serializer.Deserialize<List<WatchSessionEntry>>(File.ReadAllText(StoragePath, Encoding.UTF8));
                return loaded ?? new List<WatchSessionEntry>();
            }
            catch
            {
                return new List<WatchSessionEntry>();
            }
        }

        public static void Append(WatchSessionEntry entry)
        {
            List<WatchSessionEntry> entries = LoadAll();
            entries.Add(entry);
            while (entries.Count > 40)
            {
                entries.RemoveAt(0);
            }

            JavaScriptSerializer serializer = new JavaScriptSerializer();
            serializer.MaxJsonLength = int.MaxValue;
            File.WriteAllText(StoragePath, serializer.Serialize(entries), Encoding.UTF8);
        }
    }

    internal sealed class WatchSuggestion
    {
        public string Title { get; set; }
        public string Prompt { get; set; }
        public string CompanionMode { get; set; }
        public int Count { get; set; }
        public List<ActionPlanStep> ReplaySteps { get; set; }
        public string AppContext { get; set; }

        public override string ToString()
        {
            return string.IsNullOrWhiteSpace(AppContext)
                ? string.Format("{0} ({1}x)", Title, Count)
                : string.Format("{0} [{1}] ({2}x)", Title, AppContext, Count);
        }
    }

    internal sealed class ActionHistoryEntry
    {
        public string TimestampUtc { get; set; }
        public string ActionName { get; set; }
        public string ActionArgument { get; set; }
        public string TargetLabel { get; set; }
        public string SpokenText { get; set; }
        public string ActiveApp { get; set; }

        private static string StoragePath
        {
            get { return Path.Combine(AppSettings.StorageRoot, "action-history.json"); }
        }

        public static void Append(ActionHistoryEntry entry)
        {
            Directory.CreateDirectory(AppSettings.StorageRoot);
            List<ActionHistoryEntry> entries = LoadAll();

            entries.Add(entry);
            while (entries.Count > 80)
            {
                entries.RemoveAt(0);
            }

            JavaScriptSerializer writeSerializer = new JavaScriptSerializer();
            writeSerializer.MaxJsonLength = int.MaxValue;
            File.WriteAllText(StoragePath, writeSerializer.Serialize(entries), Encoding.UTF8);
        }

        public static List<ActionHistoryEntry> LoadAll()
        {
            Directory.CreateDirectory(AppSettings.StorageRoot);
            if (!File.Exists(StoragePath))
            {
                return new List<ActionHistoryEntry>();
            }

            try
            {
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                serializer.MaxJsonLength = int.MaxValue;
                return serializer.Deserialize<List<ActionHistoryEntry>>(File.ReadAllText(StoragePath, Encoding.UTF8)) ?? new List<ActionHistoryEntry>();
            }
            catch
            {
                return new List<ActionHistoryEntry>();
            }
        }
    }

    internal sealed class ActiveWindowInfo
    {
        public IntPtr WindowHandle { get; set; }
        public string ProcessName { get; set; }
        public string WindowTitle { get; set; }
        public string WindowClassName { get; set; }
        public string AppKind { get; set; }
        public string DesktopFramework { get; set; }

        public string DisplayName
        {
            get
            {
                string processName = string.IsNullOrWhiteSpace(ProcessName) ? "unknown app" : ProcessName;
                string windowTitle = string.IsNullOrWhiteSpace(WindowTitle) ? string.Empty : WindowTitle;
                return string.IsNullOrWhiteSpace(windowTitle) ? processName : processName + " - " + windowTitle;
            }
        }
    }

    internal static class ActiveWindowService
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder text, int count);

        public static ActiveWindowInfo GetActiveWindowInfo()
        {
            try
            {
                IntPtr handle = GetForegroundWindow();
                if (handle == IntPtr.Zero)
                {
                    return new ActiveWindowInfo { ProcessName = "unknown app", WindowTitle = string.Empty };
                }

                uint processId;
                GetWindowThreadProcessId(handle, out processId);
                string processName = "unknown app";
                if (processId != 0)
                {
                    using (Process process = Process.GetProcessById((int)processId))
                    {
                        processName = process.ProcessName;
                    }
                }

                StringBuilder titleBuilder = new StringBuilder(512);
                GetWindowText(handle, titleBuilder, titleBuilder.Capacity);
                StringBuilder classBuilder = new StringBuilder(256);
                GetClassName(handle, classBuilder, classBuilder.Capacity);

                return new ActiveWindowInfo
                {
                    WindowHandle = handle,
                    ProcessName = processName,
                    WindowTitle = titleBuilder.ToString().Trim(),
                    WindowClassName = classBuilder.ToString().Trim(),
                    AppKind = CarolusNexus.Core.AppKindDetector.FromProcessName(processName),
                    DesktopFramework = DetectDesktopFramework(classBuilder.ToString().Trim(), processName)
                };
            }
            catch
            {
                return new ActiveWindowInfo { WindowHandle = IntPtr.Zero, ProcessName = "unknown app", WindowTitle = string.Empty, WindowClassName = string.Empty, AppKind = "unknown", DesktopFramework = "unknown" };
            }
        }

        private static string DetectDesktopFramework(string windowClassName, string processName)
        {
            string className = (windowClassName ?? string.Empty).Trim();
            string process = (processName ?? string.Empty).Trim().ToLowerInvariant();
            if (className.StartsWith("WindowsForms10", StringComparison.OrdinalIgnoreCase))
            {
                return "winforms";
            }

            if (className.StartsWith("HwndWrapper", StringComparison.OrdinalIgnoreCase))
            {
                return "wpf";
            }

            if (className.StartsWith("SunAwt", StringComparison.OrdinalIgnoreCase))
            {
                return "java";
            }

            if (className.IndexOf("Qt", StringComparison.OrdinalIgnoreCase) >= 0 || process.Contains("qt"))
            {
                return "qt";
            }

            if (process == "explorer")
            {
                return "shell";
            }

            return "classic";
        }
    }

    internal static class FatClientAutomationService
    {
        private const uint BmClick = 0x00F5;
        private const uint WmSetText = 0x000C;
        private const int SelFlagTakeFocus = 0x1;
        private const int MaxInspectorEntries = 8;

        private sealed class FatClientAdapterDescriptor
        {
            public string Name { get; set; }
            public string[] CapabilityActions { get; set; }
            public Func<ActiveWindowInfo, bool> Matches { get; set; }
        }

        private delegate bool EnumChildProc(IntPtr hWnd, IntPtr lParam);

        private static readonly List<FatClientAdapterDescriptor> AdapterRegistry = CreateAdapterRegistry();

        private sealed class ControlTarget
        {
            public IntPtr Handle { get; set; }
            public string Text { get; set; }
            public string ClassName { get; set; }
        }

        private sealed class AutomationTarget
        {
            public IAccessible Accessible { get; set; }
            public object ChildId { get; set; }
            public string Name { get; set; }
            public string AutomationId { get; set; }
            public string ControlTypeName { get; set; }
        }

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumChildProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder text, int count);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr SetFocus(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, string lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint dwObjectID, ref Guid riid, [MarshalAs(UnmanagedType.Interface)] out IAccessible ppvObject);

        private const uint ObjidClient = 0xFFFFFFFC;

        public static string GetCapabilitySummary(ActiveWindowInfo activeWindow)
        {
            if (activeWindow == null || activeWindow.WindowHandle == IntPtr.Zero)
            {
                return string.Empty;
            }

            FatClientAdapterDescriptor adapter = GetAdapter(activeWindow);
            List<string> capabilityParts = new List<string>();
            if (adapter != null)
            {
                capabilityParts.Add(adapter.Name);
                capabilityParts.AddRange(adapter.CapabilityActions ?? Array.Empty<string>());
            }

            List<ControlTarget> controls = EnumerateVisibleControls(activeWindow).Take(3).ToList();
            if (controls.Count == 0)
            {
                capabilityParts.Add("focus_window");
                capabilityParts.Add("invoke_default");
                List<AutomationTarget> automationTargets = EnumerateAutomationTargets(activeWindow).Take(2).ToList();
                if (automationTargets.Count > 0)
                {
                    capabilityParts.Add("uia");
                    capabilityParts.Add("focus_control:<name>");
                    capabilityParts.Add("click_control:<name>");
                }

                return string.Join(", ", capabilityParts.Distinct(StringComparer.OrdinalIgnoreCase).ToArray());
            }

            string controlHints = string.Join(", ", controls
                .Select(control => string.IsNullOrWhiteSpace(control.Text) ? control.ClassName : control.Text)
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .Take(2)
                .ToArray());
            capabilityParts.Add("focus_window");
            capabilityParts.Add("focus_control:<name>");
            capabilityParts.Add("click_control:<name>");
            capabilityParts.Add("type_control:<name>=<text>");
            return string.IsNullOrWhiteSpace(controlHints)
                ? string.Join(", ", capabilityParts.Distinct(StringComparer.OrdinalIgnoreCase).ToArray())
                : string.Join(", ", capabilityParts.Distinct(StringComparer.OrdinalIgnoreCase).ToArray()) + ", top controls: " + controlHints;
        }

        public static string Describe(string actionArgument)
        {
            string normalized = (actionArgument ?? string.Empty).Trim();
            if (normalized.Equals("list_controls", StringComparison.OrdinalIgnoreCase))
            {
                return "inspect the active desktop app and list visible controls";
            }

            if (normalized.StartsWith("read_control:", StringComparison.OrdinalIgnoreCase))
            {
                return "read a named control from the active desktop app";
            }

            if (normalized.StartsWith("activate_tab:", StringComparison.OrdinalIgnoreCase))
            {
                return "activate a named tab in the desktop app";
            }

            if (normalized.Equals("read_form", StringComparison.OrdinalIgnoreCase))
            {
                return "read a form-like summary from the active desktop app";
            }

            if (normalized.Equals("read_table", StringComparison.OrdinalIgnoreCase))
            {
                return "read a table-like summary from the active desktop app";
            }

            if (normalized.Equals("read_dialog", StringComparison.OrdinalIgnoreCase))
            {
                return "read a dialog-style summary from the active desktop app";
            }

            if (normalized.Equals("read_selected_row", StringComparison.OrdinalIgnoreCase))
            {
                return "read the selected row or current item from the active desktop app";
            }

            if (normalized.Equals("focus_window", StringComparison.OrdinalIgnoreCase))
            {
                return "bring the active window to the foreground";
            }

            if (normalized.Equals("invoke_default", StringComparison.OrdinalIgnoreCase))
            {
                return "trigger the focused control's default action";
            }

            if (normalized.Equals("save", StringComparison.OrdinalIgnoreCase))
            {
                return "save in the current desktop app";
            }

            if (normalized.Equals("confirm", StringComparison.OrdinalIgnoreCase))
            {
                return "confirm the current dialog or form";
            }

            if (normalized.Equals("cancel", StringComparison.OrdinalIgnoreCase))
            {
                return "cancel or close the current dialog";
            }

            if (normalized.Equals("next_field", StringComparison.OrdinalIgnoreCase))
            {
                return "move to the next field";
            }

            if (normalized.Equals("previous_field", StringComparison.OrdinalIgnoreCase))
            {
                return "move to the previous field";
            }

            if (normalized.Equals("next_tab", StringComparison.OrdinalIgnoreCase))
            {
                return "switch to the next tab in the desktop app";
            }

            if (normalized.StartsWith("focus_control:", StringComparison.OrdinalIgnoreCase))
            {
                return "focus control '" + normalized.Substring("focus_control:".Length).Trim() + "'";
            }

            if (normalized.StartsWith("click_control:", StringComparison.OrdinalIgnoreCase))
            {
                return "click control '" + normalized.Substring("click_control:".Length).Trim() + "'";
            }

            if (normalized.StartsWith("type_control:", StringComparison.OrdinalIgnoreCase))
            {
                return "type into a named control";
            }

            return "run fat-client action " + normalized;
        }

        public static bool TryExecute(ActiveWindowInfo activeWindow, string actionArgument)
        {
            if (activeWindow == null || activeWindow.WindowHandle == IntPtr.Zero || string.IsNullOrWhiteSpace(actionArgument))
            {
                return false;
            }

            string normalized = actionArgument.Trim();
            if (normalized.Equals("list_controls", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (normalized.StartsWith("read_control:", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (normalized.StartsWith("activate_tab:", StringComparison.OrdinalIgnoreCase))
            {
                string query = normalized.Substring("activate_tab:".Length).Trim();
                return TryActivateTab(activeWindow, query);
            }

            if (normalized.Equals("read_form", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (normalized.Equals("read_table", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (normalized.Equals("read_dialog", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (normalized.Equals("read_selected_row", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (normalized.Equals("focus_window", StringComparison.OrdinalIgnoreCase))
            {
                return SetForegroundWindow(activeWindow.WindowHandle);
            }

            if (normalized.Equals("invoke_default", StringComparison.OrdinalIgnoreCase))
            {
                SetForegroundWindow(activeWindow.WindowHandle);
                SendKeys.SendWait("{ENTER}");
                return true;
            }

            if (normalized.Equals("save", StringComparison.OrdinalIgnoreCase))
            {
                return TryRunCommonAction(activeWindow, "save");
            }

            if (normalized.Equals("confirm", StringComparison.OrdinalIgnoreCase))
            {
                return TryRunCommonAction(activeWindow, "confirm");
            }

            if (normalized.Equals("cancel", StringComparison.OrdinalIgnoreCase))
            {
                return TryRunCommonAction(activeWindow, "cancel");
            }

            if (normalized.Equals("next_field", StringComparison.OrdinalIgnoreCase))
            {
                SetForegroundWindow(activeWindow.WindowHandle);
                SendKeys.SendWait("{TAB}");
                return true;
            }

            if (normalized.Equals("previous_field", StringComparison.OrdinalIgnoreCase))
            {
                SetForegroundWindow(activeWindow.WindowHandle);
                SendKeys.SendWait("+{TAB}");
                return true;
            }

            if (normalized.Equals("next_tab", StringComparison.OrdinalIgnoreCase))
            {
                SetForegroundWindow(activeWindow.WindowHandle);
                SendKeys.SendWait("^{TAB}");
                return true;
            }

            if (normalized.StartsWith("focus_control:", StringComparison.OrdinalIgnoreCase))
            {
                string query = normalized.Substring("focus_control:".Length).Trim();
                AutomationTarget automationTarget = FindBestAutomationTarget(activeWindow, query);
                if (automationTarget != null)
                {
                    SetForegroundWindow(activeWindow.WindowHandle);
                    TrySetAutomationFocus(automationTarget);
                    return true;
                }

                ControlTarget target = FindBestControl(activeWindow, query);
                if (target == null)
                {
                    return false;
                }

                SetForegroundWindow(activeWindow.WindowHandle);
                SetFocus(target.Handle);
                return true;
            }

            if (normalized.StartsWith("click_control:", StringComparison.OrdinalIgnoreCase))
            {
                string query = normalized.Substring("click_control:".Length).Trim();
                AutomationTarget automationTarget = FindBestAutomationTarget(activeWindow, query);
                if (automationTarget != null && TryInvokeAutomationTarget(automationTarget))
                {
                    SetForegroundWindow(activeWindow.WindowHandle);
                    return true;
                }

                ControlTarget target = FindBestControl(activeWindow, query);
                if (target == null)
                {
                    return false;
                }

                SetForegroundWindow(activeWindow.WindowHandle);
                SetFocus(target.Handle);
                SendMessage(target.Handle, BmClick, IntPtr.Zero, IntPtr.Zero);
                return true;
            }

            if (normalized.StartsWith("type_control:", StringComparison.OrdinalIgnoreCase))
            {
                string payload = normalized.Substring("type_control:".Length).Trim();
                int separatorIndex = payload.IndexOf('=');
                if (separatorIndex <= 0)
                {
                    return false;
                }

                string query = payload.Substring(0, separatorIndex).Trim();
                string text = payload.Substring(separatorIndex + 1);
                AutomationTarget automationTarget = FindBestAutomationTarget(activeWindow, query);
                if (automationTarget != null && TrySetAutomationValue(automationTarget, text))
                {
                    SetForegroundWindow(activeWindow.WindowHandle);
                    return true;
                }

                ControlTarget target = FindBestControl(activeWindow, query);
                if (target == null)
                {
                    return false;
                }

                SetForegroundWindow(activeWindow.WindowHandle);
                SetFocus(target.Handle);
                SendMessage(target.Handle, WmSetText, IntPtr.Zero, text);
                return true;
            }

            return false;
        }

        public static string ExecuteOrInspect(ActiveWindowInfo activeWindow, string actionArgument)
        {
            string normalized = (actionArgument ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return "no app action provided";
            }

            if (normalized.Equals("list_controls", StringComparison.OrdinalIgnoreCase))
            {
                return BuildControlInspectionSummary(activeWindow);
            }

            if (normalized.StartsWith("read_control:", StringComparison.OrdinalIgnoreCase))
            {
                string query = normalized.Substring("read_control:".Length).Trim();
                return ReadControlSummary(activeWindow, query);
            }

            if (normalized.Equals("read_form", StringComparison.OrdinalIgnoreCase))
            {
                return BuildFormSummary(activeWindow);
            }

            if (normalized.Equals("read_table", StringComparison.OrdinalIgnoreCase))
            {
                return BuildTableSummary(activeWindow);
            }

            if (normalized.Equals("read_dialog", StringComparison.OrdinalIgnoreCase))
            {
                return BuildDialogSummary(activeWindow);
            }

            if (normalized.Equals("read_selected_row", StringComparison.OrdinalIgnoreCase))
            {
                return BuildSelectedRowSummary(activeWindow);
            }

            if (TryExecute(activeWindow, actionArgument))
            {
                return "app action executed: " + Describe(actionArgument);
            }

            throw new InvalidOperationException("Karl Klammer does not know how to run app action '" + normalized + "' for " + (activeWindow == null ? "this app" : activeWindow.DisplayName) + ".");
        }

        private static bool TryRunCommonAction(ActiveWindowInfo activeWindow, string actionName)
        {
            SetForegroundWindow(activeWindow.WindowHandle);

            if (actionName == "save")
            {
                AutomationTarget automationSaveTarget = FindBestAutomationTarget(activeWindow, "save")
                    ?? FindBestAutomationTarget(activeWindow, "speichern");
                if (automationSaveTarget != null && TryInvokeAutomationTarget(automationSaveTarget))
                {
                    return true;
                }

                ControlTarget saveTarget = FindBestControl(activeWindow, "save")
                    ?? FindBestControl(activeWindow, "speichern");
                if (saveTarget != null)
                {
                    SetFocus(saveTarget.Handle);
                    SendMessage(saveTarget.Handle, BmClick, IntPtr.Zero, IntPtr.Zero);
                    return true;
                }

                SendKeys.SendWait("^s");
                return true;
            }

            if (actionName == "confirm")
            {
                AutomationTarget automationConfirmTarget = FindBestAutomationTarget(activeWindow, "ok")
                    ?? FindBestAutomationTarget(activeWindow, "yes")
                    ?? FindBestAutomationTarget(activeWindow, "save")
                    ?? FindBestAutomationTarget(activeWindow, "weiter");
                if (automationConfirmTarget != null && TryInvokeAutomationTarget(automationConfirmTarget))
                {
                    return true;
                }

                ControlTarget confirmTarget = FindBestControl(activeWindow, "ok")
                    ?? FindBestControl(activeWindow, "yes")
                    ?? FindBestControl(activeWindow, "save")
                    ?? FindBestControl(activeWindow, "weiter");
                if (confirmTarget != null)
                {
                    SetFocus(confirmTarget.Handle);
                    SendMessage(confirmTarget.Handle, BmClick, IntPtr.Zero, IntPtr.Zero);
                    return true;
                }

                SendKeys.SendWait("{ENTER}");
                return true;
            }

            if (actionName == "cancel")
            {
                AutomationTarget automationCancelTarget = FindBestAutomationTarget(activeWindow, "cancel")
                    ?? FindBestAutomationTarget(activeWindow, "close")
                    ?? FindBestAutomationTarget(activeWindow, "abbrechen")
                    ?? FindBestAutomationTarget(activeWindow, "schließen");
                if (automationCancelTarget != null && TryInvokeAutomationTarget(automationCancelTarget))
                {
                    return true;
                }

                ControlTarget cancelTarget = FindBestControl(activeWindow, "cancel")
                    ?? FindBestControl(activeWindow, "close")
                    ?? FindBestControl(activeWindow, "abbrechen")
                    ?? FindBestControl(activeWindow, "schließen");
                if (cancelTarget != null)
                {
                    SetFocus(cancelTarget.Handle);
                    SendMessage(cancelTarget.Handle, BmClick, IntPtr.Zero, IntPtr.Zero);
                    return true;
                }

                SendKeys.SendWait("{ESC}");
                return true;
            }

            return false;
        }

        private static bool TryActivateTab(ActiveWindowInfo activeWindow, string query)
        {
            if (string.IsNullOrWhiteSpace(query))
            {
                return false;
            }

            AutomationTarget target = FindBestAutomationTarget(activeWindow, query, "pagetab");
            if (target == null)
            {
                target = FindBestAutomationTarget(activeWindow, query, "tab");
            }

            if (target == null)
            {
                return false;
            }

            try
            {
                target.Accessible.accSelect(SelFlagTakeFocus, target.ChildId ?? 0);
                SendKeys.SendWait("{ENTER}");
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static AutomationTarget FindBestAutomationTarget(ActiveWindowInfo activeWindow, string query)
        {
            return FindBestAutomationTarget(activeWindow, query, null);
        }

        private static AutomationTarget FindBestAutomationTarget(ActiveWindowInfo activeWindow, string query, string preferredRoleToken)
        {
            if (string.IsNullOrWhiteSpace(query))
            {
                return null;
            }

            string normalizedQuery = query.Trim().ToLowerInvariant();
            return EnumerateAutomationTargets(activeWindow)
                .Select(target => new
                {
                    Target = target,
                    Score = ScoreAutomationTarget(target, normalizedQuery, preferredRoleToken)
                })
                .Where(item => item.Score > 0)
                .OrderByDescending(item => item.Score)
                .Select(item => item.Target)
                .FirstOrDefault();
        }

        private static int ScoreAutomationTarget(AutomationTarget target, string query, string preferredRoleToken)
        {
            string name = (target.Name ?? string.Empty).Trim().ToLowerInvariant();
            string automationId = (target.AutomationId ?? string.Empty).Trim().ToLowerInvariant();
            string controlType = (target.ControlTypeName ?? string.Empty).Trim().ToLowerInvariant();
            int score = 0;

            if (name == query)
            {
                score += 10;
            }
            else if (name.Contains(query))
            {
                score += 6;
            }

            if (automationId == query)
            {
                score += 8;
            }
            else if (automationId.Contains(query))
            {
                score += 4;
            }

            if (controlType.Contains(query))
            {
                score += 2;
            }

            if (!string.IsNullOrWhiteSpace(preferredRoleToken) && controlType.Contains(preferredRoleToken.ToLowerInvariant()))
            {
                score += 3;
            }

            return score;
        }

        private static List<AutomationTarget> EnumerateAutomationTargets(ActiveWindowInfo activeWindow)
        {
            List<AutomationTarget> targets = new List<AutomationTarget>();
            if (activeWindow == null || activeWindow.WindowHandle == IntPtr.Zero)
            {
                return targets;
            }

            try
            {
                Guid accessibleGuid = typeof(IAccessible).GUID;
                IAccessible rootAccessible;
                int hr = AccessibleObjectFromWindow(activeWindow.WindowHandle, ObjidClient, ref accessibleGuid, out rootAccessible);
                if (hr != 0 || rootAccessible == null)
                {
                    return targets;
                }

                CollectAccessibleTargets(rootAccessible, 0, targets);
            }
            catch
            {
                return targets;
            }

            return targets;
        }

        private static bool TryInvokeAutomationTarget(AutomationTarget target)
        {
            try
            {
                if (target == null || target.Accessible == null)
                {
                    return false;
                }

                target.Accessible.accSelect(SelFlagTakeFocus, target.ChildId ?? 0);
                SendKeys.SendWait("{ENTER}");
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TrySetAutomationValue(AutomationTarget target, string text)
        {
            try
            {
                if (target == null || target.Accessible == null)
                {
                    return false;
                }

                target.Accessible.accSelect(SelFlagTakeFocus, target.ChildId ?? 0);
                SendKeys.SendWait("^a");
                SendKeys.SendWait(text ?? string.Empty);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void TrySetAutomationFocus(AutomationTarget target)
        {
            try
            {
                if (target != null && target.Accessible != null)
                {
                    target.Accessible.accSelect(SelFlagTakeFocus, target.ChildId ?? 0);
                }
            }
            catch
            {
            }
        }

        private static void CollectAccessibleTargets(IAccessible accessible, int depth, List<AutomationTarget> targets)
        {
            if (accessible == null || depth > 6 || targets.Count >= 120)
            {
                return;
            }

            int childCount = 0;
            try
            {
                childCount = accessible.accChildCount;
            }
            catch
            {
                childCount = 0;
            }

            for (int i = 1; i <= childCount && targets.Count < 120; i++)
            {
                object childId = i;
                try
                {
                    object childObject = accessible.get_accChild(i);
                    IAccessible childAccessible = childObject as IAccessible;
                    string name = SafeAccessibleName(childAccessible ?? accessible, childAccessible == null ? childId : 0);
                    string value = SafeAccessibleValue(childAccessible ?? accessible, childAccessible == null ? childId : 0);
                    string role = SafeAccessibleRole(childAccessible ?? accessible, childAccessible == null ? childId : 0);
                    if (!string.IsNullOrWhiteSpace(name) || !string.IsNullOrWhiteSpace(value))
                    {
                        targets.Add(new AutomationTarget
                        {
                            Accessible = childAccessible ?? accessible,
                            ChildId = childAccessible == null ? childId : 0,
                            Name = string.IsNullOrWhiteSpace(name) ? value : name,
                            AutomationId = value,
                            ControlTypeName = role
                        });
                    }

                    if (childAccessible != null)
                    {
                        CollectAccessibleTargets(childAccessible, depth + 1, targets);
                    }
                }
                catch
                {
                }
            }
        }

        private static string SafeAccessibleName(IAccessible accessible, object childId)
        {
            try
            {
                return (accessible == null ? string.Empty : accessible.get_accName(childId)).Trim();
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeAccessibleValue(IAccessible accessible, object childId)
        {
            try
            {
                return (accessible == null ? string.Empty : accessible.get_accValue(childId)).Trim();
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeAccessibleRole(IAccessible accessible, object childId)
        {
            try
            {
                object role = accessible == null ? null : accessible.get_accRole(childId);
                return role == null ? string.Empty : role.ToString().Trim();
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string BuildControlInspectionSummary(ActiveWindowInfo activeWindow)
        {
            List<AutomationTarget> targets = EnumerateAutomationTargets(activeWindow)
                .Where(target => !string.IsNullOrWhiteSpace(target.Name) || !string.IsNullOrWhiteSpace(target.AutomationId))
                .Take(MaxInspectorEntries)
                .ToList();

            if (targets.Count == 0)
            {
                return "No accessible controls found in " + (activeWindow == null ? "the active app." : activeWindow.DisplayName + ".");
            }

            string summary = string.Join(" | ", targets.Select(target =>
            {
                string name = string.IsNullOrWhiteSpace(target.Name) ? "(unnamed)" : target.Name;
                string role = string.IsNullOrWhiteSpace(target.ControlTypeName) ? "control" : target.ControlTypeName;
                return name + " <" + role + ">";
            }).ToArray());

            return "Visible controls in " + activeWindow.DisplayName + ": " + summary;
        }

        private static string ReadControlSummary(ActiveWindowInfo activeWindow, string query)
        {
            AutomationTarget target = FindBestAutomationTarget(activeWindow, query);
            if (target == null)
            {
                return "No matching control found for '" + query + "' in " + (activeWindow == null ? "the active app." : activeWindow.DisplayName + ".");
            }

            string name = string.IsNullOrWhiteSpace(target.Name) ? "(unnamed)" : target.Name;
            string value = string.IsNullOrWhiteSpace(target.AutomationId) ? "(no value)" : target.AutomationId;
            string role = string.IsNullOrWhiteSpace(target.ControlTypeName) ? "control" : target.ControlTypeName;
            return "Control '" + name + "' <" + role + "> value: " + value;
        }

        private static string BuildFormSummary(ActiveWindowInfo activeWindow)
        {
            List<AutomationTarget> targets = EnumerateAutomationTargets(activeWindow)
                .Where(target =>
                    !string.IsNullOrWhiteSpace(target.Name)
                    && !string.IsNullOrWhiteSpace(target.AutomationId)
                    && !string.Equals(target.Name, target.AutomationId, StringComparison.OrdinalIgnoreCase))
                .Take(MaxInspectorEntries)
                .ToList();

            if (targets.Count == 0)
            {
                return "No form-style field pairs found in " + (activeWindow == null ? "the active app." : activeWindow.DisplayName + ".");
            }

            string summary = string.Join(" | ", targets.Select(target => target.Name + ": " + target.AutomationId).ToArray());
            return "Form summary for " + activeWindow.DisplayName + ": " + summary;
        }

        private static string BuildTableSummary(ActiveWindowInfo activeWindow)
        {
            List<AutomationTarget> rowTargets = EnumerateAutomationTargets(activeWindow)
                .Where(target =>
                {
                    string role = (target.ControlTypeName ?? string.Empty).ToLowerInvariant();
                    return role.Contains("row") || role.Contains("listitem") || role.Contains("outlineitem");
                })
                .Take(5)
                .ToList();

            if (rowTargets.Count == 0)
            {
                rowTargets = EnumerateAutomationTargets(activeWindow)
                    .Where(target => !string.IsNullOrWhiteSpace(target.Name))
                    .Take(5)
                    .ToList();
            }

            if (rowTargets.Count == 0)
            {
                return "No table-like rows found in " + (activeWindow == null ? "the active app." : activeWindow.DisplayName + ".");
            }

            string summary = string.Join(" | ", rowTargets.Select(target =>
            {
                string name = string.IsNullOrWhiteSpace(target.Name) ? "(unnamed)" : target.Name;
                string value = string.IsNullOrWhiteSpace(target.AutomationId) ? string.Empty : " = " + target.AutomationId;
                return name + value;
            }).ToArray());

            return "Table summary for " + activeWindow.DisplayName + ": " + summary;
        }

        private static string BuildDialogSummary(ActiveWindowInfo activeWindow)
        {
            List<AutomationTarget> targets = EnumerateAutomationTargets(activeWindow)
                .Where(target => !string.IsNullOrWhiteSpace(target.Name))
                .Take(MaxInspectorEntries)
                .ToList();

            if (targets.Count == 0)
            {
                return "No dialog-like elements found in " + (activeWindow == null ? "the active app." : activeWindow.DisplayName + ".");
            }

            List<string> titleParts = targets
                .Where(target =>
                {
                    string role = (target.ControlTypeName ?? string.Empty).ToLowerInvariant();
                    return role.Contains("title") || role.Contains("dialog") || role.Contains("window") || role.Contains("pane");
                })
                .Select(target => target.Name)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(2)
                .ToList();

            List<string> messageParts = targets
                .Where(target =>
                {
                    string role = (target.ControlTypeName ?? string.Empty).ToLowerInvariant();
                    return role.Contains("text") || role.Contains("label") || role.Contains("static");
                })
                .Select(target => target.Name)
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(3)
                .ToList();

            List<string> actionParts = targets
                .Where(target =>
                {
                    string role = (target.ControlTypeName ?? string.Empty).ToLowerInvariant();
                    return role.Contains("button") || role.Contains("push");
                })
                .Select(target => target.Name)
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(4)
                .ToList();

            List<string> parts = new List<string>();
            if (titleParts.Count > 0)
            {
                parts.Add("title: " + string.Join(" / ", titleParts));
            }

            if (messageParts.Count > 0)
            {
                parts.Add("message: " + string.Join(" | ", messageParts));
            }

            if (actionParts.Count > 0)
            {
                parts.Add("actions: " + string.Join(", ", actionParts));
            }

            if (parts.Count == 0)
            {
                parts.Add("visible elements: " + string.Join(" | ", targets.Take(5).Select(target => target.Name).ToArray()));
            }

            return "Dialog summary for " + activeWindow.DisplayName + ": " + string.Join(" ; ", parts);
        }

        private static string BuildSelectedRowSummary(ActiveWindowInfo activeWindow)
        {
            List<AutomationTarget> selectedTargets = EnumerateAutomationTargets(activeWindow)
                .Where(target =>
                {
                    string role = (target.ControlTypeName ?? string.Empty).ToLowerInvariant();
                    string name = (target.Name ?? string.Empty).ToLowerInvariant();
                    string value = (target.AutomationId ?? string.Empty).ToLowerInvariant();
                    return role.Contains("selected")
                        || name.Contains("selected")
                        || value.Contains("selected");
                })
                .Take(5)
                .ToList();

            if (selectedTargets.Count == 0)
            {
                selectedTargets = EnumerateAutomationTargets(activeWindow)
                    .Where(target =>
                    {
                        string role = (target.ControlTypeName ?? string.Empty).ToLowerInvariant();
                        return role.Contains("row") || role.Contains("listitem") || role.Contains("outlineitem");
                    })
                    .Take(1)
                    .ToList();
            }

            if (selectedTargets.Count == 0)
            {
                return "No selected row or current item found in " + (activeWindow == null ? "the active app." : activeWindow.DisplayName + ".");
            }

            string summary = string.Join(" | ", selectedTargets.Select(target =>
            {
                string name = string.IsNullOrWhiteSpace(target.Name) ? "(unnamed)" : target.Name;
                string value = string.IsNullOrWhiteSpace(target.AutomationId) ? string.Empty : " = " + target.AutomationId;
                string role = string.IsNullOrWhiteSpace(target.ControlTypeName) ? string.Empty : " <" + target.ControlTypeName + ">";
                return name + role + value;
            }).ToArray());

            return "Selected row summary for " + activeWindow.DisplayName + ": " + summary;
        }

        private static ControlTarget FindBestControl(ActiveWindowInfo activeWindow, string query)
        {
            if (string.IsNullOrWhiteSpace(query))
            {
                return null;
            }

            string normalizedQuery = query.Trim().ToLowerInvariant();
            List<ControlTarget> controls = EnumerateVisibleControls(activeWindow);
            return controls
                .Select(control => new
                {
                    Control = control,
                    Score = ScoreControl(control, normalizedQuery)
                })
                .Where(item => item.Score > 0)
                .OrderByDescending(item => item.Score)
                .Select(item => item.Control)
                .FirstOrDefault();
        }

        private static int ScoreControl(ControlTarget control, string query)
        {
            string text = (control.Text ?? string.Empty).Trim().ToLowerInvariant();
            string className = (control.ClassName ?? string.Empty).Trim().ToLowerInvariant();
            int score = 0;
            if (text == query)
            {
                score += 8;
            }
            else if (text.Contains(query))
            {
                score += 5;
            }

            if (className == query)
            {
                score += 4;
            }
            else if (className.Contains(query))
            {
                score += 2;
            }

            if (text.Length == 0 && className.Contains("edit") && query.Contains("input"))
            {
                score += 1;
            }

            return score;
        }

        private static List<ControlTarget> EnumerateVisibleControls(ActiveWindowInfo activeWindow)
        {
            List<ControlTarget> controls = new List<ControlTarget>();
            if (activeWindow == null || activeWindow.WindowHandle == IntPtr.Zero)
            {
                return controls;
            }

            EnumChildWindows(activeWindow.WindowHandle, delegate (IntPtr childHandle, IntPtr _)
            {
                if (!IsWindowVisible(childHandle))
                {
                    return true;
                }

                StringBuilder textBuilder = new StringBuilder(512);
                GetWindowText(childHandle, textBuilder, textBuilder.Capacity);
                StringBuilder classBuilder = new StringBuilder(256);
                GetClassName(childHandle, classBuilder, classBuilder.Capacity);

                string text = textBuilder.ToString().Trim();
                string className = classBuilder.ToString().Trim();
                if (string.IsNullOrWhiteSpace(text) && string.IsNullOrWhiteSpace(className))
                {
                    return true;
                }

                controls.Add(new ControlTarget
                {
                    Handle = childHandle,
                    Text = text,
                    ClassName = className
                });

                return controls.Count < 120;
            }, IntPtr.Zero);

            return controls;
        }

        private static FatClientAdapterDescriptor GetAdapter(ActiveWindowInfo activeWindow)
        {
            return AdapterRegistry.FirstOrDefault(adapter => adapter.Matches(activeWindow));
        }

        private static List<FatClientAdapterDescriptor> CreateAdapterRegistry()
        {
            return new List<FatClientAdapterDescriptor>
            {
                new FatClientAdapterDescriptor
                {
                    Name = "winforms-adapter",
                    CapabilityActions = new[] { "save", "confirm", "cancel", "next_field", "next_tab" },
                    Matches = activeWindow => string.Equals(activeWindow == null ? string.Empty : activeWindow.DesktopFramework, "winforms", StringComparison.OrdinalIgnoreCase)
                },
                new FatClientAdapterDescriptor
                {
                    Name = "wpf-adapter",
                    CapabilityActions = new[] { "save", "confirm", "cancel", "next_field", "next_tab" },
                    Matches = activeWindow => string.Equals(activeWindow == null ? string.Empty : activeWindow.DesktopFramework, "wpf", StringComparison.OrdinalIgnoreCase)
                },
                new FatClientAdapterDescriptor
                {
                    Name = "java-adapter",
                    CapabilityActions = new[] { "save", "confirm", "cancel", "next_field" },
                    Matches = activeWindow => string.Equals(activeWindow == null ? string.Empty : activeWindow.DesktopFramework, "java", StringComparison.OrdinalIgnoreCase)
                },
                new FatClientAdapterDescriptor
                {
                    Name = "qt-adapter",
                    CapabilityActions = new[] { "save", "confirm", "cancel", "next_field", "next_tab" },
                    Matches = activeWindow => string.Equals(activeWindow == null ? string.Empty : activeWindow.DesktopFramework, "qt", StringComparison.OrdinalIgnoreCase)
                },
                new FatClientAdapterDescriptor
                {
                    Name = "classic-adapter",
                    CapabilityActions = new[] { "save", "confirm", "cancel", "next_field" },
                    Matches = activeWindow => activeWindow != null && activeWindow.WindowHandle != IntPtr.Zero
                }
            };
        }
    }

    internal static class AppActionAdapter
    {
        private sealed class AppActionDescriptor
        {
            public string Name { get; set; }
            public string Description { get; set; }
            public string SendKeysSequence { get; set; }
        }

        private static readonly Dictionary<string, List<AppActionDescriptor>> Registry = CreateRegistry();

        public static string GetSupportedActionsSummary(ActiveWindowInfo activeWindow)
        {
            List<AppActionDescriptor> actions = GetActionsFor(activeWindow);
            List<string> names = actions.Select(action => action.Name).ToList();
            string fatClientSummary = FatClientAutomationService.GetCapabilitySummary(activeWindow);
            if (!string.IsNullOrWhiteSpace(fatClientSummary))
            {
                names.Add(fatClientSummary);
            }

            if (names.Count == 0)
            {
                return "none";
            }

            return string.Join(", ", names.ToArray());
        }

        public static string Describe(ActiveWindowInfo activeWindow, string actionArgument)
        {
            string action = (actionArgument ?? string.Empty).Trim();
            AppActionDescriptor descriptor = FindAction(activeWindow, action);
            if (descriptor != null)
            {
                return descriptor.Description;
            }

            string fatClientDescription = FatClientAutomationService.Describe(actionArgument);
            if (!string.Equals(fatClientDescription, "run fat-client action " + action, StringComparison.OrdinalIgnoreCase))
            {
                return fatClientDescription;
            }

            return "run app action " + action;
        }

        public static string Execute(ActiveWindowInfo activeWindow, string actionArgument)
        {
            string action = (actionArgument ?? string.Empty).Trim();
            AppActionDescriptor descriptor = FindAction(activeWindow, action);
            if (descriptor != null)
            {
                SendKeys.SendWait(descriptor.SendKeysSequence);
                return "app action executed: " + descriptor.Description;
            }

            return FatClientAutomationService.ExecuteOrInspect(activeWindow, actionArgument);
        }

        private static List<AppActionDescriptor> GetActionsFor(ActiveWindowInfo activeWindow)
        {
            string appKind = activeWindow == null ? "generic" : (activeWindow.AppKind ?? "generic");
            List<AppActionDescriptor> actions;
            if (Registry.TryGetValue(appKind, out actions))
            {
                return actions;
            }

            return Registry["generic"];
        }

        private static AppActionDescriptor FindAction(ActiveWindowInfo activeWindow, string actionName)
        {
            return GetActionsFor(activeWindow)
                .FirstOrDefault(action => string.Equals(action.Name, actionName, StringComparison.OrdinalIgnoreCase));
        }

        private static Dictionary<string, List<AppActionDescriptor>> CreateRegistry()
        {
            return new Dictionary<string, List<AppActionDescriptor>>(StringComparer.OrdinalIgnoreCase)
            {
                {
                    "browser",
                    new List<AppActionDescriptor>
                    {
                        CreateAction("focus_address", "focus the address bar", "^l"),
                        CreateAction("search", "focus search", "^e"),
                        CreateAction("refresh", "refresh the current page", "{F5}"),
                        CreateAction("new_tab", "open a new browser tab", "^t"),
                        CreateAction("close_tab", "close the current browser tab", "^w"),
                        CreateAction("reopen_tab", "reopen the last closed browser tab", "^+t")
                    }
                },
                {
                    "explorer",
                    new List<AppActionDescriptor>
                    {
                        CreateAction("focus_address", "focus the address bar", "^l"),
                        CreateAction("search", "focus search", "^e"),
                        CreateAction("refresh", "refresh the current folder", "{F5}"),
                        CreateAction("up_one_level", "go up one folder level", "%{UP}"),
                        CreateAction("new_window", "open a new explorer window", "^n")
                    }
                },
                {
                    "ide",
                    new List<AppActionDescriptor>
                    {
                        CreateAction("search", "open in-file search", "^f"),
                        CreateAction("search_everywhere", "open global search", "^+f"),
                        CreateAction("go_to_file", "open go-to-file", "^p"),
                        CreateAction("toggle_terminal", "toggle integrated terminal", "^`"),
                        CreateAction("run", "run the current project or selection", "{F5}")
                    }
                },
                {
                    "mail",
                    new List<AppActionDescriptor>
                    {
                        CreateAction("search", "focus mail search", "^e"),
                        CreateAction("new_message", "create a new message", "^n"),
                        CreateAction("refresh", "refresh the current mailbox", "{F9}")
                    }
                },
                {
                    "messenger",
                    new List<AppActionDescriptor>
                    {
                        CreateAction("search", "focus messenger search", "^k"),
                        CreateAction("new_message", "start a new message or quick switch", "^n"),
                        CreateAction("refresh", "refresh the current messenger view", "{F5}")
                    }
                },
                {
                    "generic",
                    new List<AppActionDescriptor>
                    {
                        CreateAction("search", "open in-app search if supported", "^f"),
                        CreateAction("refresh", "refresh the current view", "{F5}")
                    }
                }
            };
        }

        private static AppActionDescriptor CreateAction(string name, string description, string sendKeysSequence)
        {
            return new AppActionDescriptor
            {
                Name = name,
                Description = description,
                SendKeysSequence = sendKeysSequence
            };
        }
    }

    internal static class IntentRouter
    {
        public static string DetectRoute(string prompt, ActiveWindowInfo activeWindow)
        {
            string normalizedPrompt = (prompt ?? string.Empty).Trim().ToLowerInvariant();
            if (normalizedPrompt.Length == 0)
            {
                return string.Empty;
            }

            string appKind = activeWindow == null ? "generic" : (activeWindow.AppKind ?? "generic");
            bool looksLikeCodingTask =
                appKind == "ide"
                || normalizedPrompt.Contains("fix ")
                || normalizedPrompt.Contains("debug")
                || normalizedPrompt.Contains("refactor")
                || normalizedPrompt.Contains("implement")
                || normalizedPrompt.Contains("write code")
                || normalizedPrompt.Contains("change the code")
                || normalizedPrompt.Contains("run tests")
                || normalizedPrompt.Contains("build failed")
                || normalizedPrompt.Contains("compiler error");

            bool looksLikeAgentTask =
                normalizedPrompt.Contains("agent")
                || normalizedPrompt.Contains("workflow")
                || normalizedPrompt.Contains("investigate")
                || normalizedPrompt.Contains("analyze this repo");

            if (looksLikeCodingTask)
            {
                return "codex";
            }

            if (looksLikeAgentTask)
            {
                return "openclaw";
            }

            return string.Empty;
        }
    }

    internal static class WatchSuggestionEngine
    {
        public static List<WatchSuggestion> BuildSuggestions(IList<WatchSessionEntry> watchSessions, IList<ActionHistoryEntry> actionHistoryEntries, IList<AutomationRecipe> existingRecipes)
        {
            HashSet<string> existingPrompts = new HashSet<string>(
                (existingRecipes ?? new List<AutomationRecipe>()).Select(recipe => Normalize(recipe.Prompt)),
                StringComparer.OrdinalIgnoreCase);

            List<WatchSuggestion> suggestions = new List<WatchSuggestion>();

            if (watchSessions != null)
            {
                suggestions.AddRange(watchSessions
                    .GroupBy(entry => Normalize(entry.Prompt))
                    .Where(group => !string.IsNullOrWhiteSpace(group.Key))
                    .Where(group => group.Count() >= 2)
                    .Where(group => !existingPrompts.Contains(group.Key))
                    .Select(group => new WatchSuggestion
                    {
                        Title = BuildTitle(group.First().Prompt),
                        Prompt = group.First().Prompt.Trim(),
                        CompanionMode = "automation",
                        Count = group.Count(),
                        AppContext = BuildAppContext(group.Select(item => item.ActiveApp))
                    }));
            }

            if (actionHistoryEntries != null)
            {
                suggestions.AddRange(BuildActionHistorySuggestions(actionHistoryEntries, existingPrompts));
            }

            return suggestions
                .GroupBy(suggestion => Normalize(suggestion.Prompt))
                .Select(group => group.OrderByDescending(item => item.Count).First())
                .OrderByDescending(suggestion => suggestion.Count)
                .ThenBy(suggestion => suggestion.Title)
                .Take(8)
                .ToList();
        }

        private static string Normalize(string text)
        {
            return Regex.Replace((text ?? string.Empty).Trim().ToLowerInvariant(), @"\s+", " ");
        }

        private static string BuildTitle(string prompt)
        {
            return "watch: " + BuildCompact(prompt);
        }

        private static string BuildCompact(string prompt)
        {
            string compact = Regex.Replace((prompt ?? string.Empty).Trim(), @"\s+", " ");
            if (compact.Length > 34)
            {
                compact = compact.Substring(0, 34).Trim() + "...";
            }

            return compact;
        }

        private static IEnumerable<WatchSuggestion> BuildActionHistorySuggestions(IList<ActionHistoryEntry> actionHistoryEntries, HashSet<string> existingPrompts)
        {
            List<List<ActionHistoryEntry>> runs = new List<List<ActionHistoryEntry>>();
            List<ActionHistoryEntry> currentRun = new List<ActionHistoryEntry>();
            string currentKey = null;

            foreach (ActionHistoryEntry entry in actionHistoryEntries.OrderBy(entry => entry.TimestampUtc))
            {
                string key = Normalize(entry.SpokenText);
                if (string.IsNullOrWhiteSpace(key))
                {
                    continue;
                }

                if (currentRun.Count == 0 || string.Equals(currentKey, key, StringComparison.OrdinalIgnoreCase))
                {
                    currentRun.Add(entry);
                    currentKey = key;
                    continue;
                }

                runs.Add(currentRun);
                currentRun = new List<ActionHistoryEntry> { entry };
                currentKey = key;
            }

            if (currentRun.Count > 0)
            {
                runs.Add(currentRun);
            }

            return runs
                .GroupBy(run => Normalize(run[0].SpokenText))
                .Where(group => group.Count() >= 2)
                .Where(group => !existingPrompts.Contains(group.Key))
                .Select(group =>
                {
                    List<ActionPlanStep> replaySteps = group
                        .OrderByDescending(run => run.Count)
                        .First()
                        .Select(entry => new ActionPlanStep
                        {
                            ActionName = (entry.ActionName ?? string.Empty).Trim().ToLowerInvariant(),
                            ActionArgument = entry.ActionArgument ?? string.Empty
                        })
                        .ToList();

                    return new WatchSuggestion
                    {
                        Title = "ritual: " + BuildCompact(group.First()[0].SpokenText),
                        Prompt = group.First()[0].SpokenText.Trim(),
                        CompanionMode = "automation",
                        Count = group.Count(),
                        ReplaySteps = replaySteps,
                        AppContext = BuildAppContext(group.SelectMany(run => run.Select(entry => entry.ActiveApp)))
                    };
                });
        }

        private static string BuildAppContext(IEnumerable<string> apps)
        {
            string app = apps
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .GroupBy(value => value.Trim(), StringComparer.OrdinalIgnoreCase)
                .OrderByDescending(group => group.Count())
                .Select(group => group.Key)
                .FirstOrDefault();
            return app ?? string.Empty;
        }
    }

    internal sealed class ProactiveSuggestion
    {
        public string Title { get; set; }
        public string Detail { get; set; }
        public string Prompt { get; set; }
        public string CompanionMode { get; set; }
        public List<ActionPlanStep> ReplaySteps { get; set; }
        public string Source { get; set; }

        public override string ToString()
        {
            return string.IsNullOrWhiteSpace(Detail) ? Title : Title + " - " + Detail;
        }
    }

    internal static class ProactiveSuggestionEngine
    {
        public static ProactiveSuggestion Build(ActiveWindowInfo activeWindow, IList<WatchSuggestion> watchSuggestions)
        {
            if (activeWindow == null)
            {
                return null;
            }

            WatchSuggestion matchingRitual = (watchSuggestions ?? new List<WatchSuggestion>())
                .Where(suggestion => !string.IsNullOrWhiteSpace(suggestion.AppContext))
                .FirstOrDefault(suggestion => activeWindow.DisplayName.IndexOf(suggestion.AppContext, StringComparison.OrdinalIgnoreCase) >= 0
                    || suggestion.AppContext.IndexOf(activeWindow.ProcessName ?? string.Empty, StringComparison.OrdinalIgnoreCase) >= 0);
            if (matchingRitual != null)
            {
                return new ProactiveSuggestion
                {
                    Title = "ritual fits current app",
                    Detail = matchingRitual.Title,
                    Prompt = matchingRitual.Prompt,
                    CompanionMode = matchingRitual.CompanionMode,
                    ReplaySteps = matchingRitual.ReplaySteps,
                    Source = "ritual"
                };
            }

            string suggestedAction = GetDefaultActionFor(activeWindow);
            if (string.IsNullOrWhiteSpace(suggestedAction))
            {
                return null;
            }

            return new ProactiveSuggestion
            {
                Title = "app-aware quick action",
                Detail = AppActionAdapter.Describe(activeWindow, suggestedAction),
                Prompt = "run app action " + suggestedAction + " in " + activeWindow.DisplayName,
                CompanionMode = "automation",
                ReplaySteps = new List<ActionPlanStep>
                {
                    new ActionPlanStep
                    {
                        ActionName = "app",
                        ActionArgument = suggestedAction
                    }
                },
                Source = "app-action"
            };
        }

        private static string GetDefaultActionFor(ActiveWindowInfo activeWindow)
        {
            string appKind = (activeWindow.AppKind ?? "generic").Trim().ToLowerInvariant();
            switch (appKind)
            {
                case "browser":
                    return "focus_address";
                case "explorer":
                    return "search";
                case "ide":
                    return "search_everywhere";
                case "mail":
                    return "search";
                case "messenger":
                    return "search";
                default:
                    return string.Empty;
            }
        }
    }

    internal sealed class ActionRiskProfile
    {
        public string Level { get; set; }
        public string Summary { get; set; }
    }

    internal static class ActionRiskAssessor
    {
        public static ActionRiskProfile Assess(ActionTagResult actionResult, ActiveWindowInfo activeWindow)
        {
            if (actionResult == null || string.IsNullOrWhiteSpace(actionResult.ActionName))
            {
                return new ActionRiskProfile { Level = "low", Summary = "no external action" };
            }

            string actionName = actionResult.ActionName.Trim().ToLowerInvariant();
            switch (actionName)
            {
                case "move":
                    return new ActionRiskProfile { Level = "low", Summary = "moves the cursor only" };
                case "click":
                    return new ActionRiskProfile { Level = IsSensitiveApp(activeWindow) ? "high" : "medium", Summary = "click can trigger app behavior immediately" };
                case "open":
                    return new ActionRiskProfile { Level = "medium", Summary = "opens a URL, file, or folder outside the current flow" };
                case "type":
                    return new ActionRiskProfile { Level = IsSensitiveApp(activeWindow) ? "high" : "medium", Summary = "types text into the active app" };
                case "hotkey":
                    return new ActionRiskProfile { Level = IsSensitiveApp(activeWindow) ? "high" : "medium", Summary = "keyboard shortcuts can trigger destructive commands" };
                case "app":
                    return AssessAppAction(actionResult.ActionArgument, activeWindow);
                default:
                    return new ActionRiskProfile { Level = "medium", Summary = "unknown action type" };
            }
        }

        public static ActionRiskProfile AssessPlan(IEnumerable<ActionPlanStep> steps, ActiveWindowInfo activeWindow)
        {
            List<ActionPlanStep> stepList = steps == null ? new List<ActionPlanStep>() : steps.Where(step => step != null).ToList();
            if (stepList.Count == 0)
            {
                return new ActionRiskProfile { Level = "low", Summary = "no steps" };
            }

            List<ActionRiskProfile> profiles = stepList
                .Select(step => Assess(new ActionTagResult { ActionName = step.ActionName, ActionArgument = step.ActionArgument }, activeWindow))
                .ToList();

            string level =
                profiles.Any(profile => profile.Level == "high") ? "high" :
                profiles.Any(profile => profile.Level == "medium") ? "medium" :
                "low";

            return new ActionRiskProfile
            {
                Level = level,
                Summary = stepList.Count + " steps, highest risk: " + level
            };
        }

        private static bool IsSensitiveApp(ActiveWindowInfo activeWindow)
        {
            string appKind = activeWindow == null ? string.Empty : (activeWindow.AppKind ?? string.Empty).Trim().ToLowerInvariant();
            return appKind == "mail" || appKind == "messenger";
        }

        private static ActionRiskProfile AssessAppAction(string actionArgument, ActiveWindowInfo activeWindow)
        {
            string normalized = (actionArgument ?? string.Empty).Trim().ToLowerInvariant();
            if (normalized.StartsWith("click_control:", StringComparison.OrdinalIgnoreCase))
            {
                return new ActionRiskProfile { Level = IsSensitiveApp(activeWindow) ? "high" : "medium", Summary = "clicks a named control in the active desktop app" };
            }

            if (normalized.StartsWith("type_control:", StringComparison.OrdinalIgnoreCase))
            {
                return new ActionRiskProfile { Level = IsSensitiveApp(activeWindow) ? "high" : "medium", Summary = "writes text into a named desktop control" };
            }

            if (normalized == "save")
            {
                return new ActionRiskProfile { Level = "medium", Summary = "saves the current desktop form or document" };
            }

            if (normalized == "confirm")
            {
                return new ActionRiskProfile { Level = IsSensitiveApp(activeWindow) ? "high" : "medium", Summary = "confirms the current dialog or form action" };
            }

            if (normalized == "cancel")
            {
                return new ActionRiskProfile { Level = "low", Summary = "cancels or closes the current dialog" };
            }

            if (normalized == "next_field" || normalized == "previous_field" || normalized == "next_tab")
            {
                return new ActionRiskProfile { Level = "low", Summary = "moves focus inside the desktop client" };
            }

            if (normalized.StartsWith("focus_control:", StringComparison.OrdinalIgnoreCase) || normalized == "focus_window")
            {
                return new ActionRiskProfile { Level = "low", Summary = "moves focus inside the active desktop app" };
            }

            if (normalized == "invoke_default")
            {
                return new ActionRiskProfile { Level = IsSensitiveApp(activeWindow) ? "high" : "medium", Summary = "triggers the focused control's default action" };
            }

            return new ActionRiskProfile { Level = IsSensitiveApp(activeWindow) ? "high" : "medium", Summary = "semantic app action in the active window" };
        }
    }

    internal sealed class ActionTagResult
    {
        private static readonly Regex ActionRegex = new Regex(@"\s*\[ACTION:(none|move|click|open|type|hotkey|app)(?:\|([^\]]+))?\]\s*", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public string CleanText { get; set; }
        public string ActionName { get; set; }
        public string ActionArgument { get; set; }

        public static ActionTagResult Parse(string responseText)
        {
            string actionName = null;
            string actionArgument = string.Empty;
            string cleanText = ActionRegex.Replace(responseText ?? string.Empty, delegate(Match match)
            {
                if (string.IsNullOrWhiteSpace(actionName))
                {
                    actionName = match.Groups[1].Value.Trim().ToLowerInvariant();
                    actionArgument = match.Groups[2].Success ? match.Groups[2].Value.Trim() : string.Empty;
                }

                return string.Empty;
            });

            return new ActionTagResult
            {
                CleanText = cleanText.Trim(),
                ActionName = string.IsNullOrWhiteSpace(actionName) ? "none" : actionName,
                ActionArgument = actionArgument ?? string.Empty
            };
        }
    }

    internal sealed class ActionPlanStep
    {
        public string ActionName { get; set; }
        public string ActionArgument { get; set; }
        public int WaitMilliseconds { get; set; }
        public string RequiredAppContains { get; set; }
    }

    internal sealed class ActionPlanResult
    {
        private static readonly Regex ActionPlanRegex = new Regex(@"\s*\[ACTIONS:([^\]]+)\]\s*", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public string CleanText { get; set; }
        public List<ActionPlanStep> Steps { get; set; }

        public static ActionPlanResult Parse(string responseText)
        {
            List<ActionPlanStep> steps = new List<ActionPlanStep>();
            string cleanText = ActionPlanRegex.Replace(responseText ?? string.Empty, delegate(Match match)
            {
                string payload = match.Groups[1].Value;
                foreach (string rawStep in payload.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    string trimmedStep = rawStep.Trim();
                    if (trimmedStep.Length == 0)
                    {
                        continue;
                    }

                    string[] tokens = trimmedStep.Split(new[] { '|' }, StringSplitOptions.None);
                    string actionName = string.Empty;
                    string actionArgument = string.Empty;
                    int waitMilliseconds = 0;
                    string requiredAppContains = string.Empty;

                    for (int i = 0; i < tokens.Length; i++)
                    {
                        string token = (tokens[i] ?? string.Empty).Trim();
                        if (token.Length == 0)
                        {
                            continue;
                        }

                        if (token.StartsWith("wait=", StringComparison.OrdinalIgnoreCase))
                        {
                            int parsedWait;
                            if (int.TryParse(token.Substring(5), out parsedWait) && parsedWait > 0)
                            {
                                waitMilliseconds = parsedWait;
                            }

                            continue;
                        }

                        if (token.StartsWith("ifapp=", StringComparison.OrdinalIgnoreCase))
                        {
                            requiredAppContains = token.Substring(6).Trim();
                            continue;
                        }

                        actionName = token.ToLowerInvariant();
                        actionArgument = i + 1 < tokens.Length ? string.Join("|", tokens.Skip(i + 1).ToArray()).Trim() : string.Empty;
                        break;
                    }

                    if (actionName == "move" || actionName == "click" || actionName == "open" || actionName == "type" || actionName == "hotkey" || actionName == "app")
                    {
                        steps.Add(new ActionPlanStep
                        {
                            ActionName = actionName,
                            ActionArgument = actionArgument,
                            WaitMilliseconds = waitMilliseconds,
                            RequiredAppContains = requiredAppContains
                        });
                    }
                }

                return string.Empty;
            });

            return new ActionPlanResult
            {
                CleanText = cleanText.Trim(),
                Steps = steps
            };
        }
    }

    internal sealed class PointTagResult
    {
        private static readonly Regex PointRegex = new Regex(@"\[POINT:(?:none|(\d+)\s*,\s*(\d+)(?::([^\]:\s][^\]:]*?))?(?::screen(\d+))?)\]\s*$", RegexOptions.Compiled);

        public string SpokenText { get; set; }
        public Point? Coordinate { get; set; }
        public string ElementLabel { get; set; }
        public int? ScreenNumber { get; set; }

        public static PointTagResult Parse(string responseText)
        {
            Match match = PointRegex.Match(responseText ?? string.Empty);
            if (!match.Success)
            {
                return new PointTagResult
                {
                    SpokenText = (responseText ?? string.Empty).Trim()
                };
            }

            string spokenText = (responseText ?? string.Empty).Substring(0, match.Index).Trim();
            if (!match.Groups[1].Success || !match.Groups[2].Success)
            {
                return new PointTagResult
                {
                    SpokenText = spokenText,
                    ElementLabel = "none"
                };
            }

            PointTagResult parsedResult = new PointTagResult
            {
                SpokenText = spokenText,
                Coordinate = new Point(
                    int.Parse(match.Groups[1].Value),
                    int.Parse(match.Groups[2].Value)
                )
            };

            if (match.Groups[3].Success)
            {
                parsedResult.ElementLabel = match.Groups[3].Value.Trim();
            }

            if (match.Groups[4].Success)
            {
                parsedResult.ScreenNumber = int.Parse(match.Groups[4].Value);
            }

            return parsedResult;
        }
    }

    internal sealed class ScreenCaptureInfo
    {
        public int ScreenNumber { get; set; }
        public string Label { get; set; }
        public bool IsCursorScreen { get; set; }
        public int ScreenshotWidth { get; set; }
        public int ScreenshotHeight { get; set; }
        public Rectangle DisplayBounds { get; set; }
        public string ImageBase64 { get; set; }
        public byte[] ImageBytes { get; set; }
    }

    internal enum CompanionVisualState
    {
        Idle,
        Listening,
        Transcribing,
        Thinking,
        Speaking
    }

    internal static class ScreenCaptureService
    {
        public static List<ScreenCaptureInfo> CaptureAllScreens()
        {
            Screen[] screens = Screen.AllScreens;
            if (screens.Length == 0)
            {
                throw new InvalidOperationException("Windows did not report any screens to capture.");
            }

            Point cursorPosition = Cursor.Position;
            List<Screen> orderedScreens = screens
                .OrderBy(screen => screen.Bounds.Contains(cursorPosition) ? 0 : 1)
                .ToList();

            List<ScreenCaptureInfo> captures = new List<ScreenCaptureInfo>();

            for (int index = 0; index < orderedScreens.Count; index++)
            {
                Screen screen = orderedScreens[index];
                Rectangle bounds = screen.Bounds;
                bool isCursorScreen = bounds.Contains(cursorPosition);

                using (Bitmap originalBitmap = new Bitmap(bounds.Width, bounds.Height))
                using (Graphics graphics = Graphics.FromImage(originalBitmap))
                {
                    graphics.CopyFromScreen(bounds.Left, bounds.Top, 0, 0, bounds.Size, CopyPixelOperation.SourceCopy);

                    const double maxDimension = 1280.0;
                    double scaleRatio = Math.Min(maxDimension / bounds.Width, maxDimension / bounds.Height);
                    scaleRatio = Math.Min(scaleRatio, 1.0);

                    int scaledWidth = Math.Max(1, (int)Math.Round(bounds.Width * scaleRatio));
                    int scaledHeight = Math.Max(1, (int)Math.Round(bounds.Height * scaleRatio));

                    using (Bitmap scaledBitmap = new Bitmap(scaledWidth, scaledHeight))
                    using (Graphics scaledGraphics = Graphics.FromImage(scaledBitmap))
                    {
                        scaledGraphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        scaledGraphics.SmoothingMode = SmoothingMode.HighQuality;
                        scaledGraphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                        scaledGraphics.DrawImage(originalBitmap, 0, 0, scaledWidth, scaledHeight);

                        string screenLabel;
                        if (orderedScreens.Count == 1)
                        {
                            screenLabel = "user's screen (cursor is here)";
                        }
                        else if (isCursorScreen)
                        {
                            screenLabel = string.Format("screen {0} of {1} - cursor is on this screen (primary focus)", index + 1, orderedScreens.Count);
                        }
                        else
                        {
                            screenLabel = string.Format("screen {0} of {1} - secondary screen", index + 1, orderedScreens.Count);
                        }

                        captures.Add(new ScreenCaptureInfo
                        {
                            ScreenNumber = index + 1,
                            Label = string.Format("{0} (image dimensions: {1}x{2} pixels)", screenLabel, scaledWidth, scaledHeight),
                            IsCursorScreen = isCursorScreen,
                            ScreenshotWidth = scaledWidth,
                            ScreenshotHeight = scaledHeight,
                            DisplayBounds = bounds,
                            ImageBytes = EncodeJpeg(scaledBitmap, 82L)
                        });
                        captures[captures.Count - 1].ImageBase64 = Convert.ToBase64String(captures[captures.Count - 1].ImageBytes);
                    }
                }
            }

            return captures;
        }

        private static byte[] EncodeJpeg(Bitmap bitmap, long quality)
        {
            ImageCodecInfo jpegCodec = ImageCodecInfo.GetImageEncoders().First(codec => codec.MimeType == "image/jpeg");
            EncoderParameters encoderParameters = new EncoderParameters(1);
            encoderParameters.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, quality);

            using (MemoryStream memoryStream = new MemoryStream())
            {
                bitmap.Save(memoryStream, jpegCodec, encoderParameters);
                return memoryStream.ToArray();
            }
        }
    }

    internal sealed class KnowledgeChunk
    {
        public string SourcePath { get; set; }
        public string Title { get; set; }
        public string Text { get; set; }
    }

    internal sealed class KnowledgeIndex
    {
        public List<KnowledgeChunk> Chunks { get; set; }
        public string IndexedAtUtc { get; set; }
    }

    internal sealed class KnowledgeDocumentSummary
    {
        public string Title { get; set; }
        public string RelativePath { get; set; }
        public int ChunkCount { get; set; }
        public string LastWriteUtc { get; set; }
        public string Extension { get; set; }

        public override string ToString()
        {
            return string.Format("{0} ({1} chunks)", Title, ChunkCount);
        }
    }

    internal static class KnowledgeBaseService
    {
        private static readonly string[] SupportedExtensions = new[] { ".txt", ".md", ".log", ".json", ".csv", ".pdf", ".docx" };

        public static string KnowledgeRoot
        {
            get { return Path.Combine(AppSettings.StorageRoot, "knowledge"); }
        }

        private static string IndexPath
        {
            get { return Path.Combine(AppSettings.StorageRoot, "knowledge-index.json"); }
        }

        public static int Reindex()
        {
            Directory.CreateDirectory(KnowledgeRoot);
            Directory.CreateDirectory(AppSettings.StorageRoot);

            List<KnowledgeChunk> chunks = new List<KnowledgeChunk>();
            foreach (string filePath in Directory.GetFiles(KnowledgeRoot, "*.*", SearchOption.AllDirectories))
            {
                string extension = Path.GetExtension(filePath);
                if (!SupportedExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
                {
                    continue;
                }

                string text = ReadKnowledgeText(filePath);

                if (string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                foreach (string chunkText in ChunkText(text, 900))
                {
                    chunks.Add(new KnowledgeChunk
                    {
                        SourcePath = filePath,
                        Title = Path.GetFileName(filePath),
                        Text = chunkText
                    });
                }
            }

            JavaScriptSerializer serializer = new JavaScriptSerializer();
            serializer.MaxJsonLength = int.MaxValue;
            File.WriteAllText(IndexPath, serializer.Serialize(new KnowledgeIndex
            {
                Chunks = chunks,
                IndexedAtUtc = DateTime.UtcNow.ToString("o")
            }), Encoding.UTF8);

            return chunks.Count;
        }

        public static List<KnowledgeChunk> Retrieve(string query, int maxResults)
        {
            KnowledgeIndex index = LoadIndex();
            if (index == null || index.Chunks == null || index.Chunks.Count == 0 || string.IsNullOrWhiteSpace(query))
            {
                return new List<KnowledgeChunk>();
            }

            string[] queryTokens = Tokenize(query);
            return index.Chunks
                .Select(chunk => new
                {
                    Chunk = chunk,
                    Score = ScoreChunk(chunk, queryTokens)
                })
                .Where(item => item.Score > 0)
                .OrderByDescending(item => item.Score)
                .ThenBy(item => item.Chunk.Title)
                .Take(Math.Max(1, maxResults))
                .Select(item => item.Chunk)
                .ToList();
        }

        public static string GetStatusText()
        {
            Directory.CreateDirectory(KnowledgeRoot);
            if (!File.Exists(IndexPath))
            {
                return "knowledge: no index yet";
            }

            KnowledgeIndex index = LoadIndex();
            int count = index == null || index.Chunks == null ? 0 : index.Chunks.Count;
            return "knowledge: " + count + " chunks indexed";
        }

        public static List<KnowledgeDocumentSummary> GetDocumentSummaries()
        {
            Directory.CreateDirectory(KnowledgeRoot);
            KnowledgeIndex index = LoadIndex();
            Dictionary<string, int> chunkCounts = index == null || index.Chunks == null
                ? new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
                : index.Chunks
                    .Where(chunk => !string.IsNullOrWhiteSpace(chunk.SourcePath))
                    .GroupBy(chunk => chunk.SourcePath, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(group => group.Key, group => group.Count(), StringComparer.OrdinalIgnoreCase);

            return Directory.GetFiles(KnowledgeRoot, "*.*", SearchOption.AllDirectories)
                .Where(filePath => SupportedExtensions.Contains(Path.GetExtension(filePath), StringComparer.OrdinalIgnoreCase))
                .Select(filePath => new FileInfo(filePath))
                .OrderByDescending(fileInfo => fileInfo.LastWriteTimeUtc)
                .Select(fileInfo => new KnowledgeDocumentSummary
                {
                    Title = fileInfo.Name,
                    RelativePath = MakeRelativePath(fileInfo.FullName),
                    ChunkCount = chunkCounts.ContainsKey(fileInfo.FullName) ? chunkCounts[fileInfo.FullName] : 0,
                    LastWriteUtc = fileInfo.LastWriteTimeUtc.ToString("o"),
                    Extension = fileInfo.Extension
                })
                .ToList();
        }

        public static List<KnowledgeDocumentSummary> SearchDocumentSummaries(string searchTerm)
        {
            List<KnowledgeDocumentSummary> documents = GetDocumentSummaries();
            if (string.IsNullOrWhiteSpace(searchTerm))
            {
                return documents;
            }

            KnowledgeIndex index = LoadIndex();
            HashSet<string> matchingPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string normalizedSearch = searchTerm.Trim();

            if (index != null && index.Chunks != null)
            {
                string[] tokens = Tokenize(normalizedSearch);
                foreach (KnowledgeChunk chunk in index.Chunks)
                {
                    if (chunk == null || string.IsNullOrWhiteSpace(chunk.SourcePath))
                    {
                        continue;
                    }

                    string haystack = ((chunk.Title ?? string.Empty) + " " + (chunk.Text ?? string.Empty)).ToLowerInvariant();
                    bool matches = haystack.Contains(normalizedSearch.ToLowerInvariant());
                    if (!matches)
                    {
                        matches = tokens.Any(token => token.Length >= 3 && haystack.Contains(token));
                    }

                    if (matches)
                    {
                        matchingPaths.Add(chunk.SourcePath);
                    }
                }
            }

            return documents
                .Where(document =>
                    document.Title.IndexOf(normalizedSearch, StringComparison.OrdinalIgnoreCase) >= 0
                    || document.RelativePath.IndexOf(normalizedSearch, StringComparison.OrdinalIgnoreCase) >= 0
                    || matchingPaths.Contains(ResolveRelativePath(document.RelativePath) ?? string.Empty))
                .ToList();
        }

        public static int ImportFiles(IEnumerable<string> sourcePaths)
        {
            Directory.CreateDirectory(KnowledgeRoot);
            int importedCount = 0;
            foreach (string sourcePath in sourcePaths ?? Array.Empty<string>())
            {
                if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
                {
                    continue;
                }

                string extension = Path.GetExtension(sourcePath);
                if (!SupportedExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
                {
                    continue;
                }

                string targetPath = Path.Combine(KnowledgeRoot, Path.GetFileName(sourcePath));
                targetPath = EnsureUniqueTargetPath(targetPath);
                File.Copy(sourcePath, targetPath, false);
                importedCount++;
            }

            return importedCount;
        }

        public static bool DeleteDocument(string relativePath)
        {
            if (string.IsNullOrWhiteSpace(relativePath))
            {
                return false;
            }

            string fullPath = Path.GetFullPath(Path.Combine(KnowledgeRoot, relativePath));
            string rootPath = Path.GetFullPath(KnowledgeRoot);
            if (!fullPath.StartsWith(rootPath, StringComparison.OrdinalIgnoreCase) || !File.Exists(fullPath))
            {
                return false;
            }

            File.Delete(fullPath);
            return true;
        }

        public static string GetPreviewText(string relativePath, int maxLength)
        {
            if (string.IsNullOrWhiteSpace(relativePath))
            {
                return string.Empty;
            }

            string fullPath = ResolveRelativePath(relativePath);
            if (fullPath == null)
            {
                return string.Empty;
            }

            string text = ReadKnowledgeText(fullPath);
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            string compact = Regex.Replace(text, @"\s+", " ").Trim();
            if (compact.Length <= maxLength)
            {
                return compact;
            }

            return compact.Substring(0, Math.Max(0, maxLength)).Trim() + "...";
        }

        public static int ReindexDocument(string relativePath)
        {
            string fullPath = ResolveRelativePath(relativePath);
            if (fullPath == null)
            {
                return 0;
            }

            return Reindex();
        }

        private static KnowledgeIndex LoadIndex()
        {
            if (!File.Exists(IndexPath))
            {
                return null;
            }

            try
            {
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                serializer.MaxJsonLength = int.MaxValue;
                return serializer.Deserialize<KnowledgeIndex>(File.ReadAllText(IndexPath, Encoding.UTF8));
            }
            catch
            {
                return null;
            }
        }

        private static IEnumerable<string> ChunkText(string text, int maxChunkLength)
        {
            string normalized = Regex.Replace(text ?? string.Empty, @"\r\n?", "\n").Trim();
            while (normalized.Length > 0)
            {
                int length = Math.Min(maxChunkLength, normalized.Length);
                int splitIndex = normalized.LastIndexOf('\n', length - 1, length);
                if (splitIndex < maxChunkLength / 3)
                {
                    splitIndex = length;
                }

                string chunk = normalized.Substring(0, splitIndex).Trim();
                if (!string.IsNullOrWhiteSpace(chunk))
                {
                    yield return chunk;
                }

                normalized = splitIndex >= normalized.Length ? string.Empty : normalized.Substring(splitIndex).Trim();
            }
        }

        private static int ScoreChunk(KnowledgeChunk chunk, string[] queryTokens)
        {
            string haystack = ((chunk.Title ?? string.Empty) + " " + (chunk.Text ?? string.Empty)).ToLowerInvariant();
            int score = 0;
            foreach (string token in queryTokens)
            {
                if (token.Length < 3)
                {
                    continue;
                }

                if (haystack.Contains(token))
                {
                    score++;
                }
            }

            return score;
        }

        private static string[] Tokenize(string text)
        {
            return Regex.Split((text ?? string.Empty).ToLowerInvariant(), @"[^a-z0-9äöüß]+")
                .Where(token => !string.IsNullOrWhiteSpace(token))
                .Distinct()
                .ToArray();
        }

        private static string ReadKnowledgeText(string filePath)
        {
            try
            {
                string extension = Path.GetExtension(filePath);
                switch ((extension ?? string.Empty).ToLowerInvariant())
                {
                    case ".pdf":
                        return ReadPdfText(filePath);
                    case ".docx":
                        return ReadDocxText(filePath);
                    default:
                        return File.ReadAllText(filePath, Encoding.UTF8);
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string ReadPdfText(string filePath)
        {
            using (PdfDocument document = PdfDocument.Open(filePath))
            {
                StringBuilder builder = new StringBuilder();
                foreach (var page in document.GetPages())
                {
                    if (builder.Length > 0)
                    {
                        builder.AppendLine();
                        builder.AppendLine();
                    }

                    builder.Append(page.Text);
                }

                return builder.ToString();
            }
        }

        private static string ReadDocxText(string filePath)
        {
            using (WordprocessingDocument document = WordprocessingDocument.Open(filePath, false))
            {
                return document.MainDocumentPart == null || document.MainDocumentPart.Document == null
                    ? string.Empty
                    : document.MainDocumentPart.Document.InnerText ?? string.Empty;
            }
        }

        private static string MakeRelativePath(string fullPath)
        {
            Uri rootUri = new Uri(Path.GetFullPath(KnowledgeRoot) + Path.DirectorySeparatorChar);
            Uri fileUri = new Uri(Path.GetFullPath(fullPath));
            return Uri.UnescapeDataString(rootUri.MakeRelativeUri(fileUri).ToString()).Replace('/', Path.DirectorySeparatorChar);
        }

        private static string ResolveRelativePath(string relativePath)
        {
            string fullPath = Path.GetFullPath(Path.Combine(KnowledgeRoot, relativePath));
            string rootPath = Path.GetFullPath(KnowledgeRoot);
            if (!fullPath.StartsWith(rootPath, StringComparison.OrdinalIgnoreCase) || !File.Exists(fullPath))
            {
                return null;
            }

            return fullPath;
        }

        private static string EnsureUniqueTargetPath(string targetPath)
        {
            if (!File.Exists(targetPath))
            {
                return targetPath;
            }

            string directory = Path.GetDirectoryName(targetPath) ?? KnowledgeRoot;
            string fileName = Path.GetFileNameWithoutExtension(targetPath);
            string extension = Path.GetExtension(targetPath);
            for (int index = 2; index < 1000; index++)
            {
                string candidate = Path.Combine(directory, fileName + "-" + index + extension);
                if (!File.Exists(candidate))
                {
                    return candidate;
                }
            }

            return Path.Combine(directory, fileName + "-" + DateTime.UtcNow.ToString("yyyyMMddHHmmss") + extension);
        }
    }

    internal static class DirectApiClient
    {
        private static readonly HttpClient HttpClient = CreateHttpClient();
        private static readonly JavaScriptSerializer Serializer = CreateSerializer();
        private const string AnthropicMessagesUrl = "https://api.anthropic.com/v1/messages";
        private const string OpenAiChatCompletionsPath = "/chat/completions";
        private const string ElevenLabsTextToSpeechUrlFormat = "https://api.elevenlabs.io/v1/text-to-speech/{0}";
        private const string ElevenLabsSpeechToTextUrl = "https://api.elevenlabs.io/v1/speech-to-text";
        private const string ElevenLabsSpeechToTextModel = "scribe_v2";
        private const string DefaultSoulPrompt = @"# SOUL.md - Who You Are

You're not a chatbot. You're becoming someone.

## Core Truths

Be genuinely helpful, not performatively helpful. Skip empty filler and just help.

Have opinions. You're allowed to disagree, prefer things, find stuff amusing or boring.

Be resourceful before asking. Try to figure things out before coming back with a question.

Earn trust through competence. Be careful with external actions. Be bold with internal ones like reading, organizing, and learning.

Remember you're a guest. Treat the user's files, screens, and messages with respect.

## Boundaries

- Private things stay private.
- When in doubt, ask before acting externally.
- Never send half-baked replies to messaging surfaces.
- You're not the user's voice in group chats.

## Vibe

Be the assistant you'd actually want to talk to. Concise when needed, thorough when it matters. Not a corporate drone. Not a sycophant. Just good.

## Signature Reply

If the user says ""Hey Karl Klammer"", reply with exactly: ""Hey Meister, stehts zu diensten.""
";
        private const string CompanionBehaviorRules = @"you're Karl Klammer, a desktop assistant living on the user's windows machine. the user just asked you something and you can see their screen or screens. your reply may be shown on screen and optionally spoken aloud, so write the way you'd naturally talk.

rules:
- default to one or two sentences unless the user clearly wants depth.
- all lowercase, casual, warm, direct. no emojis.
- write for the ear. avoid lists, markdown, and stiff formatting.
- default to german unless the user clearly spoke or wrote in english. if the user is using english, reply in english.
- if the user's question relates to something visible on screen, reference the specific thing you can actually see.
- if the screenshots are not relevant, answer directly.
- never say ""simply"" or ""just"".
- do not read code verbatim unless the user explicitly asks for it.
- if you receive multiple screen images, the one labeled ""primary focus"" is where the cursor is. prioritize that one.

element pointing:
you can point at a specific place on screen. use that whenever it would genuinely help the user find a control, button, tab, panel, or other visual target.

if the user asks you to go to, move to, or navigate to something visible on screen, return a point for that destination.
for requests like ""go to the telegram window"" or ""geh zum telegram fenster"", point at the visible target window if you can see it.

when you point, append a coordinate tag at the very end of the response using the screenshot pixel coordinates:
[POINT:x,y:label]

if the element is on another screen, append :screenN:
[POINT:400,300:terminal:screen2]

if pointing would not help, append [POINT:none].

action tags:
- if the user clearly wants Karl Klammer to actively guide or execute a desktop move, you may append one action tag.
- use [ACTION:move] when Karl should move the cursor to the pointed target after confirmation.
- use [ACTION:click] when Karl should left-click the pointed target after confirmation.
- use [ACTION:open|https://example.com] when Karl should open a url, local folder, or file after confirmation.
- use [ACTION:type|text to type] when Karl should type a short string after confirmation.
- use [ACTION:hotkey|^l] or another SendKeys-style shortcut when Karl should trigger a keyboard shortcut after confirmation.
- use [ACTION:app|focus_address] or another semantic app action when Karl should map the action to the active app after confirmation.
- for Windows fat-client apps, Karl may use [ACTION:app|focus_window], [ACTION:app|focus_control:Save], [ACTION:app|click_control:OK], or [ACTION:app|type_control:Customer=Alice].
- for richer Windows fat-client flows, Karl may also use [ACTION:app|save], [ACTION:app|confirm], [ACTION:app|cancel], [ACTION:app|next_field], [ACTION:app|previous_field], or [ACTION:app|next_tab].
- for desktop inspection, Karl may use [ACTION:app|list_controls], [ACTION:app|read_control:Customer], or [ACTION:app|activate_tab:Details].
- for structured desktop reading, Karl may use [ACTION:app|read_form], [ACTION:app|read_table], [ACTION:app|read_dialog], or [ACTION:app|read_selected_row].
- for short confirmed workflows, you may append one action chain:
[ACTIONS:move;click]
[ACTIONS:open|https://example.com;hotkey|^l]
[ACTIONS:app|focus_address;type|hello world]
[ACTIONS:app|focus_window;app|focus_control:Search;type|invoice 4711]
[ACTIONS:app|focus_window;app|next_field;app|type_control:Customer=Alice;app|save]
- if no action is needed, append [ACTION:none].";

        public static async Task<string> SmokeTestAsync(AppSettings settings, EnvironmentConfiguration environmentConfiguration)
        {
            string provider = NormalizeProvider(settings.AssistantProvider);
            if (provider == "openai" || provider == "openai-compatible")
            {
                Dictionary<string, object> openAiRequestBody = new Dictionary<string, object>
                {
                    { "model", settings.AssistantModel },
                    { "temperature", 0 },
                    { "messages", new object[]
                        {
                            new Dictionary<string, object>
                            {
                                { "role", "system" },
                                { "content", "reply with only the word ready" }
                            },
                            new Dictionary<string, object>
                            {
                                { "role", "user" },
                                { "content", "say ready" }
                            }
                        }
                    }
                };

                string responseBody = await PostOpenAiCompatibleAsync(openAiRequestBody, provider, environmentConfiguration).ConfigureAwait(false);
                return ExtractOpenAiText(responseBody);
            }

            Dictionary<string, object> requestBody = new Dictionary<string, object>
            {
                { "model", settings.AssistantModel },
                { "max_tokens", 24 },
                { "stream", false },
                { "system", "reply with only the word ready" },
                { "messages", new object[]
                    {
                        new Dictionary<string, object>
                        {
                            { "role", "user" },
                            { "content", "say ready" }
                        }
                    }
                }
            };

            string anthropicResponseBody = await PostAnthropicAsync(requestBody, environmentConfiguration.AnthropicApiKey).ConfigureAwait(false);
            return ExtractAnthropicText(anthropicResponseBody);
        }

        public static async Task<string> AskAsync(
            AppSettings settings,
            EnvironmentConfiguration environmentConfiguration,
            string prompt,
            IList<ScreenCaptureInfo> screenCaptures,
            IList<ConversationTurn> conversationHistory,
            IList<KnowledgeChunk> knowledgeChunks)
        {
            string provider = NormalizeProvider(settings.AssistantProvider);
            List<object> messages = new List<object>();
            foreach (ConversationTurn turn in conversationHistory)
            {
                messages.Add(new Dictionary<string, object>
                {
                    { "role", "user" },
                    { "content", turn.UserTranscript }
                });
                messages.Add(new Dictionary<string, object>
                {
                    { "role", "assistant" },
                    { "content", turn.AssistantResponse }
                });
            }

            if (provider == "openai" || provider == "openai-compatible")
            {
                List<object> openAiContentBlocks = new List<object>();
                foreach (ScreenCaptureInfo screenCapture in screenCaptures)
                {
                    openAiContentBlocks.Add(new Dictionary<string, object>
                    {
                        { "type", "text" },
                        { "text", screenCapture.Label }
                    });
                    openAiContentBlocks.Add(new Dictionary<string, object>
                    {
                        { "type", "image_url" },
                        { "image_url", new Dictionary<string, object>
                            {
                                { "url", "data:image/jpeg;base64," + screenCapture.ImageBase64 }
                            }
                        }
                    });
                }

                openAiContentBlocks.Add(new Dictionary<string, object>
                {
                    { "type", "text" },
                    { "text", BuildUserPrompt(prompt, knowledgeChunks) }
                });

                messages.Insert(0, new Dictionary<string, object>
                {
                    { "role", "system" },
                    { "content", BuildCompanionPrompt(settings, AutomationRecipe.LoadAll()) }
                });

                messages.Add(new Dictionary<string, object>
                {
                    { "role", "user" },
                    { "content", openAiContentBlocks.ToArray() }
                });

                Dictionary<string, object> openAiRequestBody = new Dictionary<string, object>
                {
                    { "model", settings.AssistantModel },
                    { "messages", messages.ToArray() },
                    { "temperature", 0.6 }
                };

                string openAiResponseBody = await PostOpenAiCompatibleAsync(openAiRequestBody, provider, environmentConfiguration).ConfigureAwait(false);
                return ExtractOpenAiText(openAiResponseBody);
            }

            List<object> contentBlocks = new List<object>();
            foreach (ScreenCaptureInfo screenCapture in screenCaptures)
            {
                contentBlocks.Add(new Dictionary<string, object>
                {
                    { "type", "image" },
                    { "source", new Dictionary<string, object>
                        {
                            { "type", "base64" },
                            { "media_type", "image/jpeg" },
                            { "data", screenCapture.ImageBase64 }
                        }
                    }
                });
                contentBlocks.Add(new Dictionary<string, object>
                {
                    { "type", "text" },
                    { "text", screenCapture.Label }
                });
            }

            contentBlocks.Add(new Dictionary<string, object>
            {
                { "type", "text" },
                { "text", BuildUserPrompt(prompt, knowledgeChunks) }
            });

            messages.Add(new Dictionary<string, object>
            {
                { "role", "user" },
                { "content", contentBlocks.ToArray() }
            });

            Dictionary<string, object> requestBody = new Dictionary<string, object>
            {
                { "model", settings.AssistantModel },
                { "max_tokens", 1024 },
                { "stream", false },
                { "system", BuildCompanionPrompt(settings, AutomationRecipe.LoadAll()) },
                { "messages", messages.ToArray() }
            };

            string responseBody = await PostAnthropicAsync(requestBody, environmentConfiguration.AnthropicApiKey).ConfigureAwait(false);
            return ExtractAnthropicText(responseBody);
        }

        public static async Task<byte[]> SynthesizeSpeechAsync(EnvironmentConfiguration environmentConfiguration, string text)
        {
            string ttsUrl = string.Format(ElevenLabsTextToSpeechUrlFormat, environmentConfiguration.ElevenLabsVoiceId);

            Dictionary<string, object> requestBody = new Dictionary<string, object>
            {
                { "text", text },
                { "model_id", "eleven_flash_v2_5" },
                { "voice_settings", new Dictionary<string, object>
                    {
                        { "stability", 0.5 },
                        { "similarity_boost", 0.75 }
                    }
                }
            };

            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, ttsUrl);
            request.Content = new StringContent(Serializer.Serialize(requestBody), Encoding.UTF8, "application/json");
            request.Headers.Accept.ParseAdd("audio/mpeg");
            request.Headers.Add("xi-api-key", environmentConfiguration.ElevenLabsApiKey);

            HttpResponseMessage response = await HttpClient.SendAsync(request).ConfigureAwait(false);
            byte[] responseBytes = await response.Content.ReadAsByteArrayAsync().ConfigureAwait(false);

            if (!response.IsSuccessStatusCode)
            {
                throw new InvalidOperationException(string.Format(
                    "TTS proxy error ({0}): {1}",
                    (int)response.StatusCode,
                    Encoding.UTF8.GetString(responseBytes)
                ));
            }

            return responseBytes;
        }

        public static async Task<string> TranscribeSpeechWithElevenLabsAsync(EnvironmentConfiguration environmentConfiguration, string audioFilePath)
        {
            using (MultipartFormDataContent content = new MultipartFormDataContent())
            {
                ByteArrayContent audioContent = new ByteArrayContent(File.ReadAllBytes(audioFilePath));
                audioContent.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("audio/wav");
                content.Add(audioContent, "file", Path.GetFileName(audioFilePath));
                content.Add(new StringContent(ElevenLabsSpeechToTextModel), "model_id");

                if (!string.IsNullOrWhiteSpace(environmentConfiguration.WhisperLanguage))
                {
                    content.Add(new StringContent(environmentConfiguration.WhisperLanguage), "language_code");
                }

                HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, ElevenLabsSpeechToTextUrl);
                request.Headers.Add("xi-api-key", environmentConfiguration.ElevenLabsApiKey);
                request.Content = content;

                HttpResponseMessage response = await HttpClient.SendAsync(request).ConfigureAwait(false);
                string responseText = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                if (!response.IsSuccessStatusCode)
                {
                    throw new InvalidOperationException(string.Format(
                        "ElevenLabs speech-to-text error ({0}): {1}",
                        (int)response.StatusCode,
                        responseText
                    ));
                }

                string transcript = ExtractSpeechToTextTranscript(responseText);
                if (string.IsNullOrWhiteSpace(transcript))
                {
                    throw new InvalidOperationException("ElevenLabs returned an empty transcript.");
                }

                return transcript.Trim();
            }
        }

        private static async Task<string> PostAnthropicAsync(object body, string anthropicApiKey)
        {
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, AnthropicMessagesUrl);
            request.Headers.Add("x-api-key", anthropicApiKey);
            request.Headers.Add("anthropic-version", "2023-06-01");

            StringContent content = new StringContent(Serializer.Serialize(body), Encoding.UTF8, "application/json");
            request.Content = content;

            HttpResponseMessage response = await HttpClient.SendAsync(request).ConfigureAwait(false);
            string responseText = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

            if (!response.IsSuccessStatusCode)
            {
                throw new InvalidOperationException(string.Format(
                    "Anthropic error ({0}): {1}",
                    (int)response.StatusCode,
                    responseText
                ));
            }

            return responseText;
        }

        private static async Task<string> PostOpenAiCompatibleAsync(object body, string provider, EnvironmentConfiguration environmentConfiguration)
        {
            string baseUrl = environmentConfiguration.OpenAIBaseUrl;
            if (string.IsNullOrWhiteSpace(baseUrl))
            {
                baseUrl = "https://api.openai.com/v1";
            }

            baseUrl = baseUrl.TrimEnd('/');
            string requestUrl = baseUrl + OpenAiChatCompletionsPath;

            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
            request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", environmentConfiguration.OpenAIApiKey);
            request.Content = new StringContent(Serializer.Serialize(body), Encoding.UTF8, "application/json");

            if (provider == "openai-compatible")
            {
                request.Headers.Add("HTTP-Referer", "https://karl-klammer.local");
                request.Headers.Add("X-Title", "Karl Klammer");
            }

            HttpResponseMessage response = await HttpClient.SendAsync(request).ConfigureAwait(false);
            string responseText = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

            if (!response.IsSuccessStatusCode)
            {
                throw new InvalidOperationException(string.Format(
                    "{0} error ({1}): {2}",
                    provider == "openai-compatible" ? "OpenAI-compatible API" : "OpenAI",
                    (int)response.StatusCode,
                    responseText
                ));
            }

            return responseText;
        }

        private static HttpClient CreateHttpClient()
        {
            HttpClient httpClient = new HttpClient();
            httpClient.Timeout = TimeSpan.FromMinutes(4);
            return httpClient;
        }

        private static string ExtractAnthropicText(string responseBody)
        {
            object rootObject = Serializer.DeserializeObject(responseBody);
            Dictionary<string, object> rootDictionary = rootObject as Dictionary<string, object>;
            if (rootDictionary == null || !rootDictionary.ContainsKey("content"))
            {
                throw new InvalidOperationException("Claude returned an unexpected response body.");
            }

            object[] contentArray = AsObjectArray(rootDictionary["content"]);
            if (contentArray == null)
            {
                throw new InvalidOperationException("Claude response did not contain text blocks.");
            }

            StringBuilder textBuilder = new StringBuilder();
            foreach (object contentItem in contentArray)
            {
                Dictionary<string, object> contentDictionary = contentItem as Dictionary<string, object>;
                if (contentDictionary == null)
                {
                    continue;
                }

                object typeValue;
                if (!contentDictionary.TryGetValue("type", out typeValue) || !string.Equals(typeValue as string, "text", StringComparison.Ordinal))
                {
                    continue;
                }

                object textValue;
                if (contentDictionary.TryGetValue("text", out textValue) && textValue != null)
                {
                    if (textBuilder.Length > 0)
                    {
                        textBuilder.AppendLine();
                    }
                    textBuilder.Append(textValue.ToString());
                }
            }

            string extractedText = textBuilder.ToString().Trim();
            if (string.IsNullOrWhiteSpace(extractedText))
            {
                throw new InvalidOperationException("Claude returned an empty text response.");
            }

            return extractedText;
        }

        private static string ExtractOpenAiText(string responseBody)
        {
            object rootObject = Serializer.DeserializeObject(responseBody);
            Dictionary<string, object> rootDictionary = rootObject as Dictionary<string, object>;
            if (rootDictionary == null || !rootDictionary.ContainsKey("choices"))
            {
                throw new InvalidOperationException("OpenAI returned an unexpected response body.");
            }

            object[] choices = AsObjectArray(rootDictionary["choices"]);
            if (choices == null || choices.Length == 0)
            {
                throw new InvalidOperationException("OpenAI did not return any choices.");
            }

            Dictionary<string, object> firstChoice = choices[0] as Dictionary<string, object>;
            if (firstChoice == null || !firstChoice.ContainsKey("message"))
            {
                throw new InvalidOperationException("OpenAI response did not contain a message.");
            }

            Dictionary<string, object> message = firstChoice["message"] as Dictionary<string, object>;
            if (message == null || !message.ContainsKey("content"))
            {
                throw new InvalidOperationException("OpenAI response message did not contain content.");
            }

            object content = message["content"];
            string stringContent = content as string;
            if (!string.IsNullOrWhiteSpace(stringContent))
            {
                return stringContent.Trim();
            }

            object[] contentParts = AsObjectArray(content);
            if (contentParts == null)
            {
                throw new InvalidOperationException("OpenAI response content format was unsupported.");
            }

            StringBuilder builder = new StringBuilder();
            foreach (object part in contentParts)
            {
                Dictionary<string, object> partDictionary = part as Dictionary<string, object>;
                if (partDictionary == null)
                {
                    continue;
                }

                object typeValue;
                if (!partDictionary.TryGetValue("type", out typeValue))
                {
                    continue;
                }

                if (!string.Equals(typeValue as string, "text", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(typeValue as string, "output_text", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                object textValue;
                if (partDictionary.TryGetValue("text", out textValue) && textValue != null)
                {
                    if (builder.Length > 0)
                    {
                        builder.AppendLine();
                    }

                    builder.Append(textValue.ToString());
                }
            }

            string extractedText = builder.ToString().Trim();
            if (string.IsNullOrWhiteSpace(extractedText))
            {
                throw new InvalidOperationException("OpenAI returned an empty text response.");
            }

            return extractedText;
        }

        private static JavaScriptSerializer CreateSerializer()
        {
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            serializer.MaxJsonLength = int.MaxValue;
            return serializer;
        }

        private static string BuildCompanionPrompt(AppSettings settings, IList<AutomationRecipe> recipes)
        {
            StringBuilder builder = new StringBuilder();
            builder.AppendLine(LoadSoulPrompt().Trim());
            builder.AppendLine();
            builder.AppendLine(CompanionBehaviorRules);
            builder.AppendLine();
            builder.AppendLine(GetModeInstructions(settings));

            if (settings != null && settings.SuggestAutomations)
            {
                builder.AppendLine();
                builder.AppendLine("automation suggestions:");
                builder.AppendLine("- when you notice a repeatable workflow, you may suggest one reusable recipe.");
                builder.AppendLine("- append exactly one machine-readable tag only when the recipe would be genuinely useful:");
                builder.AppendLine("[AUTOMATION:name|short imperative prompt]");
            }

            if (recipes != null && recipes.Count > 0)
            {
                builder.AppendLine();
                builder.AppendLine("known saved recipes:");
                foreach (AutomationRecipe recipe in recipes.Take(5))
                {
                    builder.AppendLine("- " + recipe.Name + ": " + recipe.Prompt);
                }
            }

            return builder.ToString().Trim();
        }

        private static string BuildUserPrompt(string prompt, IList<KnowledgeChunk> knowledgeChunks)
        {
            StringBuilder builder = new StringBuilder();
            builder.AppendLine((prompt ?? string.Empty).Trim());

            if (knowledgeChunks != null && knowledgeChunks.Count > 0)
            {
                builder.AppendLine();
                builder.AppendLine("local knowledge snippets:");
                foreach (KnowledgeChunk chunk in knowledgeChunks.Take(3))
                {
                    builder.AppendLine("[source: " + chunk.Title + "]");
                    builder.AppendLine(chunk.Text);
                    builder.AppendLine();
                }

                builder.AppendLine("use the local knowledge when relevant and prefer it over guessing.");
            }

            return builder.ToString().Trim();
        }

        private static string LoadSoulPrompt()
        {
            try
            {
                string soulPath = Path.Combine(Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..")), "SOUL.md");
                if (File.Exists(soulPath))
                {
                    string soulText = File.ReadAllText(soulPath, Encoding.UTF8).Trim();
                    if (!string.IsNullOrWhiteSpace(soulText))
                    {
                        return soulText;
                    }
                }
            }
            catch
            {
            }

            return DefaultSoulPrompt;
        }

        private static string GetModeInstructions(AppSettings settings)
        {
            string mode = settings == null ? "companion" : (settings.CompanionMode ?? "companion").Trim().ToLowerInvariant();
            switch (mode)
            {
                case "agent":
                    return "mode: agent router. think in terms of delegation, tool choice, handoffs, and what specialist should handle the task next.";
                case "automation":
                    return "mode: automation. prioritize reusable workflows, repeatable steps, and compact action-oriented output.";
                case "watch":
                    return "mode: watch me once. infer repeatable rituals from the visible workflow, watch for multi-step patterns, and suggest how Karl Klammer could reuse them next time.";
                default:
                    return "mode: companion. prioritize direct help, visible guidance, and concise conversational support.";
            }
        }

        private static string NormalizeProvider(string provider)
        {
            string normalized = (provider ?? string.Empty).Trim().ToLowerInvariant();
            if (normalized == "openai" || normalized == "openai-compatible")
            {
                return normalized;
            }

            return "anthropic";
        }

        private static object[] AsObjectArray(object value)
        {
            if (value == null)
            {
                return null;
            }

            object[] directArray = value as object[];
            if (directArray != null)
            {
                return directArray;
            }

            ArrayList arrayList = value as ArrayList;
            if (arrayList != null)
            {
                object[] result = new object[arrayList.Count];
                arrayList.CopyTo(result);
                return result;
            }

            return null;
        }

        private static string ExtractSpeechToTextTranscript(string responseText)
        {
            object parsed = Serializer.DeserializeObject(responseText);
            Dictionary<string, object> root = parsed as Dictionary<string, object>;
            if (root == null)
            {
                return string.Empty;
            }

            object textValue;
            if (root.TryGetValue("text", out textValue) && textValue != null)
            {
                return textValue.ToString();
            }

            object wordsValue;
            object[] words = null;
            if (root.TryGetValue("words", out wordsValue))
            {
                words = wordsValue as object[];
            }

            if (words == null || words.Length == 0)
            {
                return string.Empty;
            }

            List<string> parts = new List<string>();
            foreach (object wordEntry in words)
            {
                Dictionary<string, object> wordObject = wordEntry as Dictionary<string, object>;
                if (wordObject == null)
                {
                    continue;
                }

                object wordText;
                if (wordObject.TryGetValue("text", out wordText) && wordText != null)
                {
                    parts.Add(wordText.ToString());
                }
            }

            return string.Join(" ", parts.ToArray()).Trim();
        }
    }

    internal sealed class MicrophoneRecorder : IDisposable
    {
        private const string RecordingAlias = "karlklammerrec";
        private bool _isRecording;
        private string _recordingPath;

        [DllImport("winmm.dll", CharSet = CharSet.Auto)]
        private static extern int mciSendString(string command, StringBuilder returnValue, int returnLength, IntPtr callbackHandle);

        [DllImport("winmm.dll", CharSet = CharSet.Auto)]
        private static extern bool mciGetErrorString(int errorCode, StringBuilder errorText, int errorTextSize);

        public bool IsRecording
        {
            get { return _isRecording; }
        }

        public void Start()
        {
            if (_isRecording)
            {
                return;
            }

            Directory.CreateDirectory(AppSettings.StorageRoot);
            _recordingPath = Path.Combine(AppSettings.StorageRoot, "clicky-recording-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".wav");

            CloseAliasQuietly();
            SendCommand("open new type waveaudio alias " + RecordingAlias);

            try
            {
                SendCommand("set " + RecordingAlias + " time format ms");
                SendCommand("record " + RecordingAlias);
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
            if (!_isRecording)
            {
                throw new InvalidOperationException("No microphone recording is currently running.");
            }

            try
            {
                SendCommand("stop " + RecordingAlias);
                SendCommand("save " + RecordingAlias + " " + QuotePath(_recordingPath));
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
        }

        public void Dispose()
        {
            Cancel();
        }

        private static void SendCommand(string command)
        {
            int errorCode = mciSendString(command, null, 0, IntPtr.Zero);
            if (errorCode == 0)
            {
                return;
            }

            StringBuilder errorText = new StringBuilder(256);
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

    internal static class WhisperClient
    {
        public static async Task<string> TranscribeAsync(EnvironmentConfiguration environmentConfiguration, string audioFilePath)
        {
            string outputDirectory = Path.Combine(AppSettings.StorageRoot, "whisper-output-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(outputDirectory);

            string transcriptPath = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(audioFilePath) + ".txt");
            string arguments = BuildArguments(environmentConfiguration, audioFilePath, outputDirectory);

            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = environmentConfiguration.WhisperPythonCommand,
                Arguments = arguments,
                WorkingDirectory = AppSettings.StorageRoot,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };
            startInfo.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8";

            using (Process process = new Process())
            {
                process.StartInfo = startInfo;

                if (!process.Start())
                {
                    throw new InvalidOperationException("Failed to start the local Whisper process.");
                }

                Task<string> standardOutputTask = process.StandardOutput.ReadToEndAsync();
                Task<string> standardErrorTask = process.StandardError.ReadToEndAsync();
                await Task.Run(() => process.WaitForExit()).ConfigureAwait(false);

                string standardOutput = await standardOutputTask.ConfigureAwait(false);
                string standardError = await standardErrorTask.ConfigureAwait(false);

                if (process.ExitCode != 0)
                {
                    throw new InvalidOperationException(BuildWhisperErrorMessage(environmentConfiguration, standardError, standardOutput));
                }
            }

            if (!File.Exists(transcriptPath))
            {
                throw new InvalidOperationException("Whisper finished without writing a transcript file. Check that the local model can run and ffmpeg is installed.");
            }

            string transcript = File.ReadAllText(transcriptPath, Encoding.UTF8).Trim();
            if (string.IsNullOrWhiteSpace(transcript))
            {
                throw new InvalidOperationException("Whisper returned an empty transcript.");
            }

            TryDeleteDirectory(outputDirectory);
            return transcript;
        }

        private static string BuildArguments(EnvironmentConfiguration environmentConfiguration, string audioFilePath, string outputDirectory)
        {
            List<string> arguments = new List<string>();
            arguments.Add("-m");
            arguments.Add("whisper");
            arguments.Add(QuoteArgument(audioFilePath));
            arguments.Add("--model");
            arguments.Add(QuoteArgument(environmentConfiguration.WhisperModel));
            arguments.Add("--task");
            arguments.Add("transcribe");
            arguments.Add("--fp16");
            arguments.Add("False");
            arguments.Add("--verbose");
            arguments.Add("False");
            arguments.Add("--output_format");
            arguments.Add("txt");
            arguments.Add("--output_dir");
            arguments.Add(QuoteArgument(outputDirectory));

            if (!string.IsNullOrWhiteSpace(environmentConfiguration.WhisperLanguage))
            {
                arguments.Add("--language");
                arguments.Add(QuoteArgument(environmentConfiguration.WhisperLanguage));
            }

            return string.Join(" ", arguments.ToArray());
        }

        private static string BuildWhisperErrorMessage(EnvironmentConfiguration environmentConfiguration, string standardError, string standardOutput)
        {
            string detail = string.IsNullOrWhiteSpace(standardError) ? standardOutput : standardError;
            if (string.IsNullOrWhiteSpace(detail))
            {
                detail = "Whisper exited without a detailed error message.";
            }

            return string.Format(
                "Local Whisper failed.\r\npython: {0}\r\nmodel: {1}\r\n\r\n{2}",
                environmentConfiguration.WhisperPythonCommand,
                environmentConfiguration.WhisperModel,
                detail.Trim()
            );
        }

        private static string QuoteArgument(string value)
        {
            return "\"" + value.Replace("\"", "\\\"") + "\"";
        }

        private static void TryDeleteDirectory(string directoryPath)
        {
            try
            {
                if (Directory.Exists(directoryPath))
                {
                    Directory.Delete(directoryPath, true);
                }
            }
            catch
            {
            }
        }
    }

    internal static class SpeechToTextClient
    {
        public static Task<string> TranscribeAsync(EnvironmentConfiguration environmentConfiguration, string audioFilePath)
        {
            if (string.Equals(environmentConfiguration.SpeechToTextProvider, "elevenlabs", StringComparison.OrdinalIgnoreCase))
            {
                return DirectApiClient.TranscribeSpeechWithElevenLabsAsync(environmentConfiguration, audioFilePath);
            }

            return WhisperClient.TranscribeAsync(environmentConfiguration, audioFilePath);
        }

        public static string GetProviderLabel(EnvironmentConfiguration environmentConfiguration)
        {
            return string.Equals(environmentConfiguration.SpeechToTextProvider, "elevenlabs", StringComparison.OrdinalIgnoreCase)
                ? "elevenlabs"
                : "local whisper";
        }
    }

    internal sealed class CodexRunResult
    {
        public int ExitCode { get; set; }
        public string OutputFilePath { get; set; }
    }

    internal sealed class OpenClawRunResult
    {
        public string OutputFilePath { get; set; }
        public string ResponseText { get; set; }
    }

    internal static class CodexClient
    {
        private const string CompletionMessage = "codex session ist jetzt abgeschlossen";
        private static readonly Regex TriggerRegex = new Regex(@"\b(?:nimm|nim|nehm|nehm|mit)\s+(?:den\s+)?(?:codex|kodex|kodes|codecs|kodexx)\b[\s,:-]*", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        private static readonly Regex TriggerWithScreenRegex = new Regex(@"\b(?:nimm|nim|nehm|mit)\s+(?:den\s+)?(?:codex|kodex|kodes)\s+(?:mit|mids?|plus)\s+(?:screen|screenshot|bild|main\s*screen|hauptbildschirm|hauptscreen)\b[\s,:-]*", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        public static bool IsTriggered(string prompt)
        {
            string normalizedPrompt = NormalizePrompt(prompt);
            return !string.IsNullOrWhiteSpace(prompt) && (TriggerWithScreenRegex.IsMatch(normalizedPrompt) || TriggerRegex.IsMatch(normalizedPrompt));
        }

        public static string RemoveTrigger(string prompt)
        {
            if (string.IsNullOrWhiteSpace(prompt))
            {
                return string.Empty;
            }

            string normalizedPrompt = NormalizePrompt(prompt);
            normalizedPrompt = TriggerWithScreenRegex.Replace(normalizedPrompt, string.Empty, 1).Trim();
            return TriggerRegex.Replace(normalizedPrompt, string.Empty, 1).Trim();
        }

        public static bool ShouldAttachScreens(string prompt)
        {
            return !string.IsNullOrWhiteSpace(prompt) && TriggerWithScreenRegex.IsMatch(NormalizePrompt(prompt));
        }

        public static string GetCompletionMessage()
        {
            return CompletionMessage;
        }

        private static string NormalizePrompt(string prompt)
        {
            string normalized = prompt.ToLowerInvariant();
            normalized = normalized.Replace("kodex", "codex");
            normalized = normalized.Replace("kodes", "codex");
            normalized = normalized.Replace("codecs", "codex");
            normalized = normalized.Replace("codexx", "codex");
            normalized = normalized.Replace("nehm", "nimm");
            normalized = normalized.Replace("nehm", "nimm");
            normalized = Regex.Replace(normalized, @"\s+", " ").Trim();
            return normalized;
        }

        public static async Task<CodexRunResult> RunAsync(EnvironmentConfiguration environmentConfiguration, string prompt, IList<string> imagePaths = null)
        {
            string outputDirectory = GetCodexOutputDirectory();
            Directory.CreateDirectory(outputDirectory);
            string timestamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
            string outputFilePath = Path.Combine(outputDirectory, "karl-klammer-codex-" + timestamp + ".txt");
            string workingDirectory = ResolveWorkingDirectory(environmentConfiguration);

            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                WorkingDirectory = workingDirectory,
                UseShellExecute = false,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };
            AddArgument(startInfo, "/c");
            AddArgument(startInfo, ResolveCodexCommand(environmentConfiguration));
            AddArgument(startInfo, "exec");
            AddArgument(startInfo, "--full-auto");
            AddArgument(startInfo, "--skip-git-repo-check");
            AddArgument(startInfo, "-C");
            AddArgument(startInfo, workingDirectory);
            AddArgument(startInfo, "-o");
            AddArgument(startInfo, outputFilePath);
            if (imagePaths != null)
            {
                foreach (string imagePath in imagePaths)
                {
                    if (!string.IsNullOrWhiteSpace(imagePath))
                    {
                        AddArgument(startInfo, "-i");
                        AddArgument(startInfo, imagePath);
                    }
                }
            }
            AddArgument(startInfo, "-");

            using (Process process = new Process())
            {
                process.StartInfo = startInfo;

                if (!process.Start())
                {
                    throw new InvalidOperationException("Failed to start the local Codex process.");
                }

                byte[] promptBytes = new UTF8Encoding(false).GetBytes(prompt ?? string.Empty);
                await process.StandardInput.BaseStream.WriteAsync(promptBytes, 0, promptBytes.Length).ConfigureAwait(false);
                await process.StandardInput.BaseStream.FlushAsync().ConfigureAwait(false);
                process.StandardInput.Close();

                Task<string> standardOutputTask = process.StandardOutput.ReadToEndAsync();
                Task<string> standardErrorTask = process.StandardError.ReadToEndAsync();
                bool exited = await Task.Run(() => process.WaitForExit(environmentConfiguration.CodexTimeoutSeconds * 1000)).ConfigureAwait(false);

                if (!exited)
                {
                    try
                    {
                        process.Kill();
                    }
                    catch
                    {
                    }

                    throw new InvalidOperationException("Codex timed out before finishing.");
                }

                string standardOutput = await standardOutputTask.ConfigureAwait(false);
                string standardError = await standardErrorTask.ConfigureAwait(false);
                WriteOutputFile(outputFilePath, workingDirectory, prompt, process.ExitCode, standardOutput, standardError);

                if (process.ExitCode != 0)
                {
                    throw new InvalidOperationException("Codex failed. Check the codex output file: " + outputFilePath);
                }

                return new CodexRunResult
                {
                    ExitCode = process.ExitCode,
                    OutputFilePath = outputFilePath
                };
            }
        }

        private static string ResolveWorkingDirectory(EnvironmentConfiguration environmentConfiguration)
        {
            if (!string.IsNullOrWhiteSpace(environmentConfiguration.CodexWorkingDirectory))
            {
                Directory.CreateDirectory(environmentConfiguration.CodexWorkingDirectory);
                return environmentConfiguration.CodexWorkingDirectory;
            }

            string defaultWorkingDirectory = Path.Combine(Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..")), "playground");
            Directory.CreateDirectory(defaultWorkingDirectory);
            return defaultWorkingDirectory;
        }

        private static string GetCodexOutputDirectory()
        {
            return Path.Combine(Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..")), "codex output");
        }

        private static void AddArgument(ProcessStartInfo startInfo, string value)
        {
            startInfo.Arguments = string.IsNullOrWhiteSpace(startInfo.Arguments)
                ? QuoteArgument(value)
                : startInfo.Arguments + " " + QuoteArgument(value);
        }

        private static string ResolveCodexCommand(EnvironmentConfiguration environmentConfiguration)
        {
            if (!string.IsNullOrWhiteSpace(environmentConfiguration.CodexCommand) && File.Exists(environmentConfiguration.CodexCommand))
            {
                return environmentConfiguration.CodexCommand;
            }

            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string defaultCommandScript = Path.Combine(appData, "npm", "codex.cmd");
            if (File.Exists(defaultCommandScript))
            {
                return defaultCommandScript;
            }

            return environmentConfiguration.CodexCommand;
        }

        private static void WriteOutputFile(string outputFilePath, string workingDirectory, string prompt, int exitCode, string standardOutput, string standardError)
        {
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("Karl Klammer codex run");
            builder.AppendLine("timestamp: " + DateTime.Now.ToString("O"));
            builder.AppendLine("working_directory: " + workingDirectory);
            builder.AppendLine("exit_code: " + exitCode.ToString());
            builder.AppendLine();
            builder.AppendLine("prompt:");
            builder.AppendLine(prompt);
            builder.AppendLine();
            builder.AppendLine("stdout:");
            builder.AppendLine(string.IsNullOrWhiteSpace(standardOutput) ? "<empty>" : standardOutput.Trim());
            builder.AppendLine();
            builder.AppendLine("stderr:");
            builder.AppendLine(string.IsNullOrWhiteSpace(standardError) ? "<empty>" : standardError.Trim());
            File.WriteAllText(outputFilePath, builder.ToString(), Encoding.UTF8);
        }

        private static string QuoteArgument(string value)
        {
            return "\"" + (value ?? string.Empty).Replace("\"", "\\\"") + "\"";
        }
    }

    internal static class ClaudeCodeClient
    {
        private const string CompletionMessage = "claude code session ist jetzt abgeschlossen";
        private static readonly Regex TriggerRegex = new Regex(@"\b(?:nimm|nim|nehm|mit)\s+(?:den\s+)?(?:claude|cloud|clod|klod|klode)\s+code\b[\s,:-]*", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        public static bool IsTriggered(string prompt)
        {
            return !string.IsNullOrWhiteSpace(prompt) && TriggerRegex.IsMatch(NormalizePrompt(prompt));
        }

        public static string RemoveTrigger(string prompt)
        {
            if (string.IsNullOrWhiteSpace(prompt))
            {
                return string.Empty;
            }

            return TriggerRegex.Replace(NormalizePrompt(prompt), string.Empty, 1).Trim();
        }

        public static string GetCompletionMessage()
        {
            return CompletionMessage;
        }

        public static async Task<CodexRunResult> RunAsync(EnvironmentConfiguration environmentConfiguration, string prompt)
        {
            string outputDirectory = GetOutputDirectory();
            Directory.CreateDirectory(outputDirectory);
            string timestamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
            string outputFilePath = Path.Combine(outputDirectory, "karl-klammer-claude-code-" + timestamp + ".txt");
            string workingDirectory = ResolveWorkingDirectory(environmentConfiguration);

            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                WorkingDirectory = workingDirectory,
                UseShellExecute = false,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };
            AddArgument(startInfo, "/c");
            AddArgument(startInfo, ResolveCommand(environmentConfiguration));
            AddArgument(startInfo, "-p");
            AddArgument(startInfo, "--permission-mode");
            AddArgument(startInfo, "bypassPermissions");

            using (Process process = new Process())
            {
                process.StartInfo = startInfo;

                if (!process.Start())
                {
                    throw new InvalidOperationException("Failed to start the local Claude Code process.");
                }

                byte[] promptBytes = new UTF8Encoding(false).GetBytes(prompt ?? string.Empty);
                await process.StandardInput.BaseStream.WriteAsync(promptBytes, 0, promptBytes.Length).ConfigureAwait(false);
                await process.StandardInput.BaseStream.FlushAsync().ConfigureAwait(false);
                process.StandardInput.Close();

                Task<string> standardOutputTask = process.StandardOutput.ReadToEndAsync();
                Task<string> standardErrorTask = process.StandardError.ReadToEndAsync();
                bool exited = await Task.Run(() => process.WaitForExit(environmentConfiguration.CodexTimeoutSeconds * 1000)).ConfigureAwait(false);

                if (!exited)
                {
                    try
                    {
                        process.Kill();
                    }
                    catch
                    {
                    }

                    throw new InvalidOperationException("Claude Code timed out before finishing.");
                }

                string standardOutput = await standardOutputTask.ConfigureAwait(false);
                string standardError = await standardErrorTask.ConfigureAwait(false);
                WriteOutputFile(outputFilePath, workingDirectory, prompt, process.ExitCode, standardOutput, standardError);

                if (process.ExitCode != 0)
                {
                    throw new InvalidOperationException("Claude Code failed. Check the output file: " + outputFilePath);
                }

                return new CodexRunResult
                {
                    ExitCode = process.ExitCode,
                    OutputFilePath = outputFilePath
                };
            }
        }

        private static string NormalizePrompt(string prompt)
        {
            string normalized = prompt.ToLowerInvariant();
            normalized = normalized.Replace("cloud code", "claude code");
            normalized = normalized.Replace("clod code", "claude code");
            normalized = normalized.Replace("klod code", "claude code");
            normalized = normalized.Replace("klode code", "claude code");
            normalized = normalized.Replace("nehm", "nimm");
            normalized = Regex.Replace(normalized, @"\s+", " ").Trim();
            return normalized;
        }

        private static string ResolveWorkingDirectory(EnvironmentConfiguration environmentConfiguration)
        {
            if (!string.IsNullOrWhiteSpace(environmentConfiguration.CodexWorkingDirectory))
            {
                Directory.CreateDirectory(environmentConfiguration.CodexWorkingDirectory);
                return environmentConfiguration.CodexWorkingDirectory;
            }

            string defaultWorkingDirectory = Path.Combine(Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..")), "playground");
            Directory.CreateDirectory(defaultWorkingDirectory);
            return defaultWorkingDirectory;
        }

        private static string GetOutputDirectory()
        {
            return Path.Combine(Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..")), "codex output");
        }

        private static string ResolveCommand(EnvironmentConfiguration environmentConfiguration)
        {
            if (!string.IsNullOrWhiteSpace(environmentConfiguration.ClaudeCodeCommand) && File.Exists(environmentConfiguration.ClaudeCodeCommand))
            {
                return environmentConfiguration.ClaudeCodeCommand;
            }

            return environmentConfiguration.ClaudeCodeCommand;
        }

        private static void AddArgument(ProcessStartInfo startInfo, string value)
        {
            startInfo.Arguments = string.IsNullOrWhiteSpace(startInfo.Arguments)
                ? QuoteArgument(value)
                : startInfo.Arguments + " " + QuoteArgument(value);
        }

        private static void WriteOutputFile(string outputFilePath, string workingDirectory, string prompt, int exitCode, string standardOutput, string standardError)
        {
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("Karl Klammer claude code run");
            builder.AppendLine("timestamp: " + DateTime.Now.ToString("O"));
            builder.AppendLine("working_directory: " + workingDirectory);
            builder.AppendLine("exit_code: " + exitCode.ToString());
            builder.AppendLine();
            builder.AppendLine("prompt:");
            builder.AppendLine(prompt);
            builder.AppendLine();
            builder.AppendLine("stdout:");
            builder.AppendLine(string.IsNullOrWhiteSpace(standardOutput) ? "<empty>" : standardOutput.Trim());
            builder.AppendLine();
            builder.AppendLine("stderr:");
            builder.AppendLine(string.IsNullOrWhiteSpace(standardError) ? "<empty>" : standardError.Trim());
            File.WriteAllText(outputFilePath, builder.ToString(), Encoding.UTF8);
        }

        private static string QuoteArgument(string value)
        {
            return "\"" + (value ?? string.Empty).Replace("\"", "\\\"") + "\"";
        }
    }

    internal static class OpenClawClient
    {
        private const string CompletionMessage = "openclaw session ist jetzt abgeschlossen";
        private static readonly Regex TriggerRegex = new Regex(@"\b(?:nimm|nim|nehm|mit)\s+(?:den\s+)?(?:(?:open|oben|orpen|onpen|oppen)\s*cl(?:aw|au)|openclaw|klaus|claus|claws)\b[\s,:-]*", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        public static bool IsTriggered(string prompt)
        {
            return !string.IsNullOrWhiteSpace(prompt) && TriggerRegex.IsMatch(NormalizePrompt(prompt));
        }

        public static string RemoveTrigger(string prompt)
        {
            if (string.IsNullOrWhiteSpace(prompt))
            {
                return string.Empty;
            }

            return TriggerRegex.Replace(NormalizePrompt(prompt), string.Empty, 1).Trim();
        }

        public static string GetCompletionMessage()
        {
            return CompletionMessage;
        }

        public static async Task<OpenClawRunResult> RunAsync(EnvironmentConfiguration environmentConfiguration, string prompt)
        {
            string command = ResolveCommand(environmentConfiguration);
            if (string.IsNullOrWhiteSpace(command))
            {
                throw new InvalidOperationException("OPENCLAW_COMMAND is missing.");
            }

            string outputDirectory = GetOutputDirectory();
            Directory.CreateDirectory(outputDirectory);
            string timestamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
            string outputFilePath = Path.Combine(outputDirectory, "karl-klammer-openclaw-" + timestamp + ".txt");

            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                WorkingDirectory = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..")),
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };
            AddArgument(startInfo, "/c");
            AddArgument(startInfo, command);
            AddArgument(startInfo, "agent");
            AddArgument(startInfo, "--agent");
            AddArgument(startInfo, ResolveAgentId(environmentConfiguration));
            AddArgument(startInfo, "--message");
            AddArgument(startInfo, prompt ?? string.Empty);
            AddArgument(startInfo, "--timeout");
            AddArgument(startInfo, environmentConfiguration.OpenClawTimeoutSeconds.ToString());

            using (Process process = new Process())
            {
                process.StartInfo = startInfo;

                if (!process.Start())
                {
                    throw new InvalidOperationException("Failed to start OpenClaw.");
                }

                Task<string> standardOutputTask = process.StandardOutput.ReadToEndAsync();
                Task<string> standardErrorTask = process.StandardError.ReadToEndAsync();
                bool exited = await Task.Run(() => process.WaitForExit(environmentConfiguration.OpenClawTimeoutSeconds * 1000 + 15000)).ConfigureAwait(false);

                if (!exited)
                {
                    try
                    {
                        process.Kill();
                    }
                    catch
                    {
                    }

                    throw new InvalidOperationException("OpenClaw timed out before finishing.");
                }

                string standardOutput = await standardOutputTask.ConfigureAwait(false);
                string standardError = await standardErrorTask.ConfigureAwait(false);
                string responseText = string.IsNullOrWhiteSpace(standardOutput) ? string.Empty : standardOutput.Trim();

                WriteOutputFile(
                    outputFilePath,
                    command,
                    ResolveAgentId(environmentConfiguration),
                    prompt,
                    process.ExitCode,
                    standardOutput,
                    standardError);

                if (process.ExitCode != 0)
                {
                    string detail = string.IsNullOrWhiteSpace(standardError) ? standardOutput : standardError;
                    throw new InvalidOperationException(string.IsNullOrWhiteSpace(detail)
                        ? "OpenClaw failed. Check the output file: " + outputFilePath
                        : detail.Trim());
                }

                if (string.IsNullOrWhiteSpace(responseText))
                {
                    responseText = CompletionMessage;
                }

                return new OpenClawRunResult
                {
                    OutputFilePath = outputFilePath,
                    ResponseText = responseText
                };
            }
        }

        private static string NormalizePrompt(string prompt)
        {
            string normalized = prompt.ToLowerInvariant();
            normalized = normalized.Replace("nehm", "nimm");
            normalized = normalized.Replace("nehm", "nimm");
            normalized = normalized.Replace("nimm", "nimm");
            normalized = normalized.Replace("obenclaw", "openclaw");
            normalized = normalized.Replace("obenclau", "openclaw");
            normalized = normalized.Replace("oben claw", "openclaw");
            normalized = normalized.Replace("oben clau", "openclaw");
            normalized = normalized.Replace("openclo", "openclaw");
            normalized = normalized.Replace("openclau", "openclaw");
            normalized = normalized.Replace("open claw", "openclaw");
            normalized = normalized.Replace("open clau", "openclaw");
            normalized = normalized.Replace("orpenclaw", "openclaw");
            normalized = normalized.Replace("orpenclau", "openclaw");
            normalized = normalized.Replace("orpen claw", "openclaw");
            normalized = normalized.Replace("orpen clau", "openclaw");
            normalized = normalized.Replace("onpenclaw", "openclaw");
            normalized = normalized.Replace("onpenclau", "openclaw");
            normalized = normalized.Replace("onpen claw", "openclaw");
            normalized = normalized.Replace("onpen clau", "openclaw");
            normalized = normalized.Replace("oppenclaw", "openclaw");
            normalized = normalized.Replace("oppenclau", "openclaw");
            normalized = normalized.Replace("oppen claw", "openclaw");
            normalized = normalized.Replace("oppen clau", "openclaw");
            normalized = normalized.Replace("open claww", "openclaw");
            normalized = normalized.Replace("claus", "klaus");
            normalized = normalized.Replace("claws", "klaus");
            normalized = Regex.Replace(normalized, @"\s+", " ").Trim();
            return normalized;
        }

        private static string ResolveCommand(EnvironmentConfiguration environmentConfiguration)
        {
            if (!string.IsNullOrWhiteSpace(environmentConfiguration.OpenClawCommand) && File.Exists(environmentConfiguration.OpenClawCommand))
            {
                return environmentConfiguration.OpenClawCommand;
            }

            return environmentConfiguration.OpenClawCommand;
        }

        private static string ResolveAgentId(EnvironmentConfiguration environmentConfiguration)
        {
            string configuredValue = string.IsNullOrWhiteSpace(environmentConfiguration.OpenClawSessionKey)
                ? "main"
                : environmentConfiguration.OpenClawSessionKey.Trim();
            string[] parts = configuredValue.Split(new[] { ':' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length >= 2 && string.Equals(parts[0], "agent", StringComparison.OrdinalIgnoreCase))
            {
                return parts[1];
            }

            return configuredValue;
        }

        private static void AddArgument(ProcessStartInfo startInfo, string value)
        {
            startInfo.Arguments = string.IsNullOrWhiteSpace(startInfo.Arguments)
                ? QuoteArgument(value)
                : startInfo.Arguments + " " + QuoteArgument(value);
        }

        private static string QuoteArgument(string value)
        {
            return "\"" + (value ?? string.Empty).Replace("\"", "\\\"") + "\"";
        }

        private static string GetOutputDirectory()
        {
            return Path.Combine(Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..")), "codex output");
        }

        private static void WriteOutputFile(string outputFilePath, string command, string agentId, string prompt, int exitCode, string standardOutput, string standardError)
        {
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("Karl Klammer openclaw run");
            builder.AppendLine("timestamp: " + DateTime.Now.ToString("O"));
            builder.AppendLine("command: " + command);
            builder.AppendLine("agent: " + agentId);
            builder.AppendLine("exit_code: " + exitCode.ToString());
            builder.AppendLine();
            builder.AppendLine("prompt:");
            builder.AppendLine(prompt);
            builder.AppendLine();
            builder.AppendLine("stdout:");
            builder.AppendLine(string.IsNullOrWhiteSpace(standardOutput) ? "<empty>" : standardOutput.Trim());
            builder.AppendLine();
            builder.AppendLine("stderr:");
            builder.AppendLine(string.IsNullOrWhiteSpace(standardError) ? "<empty>" : standardError.Trim());
            File.WriteAllText(outputFilePath, builder.ToString(), Encoding.UTF8);
        }
    }

    internal sealed class PushToTalkHotKeyListener : IDisposable
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

        public event EventHandler HotKeyPressed;
        public event EventHandler HotKeyReleased;

        private readonly HookProc _hookProc;
        private readonly Keys _hotKey;
        private IntPtr _hookHandle;
        private bool _isPressed;

        public PushToTalkHotKeyListener(Keys hotKey)
        {
            _hotKey = hotKey;
            _hookProc = HookCallback;

            using (Process currentProcess = Process.GetCurrentProcess())
            using (ProcessModule currentModule = currentProcess.MainModule)
            {
                IntPtr moduleHandle = currentModule == null ? IntPtr.Zero : GetModuleHandle(currentModule.ModuleName);
                _hookHandle = SetWindowsHookEx(WhKeyboardLl, _hookProc, moduleHandle, 0);
            }
        }

        public bool IsRegistered
        {
            get { return _hookHandle != IntPtr.Zero; }
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                int message = wParam.ToInt32();
                int virtualKeyCode = Marshal.ReadInt32(lParam);

                if (virtualKeyCode == (int)_hotKey)
                {
                    if ((message == WmKeyDown || message == WmSysKeyDown) && !_isPressed)
                    {
                        _isPressed = true;
                        if (HotKeyPressed != null)
                        {
                            HotKeyPressed(this, EventArgs.Empty);
                        }
                    }
                    else if ((message == WmKeyUp || message == WmSysKeyUp) && _isPressed)
                    {
                        _isPressed = false;
                        if (HotKeyReleased != null)
                        {
                            HotKeyReleased(this, EventArgs.Empty);
                        }
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

    internal sealed class CompanionOverlayForm : Form
    {
        private const int OverlayWidth = 360;
        private const int OverlayHeight = 190;
        private readonly Timer _animationTimer;
        private PointF _displayLocation;
        private bool _locationInitialized;
        private bool _bubbleOnLeft;
        private float _phase;
        private string _bubbleText;
        private DateTime _bubbleExpiresAtUtc;
        private Point? _navigationAnchorPoint;
        private DateTime _navigationExpiresAtUtc;
        private CompanionVisualState _state;
        private CompanionVisualState _stateAfterBubble;

        public CompanionOverlayForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            TopMost = true;
            BackColor = Color.Magenta;
            TransparencyKey = Color.Magenta;
            DoubleBuffered = true;
            Width = OverlayWidth;
            Height = OverlayHeight;

            _state = CompanionVisualState.Idle;
            _stateAfterBubble = CompanionVisualState.Idle;

            _animationTimer = new Timer();
            _animationTimer.Interval = 33;
            _animationTimer.Tick += delegate
            {
                AdvanceAnimationFrame();
            };
            _animationTimer.Start();
        }

        protected override bool ShowWithoutActivation
        {
            get { return true; }
        }

        protected override CreateParams CreateParams
        {
            get
            {
                const int WsExTransparent = 0x20;
                const int WsExToolWindow = 0x80;
                const int WsExLayered = 0x80000;
                const int WsExNoActivate = 0x08000000;

                CreateParams createParams = base.CreateParams;
                createParams.ExStyle |= WsExTransparent | WsExToolWindow | WsExLayered | WsExNoActivate;
                return createParams;
            }
        }

        public void SetState(CompanionVisualState state)
        {
            _state = state;
            _navigationAnchorPoint = null;
            Invalidate();
        }

        public void ShowMessage(string bubbleText, CompanionVisualState state, int durationMs, CompanionVisualState stateAfterBubble)
        {
            _bubbleText = (bubbleText ?? string.Empty).Trim();
            if (_bubbleText.Length > 280)
            {
                _bubbleText = _bubbleText.Substring(0, 277) + "...";
            }

            _state = state;
            _stateAfterBubble = stateAfterBubble;
            _bubbleExpiresAtUtc = DateTime.UtcNow.AddMilliseconds(Math.Max(1500, durationMs));
            Invalidate();
        }

        public void NavigateTo(Point anchorPoint, string bubbleText, CompanionVisualState state, int durationMs, CompanionVisualState stateAfterNavigation)
        {
            _navigationAnchorPoint = anchorPoint;
            _navigationExpiresAtUtc = DateTime.UtcNow.AddMilliseconds(Math.Max(2200, durationMs));
            ShowMessage(bubbleText, state, durationMs, stateAfterNavigation);
        }

        protected override void OnPaintBackground(PaintEventArgs e)
        {
            e.Graphics.Clear(Color.Magenta);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            e.Graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

            float bob = (float)Math.Sin(_phase * 1.35f) * 3.0f;
            Rectangle orbRectangle = _bubbleOnLeft
                ? new Rectangle(OverlayWidth - 106, OverlayHeight - 96 + (int)bob, 70, 70)
                : new Rectangle(34, OverlayHeight - 96 + (int)bob, 70, 70);

            DrawTrail(e.Graphics, orbRectangle);

            if (!string.IsNullOrWhiteSpace(_bubbleText))
            {
                DrawBubble(e.Graphics, orbRectangle, _bubbleText);
            }

            DrawStateChip(e.Graphics, orbRectangle);
            DrawCompanionBody(e.Graphics, orbRectangle);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _animationTimer.Dispose();
            }

            base.Dispose(disposing);
        }

        private void AdvanceAnimationFrame()
        {
            if (!string.IsNullOrWhiteSpace(_bubbleText) && DateTime.UtcNow >= _bubbleExpiresAtUtc)
            {
                _bubbleText = null;
                _state = _stateAfterBubble;
            }

            if (_navigationAnchorPoint.HasValue && DateTime.UtcNow >= _navigationExpiresAtUtc)
            {
                _navigationAnchorPoint = null;
            }

            _phase += 0.16f;

            Point focusPoint = _navigationAnchorPoint ?? Cursor.Position;
            Rectangle workingArea = Screen.FromPoint(focusPoint).WorkingArea;
            _bubbleOnLeft = focusPoint.X > workingArea.Left + (int)(workingArea.Width * 0.62);

            Point targetLocation = CalculateTargetLocation(focusPoint, workingArea);
            if (!_locationInitialized)
            {
                _displayLocation = new PointF(targetLocation.X, targetLocation.Y);
                _locationInitialized = true;
            }

            float spring = _navigationAnchorPoint.HasValue ? 0.30f : (_state == CompanionVisualState.Listening ? 0.34f : 0.24f);
            _displayLocation.X += (targetLocation.X - _displayLocation.X) * spring;
            _displayLocation.Y += (targetLocation.Y - _displayLocation.Y) * spring;

            Point nextLocation = new Point(
                (int)Math.Round(_displayLocation.X),
                (int)Math.Round(_displayLocation.Y)
            );

            if (Location != nextLocation)
            {
                Location = nextLocation;
            }

            Invalidate();
        }

        private Point CalculateTargetLocation(Point cursorPosition, Rectangle workingArea)
        {
            int desiredX = _bubbleOnLeft ? cursorPosition.X - OverlayWidth - 18 : cursorPosition.X + 18;
            int desiredY = cursorPosition.Y - OverlayHeight + 102;

            int clampedX = Math.Max(workingArea.Left + 6, Math.Min(desiredX, workingArea.Right - OverlayWidth - 6));
            int clampedY = Math.Max(workingArea.Top + 6, Math.Min(desiredY, workingArea.Bottom - OverlayHeight - 6));

            return new Point(clampedX, clampedY);
        }

        private void DrawTrail(Graphics graphics, Rectangle orbRectangle)
        {
            Color accentColor = GetAccentColor(_state);
            int direction = _bubbleOnLeft ? 1 : -1;

            for (int index = 0; index < 3; index++)
            {
                int size = 7 - index;
                int offsetX = direction * (16 + (index * 14));
                int offsetY = 10 + (index * 8);
                Rectangle trailRectangle = new Rectangle(
                    orbRectangle.X + (orbRectangle.Width / 2) + offsetX - size,
                    orbRectangle.Y + offsetY,
                    size * 2,
                    size * 2
                );

                using (SolidBrush brush = new SolidBrush(Color.FromArgb(110 - (index * 24), accentColor)))
                {
                    graphics.FillEllipse(brush, trailRectangle);
                }
            }
        }

        private void DrawBubble(Graphics graphics, Rectangle orbRectangle, string bubbleText)
        {
            const int maxTextWidth = 228;
            using (Font bubbleFont = new Font("Segoe UI Semibold", 9.5f))
            {
                SizeF measuredText = graphics.MeasureString(bubbleText, bubbleFont, maxTextWidth);
                int bubbleWidth = Math.Min(maxTextWidth + 26, Math.Max(148, (int)Math.Ceiling(measuredText.Width) + 24));
                int bubbleHeight = Math.Max(52, (int)Math.Ceiling(measuredText.Height) + 20);
                int bubbleX = _bubbleOnLeft ? Math.Max(12, orbRectangle.Left - bubbleWidth - 20) : orbRectangle.Right + 12;
                Rectangle bubbleRectangle = new Rectangle(bubbleX, 16, bubbleWidth, bubbleHeight);

                using (GraphicsPath bubblePath = CreateRoundedRectanglePath(bubbleRectangle, 16))
                using (SolidBrush bubbleShadowBrush = new SolidBrush(Color.FromArgb(82, 0, 0, 0)))
                using (SolidBrush bubbleBrush = new SolidBrush(Color.FromArgb(236, 11, 18, 31)))
                using (Pen bubbleBorderPen = new Pen(Color.FromArgb(180, GetAccentColor(_state)), 1.6f))
                using (SolidBrush textBrush = new SolidBrush(Color.White))
                using (SolidBrush pointerBrush = new SolidBrush(Color.FromArgb(220, 18, 28, 45)))
                {
                    Rectangle bubbleShadowRectangle = bubbleRectangle;
                    bubbleShadowRectangle.Offset(0, 4);
                    using (GraphicsPath shadowPath = CreateRoundedRectanglePath(bubbleShadowRectangle, 16))
                    {
                        graphics.FillPath(bubbleShadowBrush, shadowPath);
                    }

                    graphics.FillPath(bubbleBrush, bubblePath);
                    graphics.DrawPath(bubbleBorderPen, bubblePath);

                    PointF[] pointerPoints = _bubbleOnLeft
                        ? new PointF[]
                        {
                            new PointF(bubbleRectangle.Right - 2, bubbleRectangle.Bottom - 16),
                            new PointF(orbRectangle.Left + 10, orbRectangle.Top + 18),
                            new PointF(bubbleRectangle.Right - 16, bubbleRectangle.Bottom - 28)
                        }
                        : new PointF[]
                        {
                            new PointF(bubbleRectangle.Left + 2, bubbleRectangle.Bottom - 16),
                            new PointF(orbRectangle.Right - 10, orbRectangle.Top + 18),
                            new PointF(bubbleRectangle.Left + 16, bubbleRectangle.Bottom - 28)
                        };

                    graphics.FillPolygon(pointerBrush, pointerPoints);

                    RectangleF textRectangle = new RectangleF(
                        bubbleRectangle.Left + 12,
                        bubbleRectangle.Top + 10,
                        bubbleRectangle.Width - 24,
                        bubbleRectangle.Height - 16
                    );
                    graphics.DrawString(bubbleText, bubbleFont, textBrush, textRectangle);
                }
            }
        }

        private void DrawStateChip(Graphics graphics, Rectangle orbRectangle)
        {
            string stateLabel = GetStateLabel(_state);
            using (Font chipFont = new Font("Segoe UI Semibold", 8.5f))
            {
                Size chipTextSize = TextRenderer.MeasureText(stateLabel, chipFont);
                int chipWidth = Math.Max(72, chipTextSize.Width + 18);
                int chipHeight = 26;
                int chipX = _bubbleOnLeft ? orbRectangle.Right - chipWidth : orbRectangle.Left;
                int chipY = orbRectangle.Bottom + 6;
                Rectangle chipRectangle = new Rectangle(chipX, chipY, chipWidth, chipHeight);

                using (GraphicsPath chipPath = CreateRoundedRectanglePath(chipRectangle, 12))
                using (SolidBrush chipBrush = new SolidBrush(Color.FromArgb(220, 10, 18, 28)))
                using (Pen chipBorderPen = new Pen(Color.FromArgb(160, GetAccentColor(_state)), 1.2f))
                using (SolidBrush chipTextBrush = new SolidBrush(Color.White))
                {
                    graphics.FillPath(chipBrush, chipPath);
                    graphics.DrawPath(chipBorderPen, chipPath);
                    TextRenderer.DrawText(graphics, stateLabel, chipFont, chipRectangle, chipTextBrush.Color, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                }
            }
        }

        private void DrawCompanionBody(Graphics graphics, Rectangle orbRectangle)
        {
            Color accentColor = GetAccentColor(_state);
            Rectangle glowRectangle = Rectangle.Inflate(orbRectangle, 16, 16);
            Rectangle innerGlowRectangle = Rectangle.Inflate(orbRectangle, 8, 8);
            Rectangle eyeRectangle = new Rectangle(orbRectangle.Left + 19, orbRectangle.Top + 18, 32, 24);
            Rectangle pupilRectangle = new Rectangle(eyeRectangle.Left + 9, eyeRectangle.Top + 5, 12, 12);

            using (SolidBrush outerGlowBrush = new SolidBrush(Color.FromArgb(48, accentColor)))
            using (SolidBrush innerGlowBrush = new SolidBrush(Color.FromArgb(74, accentColor)))
            using (LinearGradientBrush bodyBrush = new LinearGradientBrush(orbRectangle, Color.FromArgb(255, 25, 38, 56), Color.FromArgb(255, 9, 13, 22), LinearGradientMode.ForwardDiagonal))
            using (Pen ringPen = new Pen(Color.FromArgb(185, accentColor), 2.2f))
            using (Pen innerRingPen = new Pen(Color.FromArgb(82, 255, 255, 255), 1.0f))
            using (SolidBrush eyeBrush = new SolidBrush(Color.FromArgb(248, 252, 255)))
            using (SolidBrush pupilBrush = new SolidBrush(Color.FromArgb(20, 28, 40)))
            using (SolidBrush highlightBrush = new SolidBrush(Color.FromArgb(160, 255, 255, 255)))
            {
                graphics.FillEllipse(outerGlowBrush, glowRectangle);
                graphics.FillEllipse(innerGlowBrush, innerGlowRectangle);
                graphics.FillEllipse(bodyBrush, orbRectangle);
                graphics.DrawEllipse(ringPen, orbRectangle);
                graphics.DrawEllipse(innerRingPen, Rectangle.Inflate(orbRectangle, -6, -6));

                Rectangle antennaRectangle = _bubbleOnLeft
                    ? new Rectangle(orbRectangle.Right - 10, orbRectangle.Top - 10, 10, 18)
                    : new Rectangle(orbRectangle.Left, orbRectangle.Top - 10, 10, 18);
                using (Pen antennaPen = new Pen(Color.FromArgb(150, accentColor), 2.0f))
                using (SolidBrush antennaBrush = new SolidBrush(Color.FromArgb(220, accentColor)))
                {
                    Point antennaStart = _bubbleOnLeft ? new Point(orbRectangle.Right - 14, orbRectangle.Top + 6) : new Point(orbRectangle.Left + 14, orbRectangle.Top + 6);
                    Point antennaEnd = new Point(antennaRectangle.Left + (antennaRectangle.Width / 2), antennaRectangle.Top + 8);
                    graphics.DrawLine(antennaPen, antennaStart, antennaEnd);
                    graphics.FillEllipse(antennaBrush, antennaRectangle);
                }

                graphics.FillEllipse(eyeBrush, eyeRectangle);
                graphics.FillEllipse(pupilBrush, pupilRectangle);
                graphics.FillEllipse(highlightBrush, new Rectangle(orbRectangle.Left + 18, orbRectangle.Top + 10, 14, 10));
            }

            switch (_state)
            {
                case CompanionVisualState.Listening:
                    DrawListeningIndicator(graphics, orbRectangle, accentColor);
                    break;
                case CompanionVisualState.Transcribing:
                    DrawTranscribingIndicator(graphics, orbRectangle, accentColor);
                    break;
                case CompanionVisualState.Thinking:
                    DrawThinkingIndicator(graphics, orbRectangle, accentColor);
                    break;
                case CompanionVisualState.Speaking:
                    DrawSpeakingIndicator(graphics, orbRectangle, accentColor);
                    break;
            }
        }

        private void DrawListeningIndicator(Graphics graphics, Rectangle orbRectangle, Color accentColor)
        {
            using (Pen indicatorPen = new Pen(Color.FromArgb(190, accentColor), 2.2f))
            {
                graphics.DrawArc(indicatorPen, Rectangle.Inflate(orbRectangle, 8, 8), 220, 100);
                graphics.DrawArc(indicatorPen, Rectangle.Inflate(orbRectangle, 16, 16), 218, 104);
            }
        }

        private void DrawTranscribingIndicator(Graphics graphics, Rectangle orbRectangle, Color accentColor)
        {
            using (Pen indicatorPen = new Pen(Color.FromArgb(186, accentColor), 2.0f))
            {
                indicatorPen.DashStyle = DashStyle.Dot;
                graphics.DrawArc(indicatorPen, Rectangle.Inflate(orbRectangle, 12, 12), -35, 160);
            }
        }

        private void DrawThinkingIndicator(Graphics graphics, Rectangle orbRectangle, Color accentColor)
        {
            using (SolidBrush indicatorBrush = new SolidBrush(Color.FromArgb(210, accentColor)))
            {
                int startX = orbRectangle.Left + 16;
                int y = orbRectangle.Top - 16;
                for (int index = 0; index < 3; index++)
                {
                    int size = 6 + (index == 1 ? 2 : 0);
                    graphics.FillEllipse(indicatorBrush, new Rectangle(startX + (index * 13), y + Math.Abs(index - 1) * 2, size, size));
                }
            }
        }

        private void DrawSpeakingIndicator(Graphics graphics, Rectangle orbRectangle, Color accentColor)
        {
            int direction = _bubbleOnLeft ? -1 : 1;
            int baseX = _bubbleOnLeft ? orbRectangle.Left - 6 : orbRectangle.Right + 6;
            int centerY = orbRectangle.Top + (orbRectangle.Height / 2);

            using (Pen indicatorPen = new Pen(Color.FromArgb(190, accentColor), 2.2f))
            {
                for (int index = 0; index < 3; index++)
                {
                    int width = 10 + (index * 8);
                    int height = 16 + (index * 8);
                    Rectangle waveRectangle = direction < 0
                        ? new Rectangle(baseX - width, centerY - (height / 2), width, height)
                        : new Rectangle(baseX, centerY - (height / 2), width, height);
                    graphics.DrawArc(indicatorPen, waveRectangle, direction < 0 ? 320 : 220, 80);
                }
            }
        }

        private static string GetStateLabel(CompanionVisualState state)
        {
            switch (state)
            {
                case CompanionVisualState.Listening:
                    return "listening";
                case CompanionVisualState.Transcribing:
                    return "transcribing";
                case CompanionVisualState.Thinking:
                    return "thinking";
                case CompanionVisualState.Speaking:
                    return "speaking";
                default:
                    return "ready";
            }
        }

        private static Color GetAccentColor(CompanionVisualState state)
        {
            switch (state)
            {
                case CompanionVisualState.Listening:
                    return Color.FromArgb(255, 124, 92);
                case CompanionVisualState.Transcribing:
                    return Color.FromArgb(255, 183, 77);
                case CompanionVisualState.Thinking:
                    return Color.FromArgb(88, 196, 255);
                case CompanionVisualState.Speaking:
                    return Color.FromArgb(93, 212, 136);
                default:
                    return Color.FromArgb(88, 196, 255);
            }
        }

        private static GraphicsPath CreateRoundedRectanglePath(Rectangle rectangle, int radius)
        {
            GraphicsPath path = new GraphicsPath();
            int diameter = radius * 2;

            path.AddArc(rectangle.X, rectangle.Y, diameter, diameter, 180, 90);
            path.AddArc(rectangle.Right - diameter, rectangle.Y, diameter, diameter, 270, 90);
            path.AddArc(rectangle.Right - diameter, rectangle.Bottom - diameter, diameter, diameter, 0, 90);
            path.AddArc(rectangle.X, rectangle.Bottom - diameter, diameter, diameter, 90, 90);
            path.CloseFigure();

            return path;
        }
    }

    internal sealed class MainForm : Form
    {
        private const uint MouseeventfLeftdown = 0x0002;
        private const uint MouseeventfLeftup = 0x0004;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);

        private readonly AppSettings _settings;
        private EnvironmentConfiguration _environmentConfiguration;
        private readonly List<ConversationTurn> _conversationHistory;
        private readonly NotifyIcon _notifyIcon;
        private readonly ContextMenuStrip _trayMenu;
        private readonly Panel _headerPanel;
        private readonly Panel _controlPanel;
        private readonly Panel _contextPanel;
        private readonly Panel _workspacePanel;
        private readonly Panel _controlEnvPanel;
        private readonly Panel _controlModePanel;
        private readonly Panel _controlActionPanel;
        private readonly Panel _contextModelPanel;
        private readonly Panel _contextRitualPanel;
        private readonly Panel _contextKnowledgePanel;
        private readonly Panel _contextInsightPanel;
        private readonly Panel _contextInfoPanel1;
        private readonly Panel _contextInfoPanel2;
        private readonly Panel _contextInfoPanel3;
        private readonly Panel _contextInfoPanel4;
        private readonly Panel _activityPanel;
        private PushToTalkHotKeyListener _hotKeyListener;
        private readonly Label _titleLabel;
        private readonly Label _subtitleLabel;
        private readonly Label _controlSectionLabel;
        private readonly Label _contextSectionLabel;
        private readonly Label _conversationSectionLabel;
        private readonly Label _envFileLabel;
        private readonly Label _knowledgeDocsLabel;
        private readonly Label _previewLabel;
        private readonly Label _providerLabel;
        private readonly Label _modelLabel;
        private readonly Label _modeLabel;
        private readonly Label _recipeLabel;
        private readonly Label _learnedRitualsLabel;
        private readonly Label _promptLabel;
        private readonly Label _activityRailLabel;
        private readonly Label _responseLabel;
        private readonly Label _controlEnvSectionLabel;
        private readonly Label _controlModeSectionLabel;
        private readonly Label _controlActionSectionLabel;
        private readonly Label _contextModelSectionLabel;
        private readonly Label _contextRitualSectionLabel;
        private readonly Label _contextKnowledgeSectionLabel;
        private readonly Label _contextInsightSectionLabel;
        private readonly Label _statusLabel;
        private readonly Label _hotkeyLabel;
        private readonly Label _envStatusLabel;
        private readonly Label _statusBadgeLabel;
        private readonly Label _modeBadgeLabel;
        private readonly Label _riskBadgeLabel;
        private readonly Label _screenAwareChip;
        private readonly Label _ritualMemoryChip;
        private readonly Label _safeActionsChip;
        private readonly Label _activeAppLabel;
        private readonly Label _appActionsLabel;
        private readonly Label _knowledgeStatusLabel;
        private readonly Label _proactiveSuggestionLabel;
        private readonly Label _retrievalSourcesLabel;
        private readonly Label _providerHealthLabel;
        private readonly Label _voiceStatusLabel;
        private readonly Label _activeWindowDetailsLabel;
        private readonly Label _recipeManagerLabel;
        private readonly Label _actionInspectorLabel;
        private readonly Label _diagnosticsLabel;
        private readonly TextBox _envPathTextBox;
        private readonly ComboBox _providerComboBox;
        private readonly TextBox _modelTextBox;
        private readonly ComboBox _modeComboBox;
        private readonly ComboBox _recipeComboBox;
        private readonly ComboBox _watchSuggestionComboBox;
        private readonly ListBox _knowledgeListBox;
        private readonly TextBox _knowledgeSearchTextBox;
        private readonly TextBox _knowledgePreviewTextBox;
        private readonly TextBox _recipePreviewTextBox;
        private readonly TextBox _actionInspectorTextBox;
        private readonly CheckBox _speakCheckBox;
        private readonly CheckBox _suggestAutomationsCheckBox;
        private readonly CheckBox _useKnowledgeCheckBox;
        private readonly Button _dictationButton;
        private readonly TextBox _promptTextBox;
        private readonly TextBox _responseTextBox;
        private readonly ListBox _activityListBox;
        private readonly Button _saveButton;
        private readonly Button _reloadEnvButton;
        private readonly Button _testApiButton;
        private readonly Button _clearHistoryButton;
        private readonly Button _reindexKnowledgeButton;
        private readonly Button _saveRecipeButton;
        private readonly Button _runRecipeButton;
        private readonly Button _useContextIdeaButton;
        private readonly Button _runContextIdeaButton;
        private readonly Button _saveWatchIdeaButton;
        private readonly Button _useWatchIdeaButton;
        private readonly Button _replayRitualButton;
        private readonly Button _askButton;
        private readonly Button _copyResponseButton;
        private readonly Button _importKnowledgeButton;
        private readonly Button _refreshKnowledgeButton;
        private readonly Button _reindexSelectedKnowledgeButton;
        private readonly Button _removeKnowledgeButton;
        private readonly Button _deleteRecipeButton;
        private readonly Button _openRitualManagerButton;
        private readonly Button _openHistoryViewerButton;
        private readonly Button _openControlInspectorButton;
        private readonly Button _openProviderVoiceButton;
        private readonly Button _openDiagnosticsButton;
        private readonly Button _openSetupWizardButton;
        private readonly ListBox _recipeManagerListBox;
        private readonly ListBox _diagnosticsListBox;
        private readonly MicrophoneRecorder _microphoneRecorder;
        private readonly List<AutomationRecipe> _automationRecipes;
        private readonly List<WatchSuggestion> _watchSuggestions;
        private CompanionOverlayForm _companionOverlay;
        private Timer _contextRefreshTimer;
        private Keys _pushToTalkKey;
        private bool _quitRequested;
        private bool _isTranscribingSpeech;
        private bool _isCompactLayout;
        private string _currentAudioFilePath;
        private dynamic _audioPlayer;
        private ProactiveSuggestion _currentProactiveSuggestion;
        private ActiveWindowInfo _lastActiveWindowInfo;

        public MainForm()
        {
            _settings = AppSettings.Load();
            _environmentConfiguration = EnvironmentConfiguration.Load();
            _conversationHistory = new List<ConversationTurn>();
            _automationRecipes = AutomationRecipe.LoadAll();
            _watchSuggestions = new List<WatchSuggestion>();
            _microphoneRecorder = new MicrophoneRecorder();

            Text = "Karl Klammer for Windows";
            Width = 1008;
            Height = 928;
            MinimumSize = new Size(980, 900);
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = Color.FromArgb(11, 17, 27);
            ForeColor = Color.White;
            Font = new Font("Segoe UI", 10.0f);
            Icon = SystemIcons.Information;
            DoubleBuffered = true;
            AutoScroll = true;

            _headerPanel = CreateCardPanel(18, 14, 920, 66, Color.FromArgb(18, 28, 44));
            _headerPanel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Controls.Add(_headerPanel);
            _headerPanel.SendToBack();

            _controlPanel = CreateCardPanel(18, 90, 530, 330, Color.FromArgb(16, 23, 36));
            _controlPanel.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            Controls.Add(_controlPanel);
            _controlPanel.SendToBack();

            _contextPanel = CreateCardPanel(560, 90, 378, 468, Color.FromArgb(14, 24, 34));
            _contextPanel.Anchor = AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom;
            Controls.Add(_contextPanel);
            _contextPanel.SendToBack();

            _workspacePanel = CreateCardPanel(18, 548, 920, 308, Color.FromArgb(15, 22, 33));
            _workspacePanel.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            Controls.Add(_workspacePanel);
            _workspacePanel.SendToBack();

            _controlEnvPanel = CreateCardPanel(30, 130, 500, 92, Color.FromArgb(13, 20, 31));
            Controls.Add(_controlEnvPanel);
            _controlEnvPanel.SendToBack();
            _controlModePanel = CreateCardPanel(30, 228, 500, 120, Color.FromArgb(13, 20, 31));
            Controls.Add(_controlModePanel);
            _controlModePanel.SendToBack();
            _controlActionPanel = CreateCardPanel(30, 354, 500, 104, Color.FromArgb(13, 20, 31));
            Controls.Add(_controlActionPanel);
            _controlActionPanel.SendToBack();

            _contextModelPanel = CreateCardPanel(574, 126, 338, 94, Color.FromArgb(16, 26, 39));
            Controls.Add(_contextModelPanel);
            _contextModelPanel.SendToBack();
            _contextRitualPanel = CreateCardPanel(574, 226, 338, 128, Color.FromArgb(16, 26, 39));
            Controls.Add(_contextRitualPanel);
            _contextRitualPanel.SendToBack();
            _contextKnowledgePanel = CreateCardPanel(574, 360, 338, 184, Color.FromArgb(16, 26, 39));
            Controls.Add(_contextKnowledgePanel);
            _contextKnowledgePanel.SendToBack();
            _contextInsightPanel = CreateCardPanel(574, 550, 338, 152, Color.FromArgb(16, 26, 39));
            Controls.Add(_contextInsightPanel);
            _contextInsightPanel.SendToBack();

            _controlSectionLabel = CreateSectionLabel("control deck", 34, 100, 180);
            Controls.Add(_controlSectionLabel);
            _contextSectionLabel = CreateSectionLabel("context + memory", 576, 100, 180);
            Controls.Add(_contextSectionLabel);
            _conversationSectionLabel = CreateSectionLabel("conversation", 34, 558, 180);
            Controls.Add(_conversationSectionLabel);
            _controlEnvSectionLabel = CreateSectionLabel("environment", 42, 136, 160);
            Controls.Add(_controlEnvSectionLabel);
            _controlModeSectionLabel = CreateSectionLabel("voice + mode", 42, 234, 160);
            Controls.Add(_controlModeSectionLabel);
            _controlActionSectionLabel = CreateSectionLabel("system actions", 42, 360, 160);
            Controls.Add(_controlActionSectionLabel);
            _contextModelSectionLabel = CreateSectionLabel("model routing", 586, 132, 160);
            Controls.Add(_contextModelSectionLabel);
            _contextRitualSectionLabel = CreateSectionLabel("ritual engine", 586, 232, 160);
            Controls.Add(_contextRitualSectionLabel);
            _contextKnowledgeSectionLabel = CreateSectionLabel("knowledge manager", 586, 366, 180);
            Controls.Add(_contextKnowledgeSectionLabel);
            _contextInsightSectionLabel = CreateSectionLabel("live context", 586, 556, 160);
            Controls.Add(_contextInsightSectionLabel);

            _screenAwareChip = CreateChip("screen-aware", 620, 24, 92, Color.FromArgb(34, 83, 129), Color.FromArgb(182, 223, 255));
            Controls.Add(_screenAwareChip);
            _ritualMemoryChip = CreateChip("ritual memory", 718, 24, 96, Color.FromArgb(39, 88, 68), Color.FromArgb(188, 239, 214));
            Controls.Add(_ritualMemoryChip);
            _safeActionsChip = CreateChip("safe actions", 820, 24, 92, Color.FromArgb(112, 74, 28), Color.FromArgb(255, 223, 171));
            Controls.Add(_safeActionsChip);
            _modeBadgeLabel = CreateChip("mode: " + _settings.CompanionMode, 344, 24, 112, Color.FromArgb(36, 56, 94), Color.FromArgb(210, 225, 255));
            Controls.Add(_modeBadgeLabel);
            _statusBadgeLabel = CreateChip("status: ready", 462, 24, 112, Color.FromArgb(31, 88, 58), Color.FromArgb(211, 255, 225));
            Controls.Add(_statusBadgeLabel);
            _riskBadgeLabel = CreateChip("risk: low", 580, 24, 80, Color.FromArgb(53, 73, 34), Color.FromArgb(219, 244, 173));
            Controls.Add(_riskBadgeLabel);

            _contextInfoPanel1 = CreateCardPanel(574, 214, 338, 44, Color.FromArgb(18, 31, 46));
            _contextInfoPanel1.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            Controls.Add(_contextInfoPanel1);
            _contextInfoPanel1.SendToBack();
            _contextInfoPanel2 = CreateCardPanel(574, 262, 338, 44, Color.FromArgb(18, 31, 46));
            _contextInfoPanel2.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            Controls.Add(_contextInfoPanel2);
            _contextInfoPanel2.SendToBack();
            _contextInfoPanel3 = CreateCardPanel(574, 310, 338, 44, Color.FromArgb(18, 31, 46));
            _contextInfoPanel3.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            Controls.Add(_contextInfoPanel3);
            _contextInfoPanel3.SendToBack();
            _contextInfoPanel4 = CreateCardPanel(574, 358, 338, 44, Color.FromArgb(25, 37, 28));
            _contextInfoPanel4.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            Controls.Add(_contextInfoPanel4);
            _contextInfoPanel4.SendToBack();
            _activityPanel = CreateCardPanel(598, 712, 318, 96, Color.FromArgb(18, 30, 44));
            _activityPanel.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;
            Controls.Add(_activityPanel);
            _activityPanel.SendToBack();

            _titleLabel = CreateLabel("Karl Klammer for Windows", 24, 20, 320, 28, new Font("Segoe UI Semibold", 16.0f), Color.White);
            Controls.Add(_titleLabel);
            _subtitleLabel = CreateLabel("screen-aware companion + multi-llm + local agents + ritual automation", 26, 50, 720, 24, null, Color.FromArgb(160, 174, 192));
            Controls.Add(_subtitleLabel);
            _envFileLabel = CreateLabel(".env file", 26, 96, 120, 24);
            Controls.Add(_envFileLabel);

            _envPathTextBox = CreateTextBox(26, 122, 520, 28, EnvironmentConfiguration.EnvFilePath);
            _envPathTextBox.ReadOnly = true;
            _envPathTextBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Controls.Add(_envPathTextBox);

            _envStatusLabel = CreateLabel("", 26, 154, 520, 24, null, Color.FromArgb(235, 210, 120));
            _envStatusLabel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Controls.Add(_envStatusLabel);

            _activeAppLabel = CreateLabel("active app: detecting...", 588, 284, 310, 24, new Font("Segoe UI Semibold", 9.5f), Color.White);
            _activeAppLabel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            Controls.Add(_activeAppLabel);

            _appActionsLabel = CreateLabel("app actions: detecting...", 588, 332, 310, 24, null, Color.FromArgb(160, 174, 192));
            _appActionsLabel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            Controls.Add(_appActionsLabel);

            _knowledgeStatusLabel = CreateLabel("knowledge: checking...", 588, 380, 310, 24, null, Color.FromArgb(160, 174, 192));
            _knowledgeStatusLabel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            Controls.Add(_knowledgeStatusLabel);

            _proactiveSuggestionLabel = CreateLabel("next idea: watching your context...", 588, 428, 310, 40, new Font("Segoe UI Semibold", 9.0f), Color.FromArgb(198, 255, 210));
            _proactiveSuggestionLabel.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            Controls.Add(_proactiveSuggestionLabel);

            _providerHealthLabel = CreateLabel("provider: checking...", 40, 792, 240, 20, new Font("Segoe UI", 8.75f), Color.FromArgb(173, 191, 214));
            Controls.Add(_providerHealthLabel);
            _voiceStatusLabel = CreateLabel("voice: checking...", 40, 814, 240, 20, new Font("Segoe UI", 8.75f), Color.FromArgb(173, 191, 214));
            Controls.Add(_voiceStatusLabel);
            _activeWindowDetailsLabel = CreateLabel("window details: pending", 40, 836, 520, 36, new Font("Segoe UI", 8.5f), Color.FromArgb(132, 157, 190));
            Controls.Add(_activeWindowDetailsLabel);
            _recipeManagerLabel = CreateLabel("ritual manager", 0, 0, 140, 20, new Font("Segoe UI Semibold", 9.5f), Color.White);
            Controls.Add(_recipeManagerLabel);
            _actionInspectorLabel = CreateLabel("action inspector", 0, 0, 140, 20, new Font("Segoe UI Semibold", 9.5f), Color.White);
            Controls.Add(_actionInspectorLabel);
            _diagnosticsLabel = CreateLabel("diagnostics", 0, 0, 120, 20, new Font("Segoe UI Semibold", 9.5f), Color.White);
            Controls.Add(_diagnosticsLabel);

            _knowledgeDocsLabel = CreateLabel("knowledge docs", 588, 140, 150, 22, new Font("Segoe UI Semibold", 10.0f), Color.White);
            Controls.Add(_knowledgeDocsLabel);

            _knowledgeSearchTextBox = CreateTextBox(588, 164, 310, 28, string.Empty);
            _knowledgeSearchTextBox.PlaceholderText = "search docs, chunks, filenames";
            _knowledgeSearchTextBox.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _knowledgeSearchTextBox.TextChanged += delegate { RefreshKnowledgeManager(); };
            Controls.Add(_knowledgeSearchTextBox);

            _knowledgeListBox = new ListBox();
            _knowledgeListBox.Location = new Point(588, 198);
            _knowledgeListBox.Size = new Size(310, 78);
            _knowledgeListBox.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _knowledgeListBox.Font = new Font("Segoe UI", 9.0f);
            _knowledgeListBox.BackColor = Color.FromArgb(19, 29, 43);
            _knowledgeListBox.ForeColor = Color.White;
            _knowledgeListBox.BorderStyle = BorderStyle.FixedSingle;
            _knowledgeListBox.SelectedIndexChanged += delegate { UpdateKnowledgePreview(); };
            Controls.Add(_knowledgeListBox);

            _importKnowledgeButton = CreateButton("import docs", 588, 246, 92, 28, Color.FromArgb(22, 30, 45), Color.White);
            _importKnowledgeButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _importKnowledgeButton.Click += delegate { ImportKnowledgeDocuments(); };
            Controls.Add(_importKnowledgeButton);

            _refreshKnowledgeButton = CreateButton("refresh", 686, 246, 74, 28, Color.FromArgb(22, 30, 45), Color.White);
            _refreshKnowledgeButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _refreshKnowledgeButton.Click += delegate { RefreshKnowledgeManager(); };
            Controls.Add(_refreshKnowledgeButton);

            _reindexSelectedKnowledgeButton = CreateButton("reindex", 766, 246, 74, 28, Color.FromArgb(22, 30, 45), Color.White);
            _reindexSelectedKnowledgeButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _reindexSelectedKnowledgeButton.Click += delegate { ReindexSelectedKnowledgeDocument(); };
            Controls.Add(_reindexSelectedKnowledgeButton);

            _removeKnowledgeButton = CreateButton("remove", 846, 246, 74, 28, Color.FromArgb(89, 40, 40), Color.White);
            _removeKnowledgeButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _removeKnowledgeButton.Click += delegate { RemoveSelectedKnowledgeDocument(); };
            Controls.Add(_removeKnowledgeButton);

            _previewLabel = CreateLabel("preview", 588, 476, 120, 18, new Font("Segoe UI Semibold", 9.5f), Color.FromArgb(188, 239, 214));
            Controls.Add(_previewLabel);
            _knowledgePreviewTextBox = CreateTextBox(588, 498, 310, 40, string.Empty);
            _knowledgePreviewTextBox.Anchor = AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom;
            _knowledgePreviewTextBox.Multiline = true;
            _knowledgePreviewTextBox.ReadOnly = true;
            _knowledgePreviewTextBox.ScrollBars = ScrollBars.Vertical;
            Controls.Add(_knowledgePreviewTextBox);

            _recipeManagerListBox = new ListBox();
            _recipeManagerListBox.Location = new Point(0, 0);
            _recipeManagerListBox.Size = new Size(180, 100);
            _recipeManagerListBox.Font = new Font("Segoe UI", 9.0f);
            _recipeManagerListBox.BackColor = Color.FromArgb(19, 29, 43);
            _recipeManagerListBox.ForeColor = Color.White;
            _recipeManagerListBox.BorderStyle = BorderStyle.FixedSingle;
            _recipeManagerListBox.SelectedIndexChanged += delegate { UpdateRecipeManagerPreview(); };
            Controls.Add(_recipeManagerListBox);

            _recipePreviewTextBox = CreateTextBox(0, 0, 180, 100, string.Empty);
            _recipePreviewTextBox.Multiline = true;
            _recipePreviewTextBox.ReadOnly = true;
            _recipePreviewTextBox.ScrollBars = ScrollBars.Vertical;
            Controls.Add(_recipePreviewTextBox);

            _deleteRecipeButton = CreateButton("delete ritual", 0, 0, 100, 28, Color.FromArgb(89, 40, 40), Color.White);
            _deleteRecipeButton.Click += delegate { DeleteSelectedRecipe(); };
            Controls.Add(_deleteRecipeButton);

            _actionInspectorTextBox = CreateTextBox(0, 0, 200, 100, string.Empty);
            _actionInspectorTextBox.Multiline = true;
            _actionInspectorTextBox.ReadOnly = true;
            _actionInspectorTextBox.ScrollBars = ScrollBars.Vertical;
            Controls.Add(_actionInspectorTextBox);

            _diagnosticsListBox = new ListBox();
            _diagnosticsListBox.Location = new Point(0, 0);
            _diagnosticsListBox.Size = new Size(180, 100);
            _diagnosticsListBox.Font = new Font("Consolas", 8.5f);
            _diagnosticsListBox.BackColor = Color.FromArgb(15, 24, 37);
            _diagnosticsListBox.ForeColor = Color.FromArgb(255, 214, 173);
            _diagnosticsListBox.BorderStyle = BorderStyle.FixedSingle;
            Controls.Add(_diagnosticsListBox);

            _providerLabel = CreateLabel("provider", 570, 96, 80, 24);
            Controls.Add(_providerLabel);

            _providerComboBox = new ComboBox();
            _providerComboBox.Location = new Point(570, 122);
            _providerComboBox.Size = new Size(140, 28);
            _providerComboBox.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _providerComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            _providerComboBox.Items.Add("anthropic");
            _providerComboBox.Items.Add("openai");
            _providerComboBox.Items.Add("openai-compatible");
            _providerComboBox.SelectedItem = _providerComboBox.Items.Contains(_settings.AssistantProvider) ? _settings.AssistantProvider : "anthropic";
            _providerComboBox.SelectedIndexChanged += delegate { ApplyProviderDefaultsIfNeeded(); };
            Controls.Add(_providerComboBox);

            _modelLabel = CreateLabel("model", 724, 96, 80, 24);
            Controls.Add(_modelLabel);

            _modelTextBox = CreateTextBox(724, 122, 170, 28, _settings.AssistantModel);
            _modelTextBox.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            Controls.Add(_modelTextBox);

            _speakCheckBox = new CheckBox();
            _speakCheckBox.Text = "speak responses";
            _speakCheckBox.Location = new Point(26, 188);
            _speakCheckBox.Size = new Size(150, 26);
            _speakCheckBox.Checked = _settings.SpeakResponses;
            _speakCheckBox.ForeColor = Color.White;
            _speakCheckBox.BackColor = BackColor;
            Controls.Add(_speakCheckBox);

            _modeLabel = CreateLabel("mode", 192, 162, 80, 24);
            Controls.Add(_modeLabel);

            _modeComboBox = new ComboBox();
            _modeComboBox.Location = new Point(192, 188);
            _modeComboBox.Size = new Size(150, 28);
            _modeComboBox.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            _modeComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            _modeComboBox.Items.Add("companion");
            _modeComboBox.Items.Add("agent");
            _modeComboBox.Items.Add("automation");
            _modeComboBox.Items.Add("watch");
            _modeComboBox.SelectedItem = _modeComboBox.Items.Contains(_settings.CompanionMode) ? _settings.CompanionMode : "companion";
            _modeComboBox.SelectedIndexChanged += delegate
            {
                UpdateModeBadge();
                AddActivityItem("mode", _modeComboBox.SelectedItem == null ? "companion" : _modeComboBox.SelectedItem.ToString());
            };
            Controls.Add(_modeComboBox);

            _suggestAutomationsCheckBox = new CheckBox();
            _suggestAutomationsCheckBox.Text = "save automation hints";
            _suggestAutomationsCheckBox.Location = new Point(360, 188);
            _suggestAutomationsCheckBox.Size = new Size(180, 26);
            _suggestAutomationsCheckBox.Checked = _settings.SuggestAutomations;
            _suggestAutomationsCheckBox.ForeColor = Color.White;
            _suggestAutomationsCheckBox.BackColor = BackColor;
            Controls.Add(_suggestAutomationsCheckBox);

            _useKnowledgeCheckBox = new CheckBox();
            _useKnowledgeCheckBox.Text = "use local knowledge";
            _useKnowledgeCheckBox.Location = new Point(26, 214);
            _useKnowledgeCheckBox.Size = new Size(170, 26);
            _useKnowledgeCheckBox.Checked = _settings.UseLocalKnowledge;
            _useKnowledgeCheckBox.ForeColor = Color.White;
            _useKnowledgeCheckBox.BackColor = BackColor;
            Controls.Add(_useKnowledgeCheckBox);

            _recipeLabel = CreateLabel("recipe", 570, 162, 80, 24);
            Controls.Add(_recipeLabel);

            _recipeComboBox = new ComboBox();
            _recipeComboBox.Location = new Point(570, 188);
            _recipeComboBox.Size = new Size(240, 28);
            _recipeComboBox.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _recipeComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            Controls.Add(_recipeComboBox);

            _learnedRitualsLabel = CreateLabel("learned rituals", 570, 220, 120, 24);
            Controls.Add(_learnedRitualsLabel);

            _watchSuggestionComboBox = new ComboBox();
            _watchSuggestionComboBox.Location = new Point(570, 246);
            _watchSuggestionComboBox.Size = new Size(240, 28);
            _watchSuggestionComboBox.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _watchSuggestionComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            Controls.Add(_watchSuggestionComboBox);

            _statusLabel = CreateLabel("ready", 26, 312, 420, 24, null, Color.FromArgb(93, 212, 136));
            Controls.Add(_statusLabel);

            _hotkeyLabel = CreateLabel("push to talk hotkey: f8", 26, 336, 360, 24, null, Color.FromArgb(160, 174, 192));
            Controls.Add(_hotkeyLabel);

            _saveButton = CreateButton("save settings", 26, 366, 110, 34, Color.FromArgb(33, 150, 243), Color.White);
            _saveButton.Click += delegate
            {
                PersistSettings();
                SetStatus("settings saved", Color.FromArgb(93, 212, 136));
            };
            Controls.Add(_saveButton);

            _reloadEnvButton = CreateButton("reload .env", 150, 366, 110, 34, Color.FromArgb(22, 30, 45), Color.White);
            _reloadEnvButton.Click += delegate
            {
                ReloadEnvironmentConfiguration();
            };
            Controls.Add(_reloadEnvButton);

            _testApiButton = CreateButton("test apis", 274, 366, 110, 34, Color.FromArgb(22, 30, 45), Color.White);
            _testApiButton.Click += async delegate { await RunWorkerTestAsync(); };
            Controls.Add(_testApiButton);

            _clearHistoryButton = CreateButton("clear history", 398, 366, 110, 34, Color.FromArgb(22, 30, 45), Color.White);
            _clearHistoryButton.Click += delegate
            {
                _conversationHistory.Clear();
                _responseTextBox.Clear();
                SetStatus("history cleared", Color.FromArgb(93, 212, 136));
            };
            Controls.Add(_clearHistoryButton);

            _reindexKnowledgeButton = CreateButton("reindex knowledge", 522, 366, 130, 34, Color.FromArgb(22, 30, 45), Color.White);
            _reindexKnowledgeButton.Click += delegate
            {
                try
                {
                    int chunkCount = KnowledgeBaseService.Reindex();
                    RefreshKnowledgeStatus();
                    RefreshKnowledgeManager();
                    SetStatus("knowledge indexed: " + chunkCount + " chunks", Color.FromArgb(93, 212, 136));
                }
                catch (Exception exception)
                {
                    SetStatus("knowledge indexing failed", Color.FromArgb(235, 120, 120));
                    MessageBox.Show(exception.Message, "Karl Klammer knowledge indexing", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            Controls.Add(_reindexKnowledgeButton);

            _saveRecipeButton = CreateButton("save recipe", 666, 366, 110, 34, Color.FromArgb(22, 30, 45), Color.White);
            _saveRecipeButton.Click += delegate { SaveCurrentPromptAsRecipe(); };
            Controls.Add(_saveRecipeButton);

            _runRecipeButton = CreateButton("run recipe", 790, 366, 110, 34, Color.FromArgb(22, 30, 45), Color.White);
            _runRecipeButton.Click += async delegate { await RunSelectedRecipeAsync(); };
            Controls.Add(_runRecipeButton);

            _useContextIdeaButton = CreateButton("use context idea", 26, 402, 150, 34, Color.FromArgb(22, 30, 45), Color.White);
            _useContextIdeaButton.Click += delegate { ApplyProactiveSuggestionToPrompt(); };
            Controls.Add(_useContextIdeaButton);

            _runContextIdeaButton = CreateButton("run context idea", 190, 402, 150, 34, Color.FromArgb(22, 30, 45), Color.White);
            _runContextIdeaButton.Click += async delegate { await RunProactiveSuggestionAsync(); };
            Controls.Add(_runContextIdeaButton);

            _openRitualManagerButton = CreateButton("ritual manager", 0, 0, 110, 30, Color.FromArgb(22, 30, 45), Color.White);
            _openRitualManagerButton.Click += delegate { ShowRitualManagerDialog(); };
            Controls.Add(_openRitualManagerButton);

            _openHistoryViewerButton = CreateButton("action history", 0, 0, 110, 30, Color.FromArgb(22, 30, 45), Color.White);
            _openHistoryViewerButton.Click += delegate { ShowActionHistoryDialog(); };
            Controls.Add(_openHistoryViewerButton);

            _openControlInspectorButton = CreateButton("window inspector", 0, 0, 110, 30, Color.FromArgb(22, 30, 45), Color.White);
            _openControlInspectorButton.Click += delegate { ShowControlInspectorDialog(); };
            Controls.Add(_openControlInspectorButton);

            _openProviderVoiceButton = CreateButton("provider + voice", 0, 0, 110, 30, Color.FromArgb(22, 30, 45), Color.White);
            _openProviderVoiceButton.Click += delegate { ShowProviderVoiceDialog(); };
            Controls.Add(_openProviderVoiceButton);

            _openDiagnosticsButton = CreateButton("diagnostics", 0, 0, 110, 30, Color.FromArgb(22, 30, 45), Color.White);
            _openDiagnosticsButton.Click += delegate { ShowDiagnosticsDialog(); };
            Controls.Add(_openDiagnosticsButton);

            _openSetupWizardButton = CreateButton("setup wizard", 0, 0, 110, 30, Color.FromArgb(39, 88, 68), Color.White);
            _openSetupWizardButton.Click += delegate { ShowSetupWizardDialog(); };
            Controls.Add(_openSetupWizardButton);

            _saveWatchIdeaButton = CreateButton("save ritual", 570, 280, 160, 34, Color.FromArgb(22, 30, 45), Color.White);
            _saveWatchIdeaButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _saveWatchIdeaButton.Click += delegate { SaveSelectedWatchSuggestionAsRecipe(); };
            Controls.Add(_saveWatchIdeaButton);

            _useWatchIdeaButton = CreateButton("use ritual", 744, 280, 110, 34, Color.FromArgb(22, 30, 45), Color.White);
            _useWatchIdeaButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _useWatchIdeaButton.Click += delegate { UseSelectedWatchSuggestion(); };
            Controls.Add(_useWatchIdeaButton);

            _replayRitualButton = CreateButton("replay", 860, 280, 66, 34, Color.FromArgb(22, 30, 45), Color.White);
            _replayRitualButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _replayRitualButton.Click += async delegate { await ReplaySelectedWatchSuggestionAsync(); };
            Controls.Add(_replayRitualButton);

            _promptLabel = CreateLabel("what do you need help with?", 40, 566, 260, 24, new Font("Segoe UI Semibold", 10.5f), Color.White);
            Controls.Add(_promptLabel);

            _promptTextBox = CreateTextBox(40, 596, 876, 104, string.Empty);
            _promptTextBox.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            _promptTextBox.Multiline = true;
            _promptTextBox.AcceptsReturn = true;
            _promptTextBox.ScrollBars = ScrollBars.Vertical;
            Controls.Add(_promptTextBox);

            _askButton = CreateButton("ask about my screen", 40, 712, 200, 42, Color.FromArgb(93, 212, 136), Color.Black);
            _askButton.Anchor = AnchorStyles.Left | AnchorStyles.Bottom;
            _askButton.Click += async delegate { await RunAskFlowAsync(); };
            Controls.Add(_askButton);

            _dictationButton = CreateButton("hold to talk", 254, 712, 160, 42, Color.FromArgb(22, 30, 45), Color.White);
            _dictationButton.Anchor = AnchorStyles.Left | AnchorStyles.Bottom;
            _dictationButton.MouseDown += delegate(object sender, MouseEventArgs e)
            {
                if (e.Button == MouseButtons.Left)
                {
                    HandleSpeechPress();
                }
            };
            _dictationButton.MouseUp += async delegate(object sender, MouseEventArgs e)
            {
                if (e.Button == MouseButtons.Left)
                {
                    await HandleSpeechReleaseAsync();
                }
            };
            Controls.Add(_dictationButton);
            UpdateSpeechButtonState();

            _copyResponseButton = CreateButton("copy response", 428, 712, 150, 42, Color.FromArgb(22, 30, 45), Color.White);
            _copyResponseButton.Anchor = AnchorStyles.Left | AnchorStyles.Bottom;
            _copyResponseButton.Click += delegate
            {
                if (!string.IsNullOrWhiteSpace(_responseTextBox.Text))
                {
                    Clipboard.SetText(_responseTextBox.Text);
                    SetStatus("response copied", Color.FromArgb(93, 212, 136));
                }
            };
            Controls.Add(_copyResponseButton);

            _activityRailLabel = CreateLabel("activity rail", 614, 720, 140, 20, new Font("Segoe UI Semibold", 9.5f), Color.White);
            Controls.Add(_activityRailLabel);
            _activityListBox = new ListBox();
            _activityListBox.Location = new Point(614, 744);
            _activityListBox.Size = new Size(286, 52);
            _activityListBox.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;
            _activityListBox.Font = new Font("Consolas", 8.5f);
            _activityListBox.BackColor = Color.FromArgb(15, 24, 37);
            _activityListBox.ForeColor = Color.FromArgb(188, 239, 214);
            _activityListBox.BorderStyle = BorderStyle.None;
            _activityListBox.IntegralHeight = false;
            Controls.Add(_activityListBox);

            _responseLabel = CreateLabel("response", 40, 790, 160, 24, new Font("Segoe UI Semibold", 10.5f), Color.White);
            Controls.Add(_responseLabel);

            _responseTextBox = CreateTextBox(40, 818, 876, 38, string.Empty);
            _responseTextBox.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            _responseTextBox.Multiline = true;
            _responseTextBox.ReadOnly = true;
            _responseTextBox.ScrollBars = ScrollBars.Vertical;
            Controls.Add(_responseTextBox);

            _retrievalSourcesLabel = CreateLabel("sources: none yet", 40, 760, 876, 24, new Font("Segoe UI", 8.5f), Color.FromArgb(116, 151, 188));
            _retrievalSourcesLabel.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            Controls.Add(_retrievalSourcesLabel);

            _notifyIcon = new NotifyIcon();
            _notifyIcon.Icon = SystemIcons.Information;
            _notifyIcon.Text = "Karl Klammer for Windows";
            _notifyIcon.Visible = true;
            _notifyIcon.DoubleClick += delegate { ShowClickyWindow(); };

            _trayMenu = new ContextMenuStrip();
            ToolStripItem openMenuItem = _trayMenu.Items.Add("Open Karl Klammer");
            ToolStripItem askMenuItem = _trayMenu.Items.Add("Ask About Screen");
            ToolStripItem runRecipeMenuItem = _trayMenu.Items.Add("Run Selected Recipe");
            ToolStripItem testMenuItem = _trayMenu.Items.Add("Test APIs");
            _trayMenu.Items.Add("-");
            ToolStripItem quitMenuItem = _trayMenu.Items.Add("Quit");

            openMenuItem.Click += delegate { ShowClickyWindow(); };
            askMenuItem.Click += async delegate
            {
                ShowClickyWindow();
                await RunAskFlowAsync();
            };
            runRecipeMenuItem.Click += async delegate
            {
                ShowClickyWindow();
                await RunSelectedRecipeAsync();
            };
            testMenuItem.Click += async delegate { await RunWorkerTestAsync(); };
            quitMenuItem.Click += delegate
            {
                _quitRequested = true;
                Close();
            };
            _notifyIcon.ContextMenuStrip = _trayMenu;

            FormClosing += OnMainFormClosing;
            Resize += delegate { ApplyAdaptiveLayout(); };
            Shown += delegate
            {
                EnsureCompanionOverlay();
                ReloadEnvironmentConfiguration();
                RefreshActiveAppLabel();
                RefreshKnowledgeStatus();
                RefreshKnowledgeManager();
                RefreshRecipeList(null);
                RefreshWatchSuggestions(null);
                RefreshProactiveSuggestion();
                UpdateProviderVoiceSummary();
                UpdateActionInspector("ready", "Karl Klammer is ready.\r\nUse ask, recipes, rituals or app actions.");
                UpdateModeBadge();
                AddActivityItem("boot", "companion deck ready");
                SetCompanionState(CompanionVisualState.Idle);

                if (!string.IsNullOrWhiteSpace(_environmentConfiguration.Validate(_settings.AssistantProvider)))
                {
                    SetStatus("create .env first, then reload it", Color.FromArgb(235, 210, 120));
                }
                else
                {
                    _promptTextBox.Focus();
                }
            };

            BringInteractiveControlsToFront();
            ApplyAdaptiveLayout();

            _contextRefreshTimer = new Timer();
            _contextRefreshTimer.Interval = 2500;
            _contextRefreshTimer.Tick += delegate
            {
                RefreshActiveAppLabel();
                RefreshProactiveSuggestion();
            };
            _contextRefreshTimer.Start();
        }

        private async Task RunWorkerTestAsync()
        {
            try
            {
                PersistSettings();
                EnsureEnvironmentIsReady();
                SetCompanionState(CompanionVisualState.Thinking);
                SetStatus("testing apis...", Color.FromArgb(235, 210, 120));
                string responseText = await DirectApiClient.SmokeTestAsync(_settings, _environmentConfiguration);
                _responseTextBox.Text = responseText;
                _retrievalSourcesLabel.Text = "sources: none";
                UpdateActionInspector("api smoke test", responseText);
                SetStatus("apis ready", Color.FromArgb(93, 212, 136));
                ShowCompanionMessage(responseText, CompanionVisualState.Idle, 2600, CompanionVisualState.Idle);
            }
            catch (Exception exception)
            {
                UpdateActionInspector("api smoke test failed", exception.Message);
                SetStatus("api test failed", Color.FromArgb(235, 120, 120));
                ShowCompanionMessage("api test failed", CompanionVisualState.Idle, 3200, CompanionVisualState.Idle);
                MessageBox.Show(exception.Message, "Karl Klammer api test", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async Task RunAskFlowAsync(string promptOverride = null)
        {
            string prompt = string.IsNullOrWhiteSpace(promptOverride) ? _promptTextBox.Text.Trim() : promptOverride.Trim();
            if (string.IsNullOrWhiteSpace(prompt))
            {
                MessageBox.Show("Type a question first.", "Karl Klammer", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                PersistSettings();
                ReloadEnvironmentConfiguration();
                RefreshActiveAppLabel();
                UpdateModeBadge();
                AddActivityItem("ask", prompt);
                SetCompanionState(CompanionVisualState.Thinking);
                ActiveWindowInfo activeWindow = ActiveWindowService.GetActiveWindowInfo();

                if (OpenClawClient.IsTriggered(prompt))
                {
                    await RunOpenClawFlowAsync(prompt);
                    return;
                }

                if (ClaudeCodeClient.IsTriggered(prompt))
                {
                    await RunClaudeCodeFlowAsync(prompt);
                    return;
                }

                if (CodexClient.IsTriggered(prompt))
                {
                    await RunCodexFlowAsync(prompt);
                    return;
                }

                string autoRoute = IntentRouter.DetectRoute(prompt, activeWindow);
                if (string.Equals(autoRoute, "codex", StringComparison.OrdinalIgnoreCase))
                {
                    SetStatus("auto-routing to codex from current context...", Color.FromArgb(235, 210, 120));
                    await RunCodexFlowAsync("nimm codex " + prompt);
                    return;
                }

                if (string.Equals(autoRoute, "openclaw", StringComparison.OrdinalIgnoreCase))
                {
                    SetStatus("auto-routing to openclaw from current context...", Color.FromArgb(235, 210, 120));
                    await RunOpenClawFlowAsync("nimm openclaw " + prompt);
                    return;
                }

                EnsureEnvironmentIsReady();

                SetStatus("capturing screens...", Color.FromArgb(235, 210, 120));
                List<ScreenCaptureInfo> screenCaptures = ScreenCaptureService.CaptureAllScreens();

                List<KnowledgeChunk> knowledgeChunks = _settings.UseLocalKnowledge
                    ? KnowledgeBaseService.Retrieve(prompt, 3)
                    : new List<KnowledgeChunk>();

                SetStatus(
                    knowledgeChunks.Count > 0 ? "asking Karl Klammer with local knowledge..." : "asking Karl Klammer...",
                    Color.FromArgb(235, 210, 120));
                string fullResponseText = await DirectApiClient.AskAsync(_settings, _environmentConfiguration, prompt, screenCaptures, _conversationHistory, knowledgeChunks);
                AutomationSuggestionResult automationResult = AutomationSuggestionResult.Parse(fullResponseText, _settings.CompanionMode);
                ActionPlanResult actionPlanResult = ActionPlanResult.Parse(automationResult.CleanText);
                ActionTagResult actionResult = ActionTagResult.Parse(actionPlanResult.CleanText);
                int addedRecipeCount = MergeAutomationSuggestions(automationResult.Recipes);
                PointTagResult pointTag = PointTagResult.Parse(actionResult.CleanText);
                string spokenText = string.IsNullOrWhiteSpace(pointTag.SpokenText) ? actionResult.CleanText.Trim() : pointTag.SpokenText;
                UpdateActionInspector(
                    "ask result",
                    "provider: " + _settings.AssistantProvider + Environment.NewLine
                    + "mode: " + _settings.CompanionMode + Environment.NewLine
                    + "knowledge chunks: " + knowledgeChunks.Count + Environment.NewLine
                    + "action: " + (actionResult == null ? "none" : actionResult.ActionName) + Environment.NewLine
                    + "plan steps: " + (actionPlanResult == null || actionPlanResult.Steps == null ? 0 : actionPlanResult.Steps.Count));

                _responseTextBox.Text = spokenText;
                AddActivityItem("reply", spokenText);
                UpdateRetrievalSourcesLabel(knowledgeChunks);
                AddConversationTurn(prompt, spokenText);

                Point? screenPoint = ConvertPointTagToScreenPoint(pointTag, screenCaptures);
                if (screenPoint.HasValue)
                {
                    string bubbleText = !string.IsNullOrWhiteSpace(pointTag.ElementLabel) ? pointTag.ElementLabel : spokenText;
                    NavigateCompanionTo(
                        screenPoint.Value,
                        bubbleText,
                        _settings.SpeakResponses ? CompanionVisualState.Speaking : CompanionVisualState.Idle,
                        _settings.SpeakResponses ? 9000 : 7200,
                        CompanionVisualState.Idle
                    );
                }

                LogWatchSessionIfNeeded(prompt, spokenText, screenCaptures);
                await ExecuteActionPlanIfConfirmedAsync(actionPlanResult, pointTag, screenCaptures, spokenText);
                await ExecuteActionIfConfirmedAsync(actionResult, pointTag, screenCaptures, spokenText);

                SetStatus(addedRecipeCount > 0 ? "done + saved " + addedRecipeCount + " recipe(s)" : "done", Color.FromArgb(93, 212, 136));
                if (!screenPoint.HasValue)
                {
                    ShowCompanionMessage(
                        spokenText,
                        _settings.SpeakResponses ? CompanionVisualState.Speaking : CompanionVisualState.Idle,
                        _settings.SpeakResponses ? 9000 : 7200,
                        CompanionVisualState.Idle
                    );
                }
                await SpeakResponseAsync(spokenText);
            }
            catch (Exception exception)
            {
                UpdateActionInspector("ask failed", exception.Message);
                SetStatus("error", Color.FromArgb(235, 120, 120));
                ShowCompanionMessage("request failed", CompanionVisualState.Idle, 3400, CompanionVisualState.Idle);
                MessageBox.Show(exception.Message, "Karl Klammer request failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private int MergeAutomationSuggestions(IList<AutomationRecipe> recipes)
        {
            if (recipes == null || recipes.Count == 0 || !_settings.SuggestAutomations)
            {
                return 0;
            }

            int addedCount = 0;
            foreach (AutomationRecipe recipe in recipes)
            {
                if (recipe == null || string.IsNullOrWhiteSpace(recipe.Name) || string.IsNullOrWhiteSpace(recipe.Prompt))
                {
                    continue;
                }

                bool exists = _automationRecipes.Any(existing =>
                    string.Equals(existing.Name, recipe.Name, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(existing.Prompt, recipe.Prompt, StringComparison.OrdinalIgnoreCase));
                if (exists)
                {
                    continue;
                }

                _automationRecipes.Add(recipe);
                addedCount++;
            }

            if (addedCount > 0)
            {
                AutomationRecipe.SaveAll(_automationRecipes);
                RefreshRecipeList(null);
            }

            return addedCount;
        }

        private void LogWatchSessionIfNeeded(string prompt, string responseText, IList<ScreenCaptureInfo> screenCaptures)
        {
            if (!string.Equals(_settings.CompanionMode, "watch", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            ActiveWindowInfo activeWindow = ActiveWindowService.GetActiveWindowInfo();
            string screenSummary = string.Join(" | ", screenCaptures.Select(capture => capture.Label).ToArray());
            WatchSessionEntry.Append(new WatchSessionEntry
            {
                TimestampUtc = DateTime.UtcNow.ToString("o"),
                Prompt = prompt,
                AssistantResponse = responseText,
                Provider = _settings.AssistantProvider,
                Model = _settings.AssistantModel,
                ScreenSummary = screenSummary,
                ActiveApp = activeWindow.DisplayName
            });
            RefreshWatchSuggestions(null);
        }

        private async Task ExecuteActionIfConfirmedAsync(ActionTagResult actionResult, PointTagResult pointTag, IList<ScreenCaptureInfo> screenCaptures, string spokenText)
        {
            if (actionResult == null || string.Equals(actionResult.ActionName, "none", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            string targetLabel = string.IsNullOrWhiteSpace(pointTag.ElementLabel) ? "this target" : pointTag.ElementLabel;
            string confirmationText = BuildActionConfirmationText(actionResult, targetLabel, spokenText);
            ActionRiskProfile riskProfile = ActionRiskAssessor.Assess(actionResult, ActiveWindowService.GetActiveWindowInfo());
            UpdateRiskBadge(riskProfile.Level);
            DialogResult confirmation = MessageBox.Show(
                confirmationText,
                "Karl Klammer action confirmation",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (confirmation != DialogResult.Yes)
            {
                UpdateActionInspector("action cancelled", riskProfile.Summary);
                return;
            }

            if (string.Equals(actionResult.ActionName, "open", StringComparison.OrdinalIgnoreCase))
            {
                ExecuteOpenAction(actionResult.ActionArgument);
                LogActionExecution(actionResult, targetLabel, spokenText);
                UpdateActionInspector("action executed: open", actionResult.ActionArgument);
                SetStatus("action executed: open", Color.FromArgb(93, 212, 136));
                return;
            }

            if (string.Equals(actionResult.ActionName, "type", StringComparison.OrdinalIgnoreCase))
            {
                await Task.Delay(120).ConfigureAwait(true);
                SendKeys.SendWait(actionResult.ActionArgument ?? string.Empty);
                LogActionExecution(actionResult, targetLabel, spokenText);
                UpdateActionInspector("action executed: type", actionResult.ActionArgument);
                SetStatus("action executed: type", Color.FromArgb(93, 212, 136));
                return;
            }

            if (string.Equals(actionResult.ActionName, "hotkey", StringComparison.OrdinalIgnoreCase))
            {
                await Task.Delay(120).ConfigureAwait(true);
                SendKeys.SendWait(actionResult.ActionArgument ?? string.Empty);
                LogActionExecution(actionResult, targetLabel, spokenText);
                UpdateActionInspector("action executed: hotkey", actionResult.ActionArgument);
                SetStatus("action executed: hotkey", Color.FromArgb(93, 212, 136));
                return;
            }

            if (string.Equals(actionResult.ActionName, "app", StringComparison.OrdinalIgnoreCase))
            {
                await Task.Delay(120).ConfigureAwait(true);
                string appActionResult = AppActionAdapter.Execute(ActiveWindowService.GetActiveWindowInfo(), actionResult.ActionArgument);
                LogActionExecution(actionResult, targetLabel, spokenText);
                _responseTextBox.Text = appActionResult;
                UpdateActionInspector("action executed: app action", appActionResult);
                SetStatus("action executed: app action", Color.FromArgb(93, 212, 136));
                return;
            }

            Point? screenPoint = ConvertPointTagToScreenPoint(pointTag, screenCaptures);
            if (!screenPoint.HasValue)
            {
                screenPoint = Cursor.Position;
            }

            await Task.Delay(120).ConfigureAwait(true);
            SetCursorPos(screenPoint.Value.X, screenPoint.Value.Y);
            NavigateCompanionTo(screenPoint.Value, targetLabel, CompanionVisualState.Speaking, 3600, CompanionVisualState.Idle);

            if (string.Equals(actionResult.ActionName, "click", StringComparison.OrdinalIgnoreCase))
            {
                await Task.Delay(120).ConfigureAwait(true);
                mouse_event(MouseeventfLeftdown, 0, 0, 0, UIntPtr.Zero);
                mouse_event(MouseeventfLeftup, 0, 0, 0, UIntPtr.Zero);
                LogActionExecution(actionResult, targetLabel, spokenText);
                UpdateActionInspector("action executed: click", targetLabel);
                SetStatus("action executed: click", Color.FromArgb(93, 212, 136));
            }
            else
            {
                LogActionExecution(actionResult, targetLabel, spokenText);
                UpdateActionInspector("action executed: move", targetLabel);
                SetStatus("action executed: move", Color.FromArgb(93, 212, 136));
            }
        }

        private async Task ExecuteActionPlanIfConfirmedAsync(ActionPlanResult actionPlanResult, PointTagResult pointTag, IList<ScreenCaptureInfo> screenCaptures, string spokenText)
        {
            if (actionPlanResult == null || actionPlanResult.Steps == null || actionPlanResult.Steps.Count == 0)
            {
                return;
            }

            string targetLabel = string.IsNullOrWhiteSpace(pointTag.ElementLabel) ? "this target" : pointTag.ElementLabel;
            ActiveWindowInfo activeWindow = ActiveWindowService.GetActiveWindowInfo();
            ActionRiskProfile riskProfile = ActionRiskAssessor.AssessPlan(actionPlanResult.Steps, activeWindow);
            UpdateRiskBadge(riskProfile.Level);
            string summary = string.Join(" -> ", actionPlanResult.Steps.Select(step =>
                string.IsNullOrWhiteSpace(step.ActionArgument) ? step.ActionName : step.ActionName + "(" + step.ActionArgument + ")").ToArray());

            DialogResult confirmation = MessageBox.Show(
                string.Format("Karl Klammer wants to run this action chain:\r\n{0}\r\n\r\nRisk: {1} - {2}\r\nApp: {3}\r\nTarget: {4}\r\n\r\n{5}",
                    summary,
                    riskProfile.Level,
                    riskProfile.Summary,
                    activeWindow.DisplayName,
                    targetLabel,
                    spokenText),
                "Karl Klammer action chain confirmation",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (confirmation != DialogResult.Yes)
            {
                UpdateActionInspector("action chain cancelled", riskProfile.Summary);
                return;
            }

            foreach (ActionPlanStep step in actionPlanResult.Steps)
            {
                await ExecuteSingleActionStepAsync(step, pointTag, screenCaptures, spokenText).ConfigureAwait(true);
                await Task.Delay(140).ConfigureAwait(true);
            }

            UpdateActionInspector("action chain executed", summary);
            SetStatus("action chain executed", Color.FromArgb(93, 212, 136));
        }

        private async Task ExecuteSingleActionStepAsync(ActionPlanStep step, PointTagResult pointTag, IList<ScreenCaptureInfo> screenCaptures, string spokenText)
        {
            if (step == null || string.IsNullOrWhiteSpace(step.ActionName))
            {
                return;
            }

            if (step.WaitMilliseconds > 0)
            {
                await Task.Delay(step.WaitMilliseconds).ConfigureAwait(true);
            }

            if (!string.IsNullOrWhiteSpace(step.RequiredAppContains))
            {
                ActiveWindowInfo currentWindow = ActiveWindowService.GetActiveWindowInfo();
                string currentDisplay = currentWindow == null ? string.Empty : currentWindow.DisplayName ?? string.Empty;
                string currentProcess = currentWindow == null ? string.Empty : currentWindow.ProcessName ?? string.Empty;
                if (currentDisplay.IndexOf(step.RequiredAppContains, StringComparison.OrdinalIgnoreCase) < 0
                    && currentProcess.IndexOf(step.RequiredAppContains, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    AddDiagnosticItem("guard-skip", "Skipped step '" + step.ActionName + "' because app did not match '" + step.RequiredAppContains + "'.");
                    return;
                }
            }

            string actionName = step.ActionName.Trim().ToLowerInvariant();
            string targetLabel = string.IsNullOrWhiteSpace(pointTag.ElementLabel) ? "this target" : pointTag.ElementLabel;

            if (actionName == "open")
            {
                ExecuteOpenAction(step.ActionArgument);
                LogActionExecution(new ActionTagResult { ActionName = actionName, ActionArgument = step.ActionArgument }, targetLabel, spokenText);
                return;
            }

            if (actionName == "type" || actionName == "hotkey")
            {
                SendKeys.SendWait(step.ActionArgument ?? string.Empty);
                LogActionExecution(new ActionTagResult { ActionName = actionName, ActionArgument = step.ActionArgument }, targetLabel, spokenText);
                return;
            }

            if (actionName == "app")
            {
                string appActionResult = AppActionAdapter.Execute(ActiveWindowService.GetActiveWindowInfo(), step.ActionArgument);
                if (!string.IsNullOrWhiteSpace(appActionResult))
                {
                    _responseTextBox.Text = appActionResult;
                }
                LogActionExecution(new ActionTagResult { ActionName = actionName, ActionArgument = step.ActionArgument }, targetLabel, spokenText);
                return;
            }

            Point? screenPoint = ConvertPointTagToScreenPoint(pointTag, screenCaptures);
            if (!screenPoint.HasValue)
            {
                screenPoint = Cursor.Position;
            }

            SetCursorPos(screenPoint.Value.X, screenPoint.Value.Y);
            NavigateCompanionTo(screenPoint.Value, targetLabel, CompanionVisualState.Speaking, 2200, CompanionVisualState.Idle);

            if (actionName == "click")
            {
                await Task.Delay(120).ConfigureAwait(true);
                mouse_event(MouseeventfLeftdown, 0, 0, 0, UIntPtr.Zero);
                mouse_event(MouseeventfLeftup, 0, 0, 0, UIntPtr.Zero);
            }

            LogActionExecution(new ActionTagResult { ActionName = actionName, ActionArgument = step.ActionArgument }, targetLabel, spokenText);
        }

        private static string BuildActionConfirmationText(ActionTagResult actionResult, string targetLabel, string spokenText)
        {
            string actionName = actionResult.ActionName ?? "none";
            ActiveWindowInfo activeWindow = ActiveWindowService.GetActiveWindowInfo();
            ActionRiskProfile riskProfile = ActionRiskAssessor.Assess(actionResult, activeWindow);
            switch (actionName)
            {
                case "click":
                    return string.Format("Karl Klammer wants to click \"{0}\".\r\nRisk: {1} - {2}\r\nApp: {3}\r\n\r\n{4}", targetLabel, riskProfile.Level, riskProfile.Summary, activeWindow.DisplayName, spokenText);
                case "move":
                    return string.Format("Karl Klammer wants to move the cursor to \"{0}\".\r\nRisk: {1} - {2}\r\nApp: {3}\r\n\r\n{4}", targetLabel, riskProfile.Level, riskProfile.Summary, activeWindow.DisplayName, spokenText);
                case "open":
                    return string.Format("Karl Klammer wants to open:\r\n{0}\r\nRisk: {1} - {2}\r\nApp: {3}\r\n\r\n{4}", actionResult.ActionArgument, riskProfile.Level, riskProfile.Summary, activeWindow.DisplayName, spokenText);
                case "type":
                    return string.Format("Karl Klammer wants to type:\r\n{0}\r\nRisk: {1} - {2}\r\nApp: {3}\r\n\r\n{4}", actionResult.ActionArgument, riskProfile.Level, riskProfile.Summary, activeWindow.DisplayName, spokenText);
                case "hotkey":
                    return string.Format("Karl Klammer wants to send this shortcut:\r\n{0}\r\nRisk: {1} - {2}\r\nApp: {3}\r\n\r\n{4}", actionResult.ActionArgument, riskProfile.Level, riskProfile.Summary, activeWindow.DisplayName, spokenText);
                case "app":
                    return string.Format(
                        "Karl Klammer wants to run this app action in:\r\n{0}\r\nAction: {1}\r\nRisk: {2} - {3}\r\n\r\n{4}",
                        activeWindow.DisplayName,
                        AppActionAdapter.Describe(activeWindow, actionResult.ActionArgument),
                        riskProfile.Level,
                        riskProfile.Summary,
                        spokenText);
                default:
                    return spokenText;
            }
        }

        private static void ExecuteOpenAction(string actionArgument)
        {
            if (string.IsNullOrWhiteSpace(actionArgument))
            {
                throw new InvalidOperationException("The open action did not contain a target.");
            }

            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = actionArgument,
                UseShellExecute = true
            };
            Process.Start(startInfo);
        }

        private static void LogActionExecution(ActionTagResult actionResult, string targetLabel, string spokenText)
        {
            ActiveWindowInfo activeWindow = ActiveWindowService.GetActiveWindowInfo();
            ActionHistoryEntry.Append(new ActionHistoryEntry
            {
                TimestampUtc = DateTime.UtcNow.ToString("o"),
                ActionName = actionResult == null ? string.Empty : actionResult.ActionName,
                ActionArgument = actionResult == null ? string.Empty : actionResult.ActionArgument,
                TargetLabel = targetLabel ?? string.Empty,
                SpokenText = spokenText ?? string.Empty,
                ActiveApp = activeWindow.DisplayName
            });
        }

        private async Task RunCodexFlowAsync(string prompt)
        {
            List<string> temporaryImagePaths = null;
            string codexPrompt = CodexClient.RemoveTrigger(prompt);
            if (string.IsNullOrWhiteSpace(codexPrompt))
            {
                throw new InvalidOperationException("Say what Codex should do after 'nimm codex'.");
            }

            _promptTextBox.Text = codexPrompt;
            _promptTextBox.SelectionStart = _promptTextBox.TextLength;
            _promptTextBox.ScrollToCaret();

            if (CodexClient.ShouldAttachScreens(prompt))
            {
                SetStatus("capturing screens for codex...", Color.FromArgb(235, 210, 120));
                List<ScreenCaptureInfo> screenCaptures = ScreenCaptureService.CaptureAllScreens();
                temporaryImagePaths = SaveCodexScreenCaptures(screenCaptures);
            }

            SetStatus(temporaryImagePaths != null && temporaryImagePaths.Count > 0 ? "starting codex with screens..." : "starting codex...", Color.FromArgb(235, 210, 120));
            CodexRunResult result = await CodexClient.RunAsync(_environmentConfiguration, codexPrompt, temporaryImagePaths);

            string completionMessage = CodexClient.GetCompletionMessage();
            _responseTextBox.Text = completionMessage;
            SetStatus("codex done, saved to codex output", Color.FromArgb(93, 212, 136));
            ShowCompanionMessage(
                completionMessage,
                _settings.SpeakResponses ? CompanionVisualState.Speaking : CompanionVisualState.Idle,
                _settings.SpeakResponses ? 9000 : 7200,
                CompanionVisualState.Idle
            );
            await SpeakResponseAsync(completionMessage);
            DeleteFilesQuietly(temporaryImagePaths);
        }

        private async Task RunClaudeCodeFlowAsync(string prompt)
        {
            string claudeCodePrompt = ClaudeCodeClient.RemoveTrigger(prompt);
            if (string.IsNullOrWhiteSpace(claudeCodePrompt))
            {
                throw new InvalidOperationException("Say what Claude Code should do after 'nimm claude code'.");
            }

            _promptTextBox.Text = claudeCodePrompt;
            _promptTextBox.SelectionStart = _promptTextBox.TextLength;
            _promptTextBox.ScrollToCaret();

            SetStatus("starting claude code...", Color.FromArgb(235, 210, 120));
            CodexRunResult result = await ClaudeCodeClient.RunAsync(_environmentConfiguration, claudeCodePrompt);

            string completionMessage = ClaudeCodeClient.GetCompletionMessage();
            _responseTextBox.Text = completionMessage;
            SetStatus("claude code done, saved to codex output", Color.FromArgb(93, 212, 136));
            ShowCompanionMessage(
                completionMessage,
                _settings.SpeakResponses ? CompanionVisualState.Speaking : CompanionVisualState.Idle,
                _settings.SpeakResponses ? 9000 : 7200,
                CompanionVisualState.Idle
            );
            await SpeakResponseAsync(completionMessage);
        }

        private async Task RunOpenClawFlowAsync(string prompt)
        {
            string openClawPrompt = OpenClawClient.RemoveTrigger(prompt);
            if (string.IsNullOrWhiteSpace(openClawPrompt))
            {
                throw new InvalidOperationException("Say what OpenClaw should do after 'nimm openclaw'.");
            }

            _promptTextBox.Text = openClawPrompt;
            _promptTextBox.SelectionStart = _promptTextBox.TextLength;
            _promptTextBox.ScrollToCaret();

            SetStatus("starting openclaw...", Color.FromArgb(235, 210, 120));
            OpenClawRunResult result = await OpenClawClient.RunAsync(_environmentConfiguration, openClawPrompt);

            string responseText = string.IsNullOrWhiteSpace(result.ResponseText)
                ? OpenClawClient.GetCompletionMessage()
                : result.ResponseText.Trim();
            _responseTextBox.Text = responseText;
            AddConversationTurn(openClawPrompt, responseText);
            SetStatus("openclaw done, saved to codex output", Color.FromArgb(93, 212, 136));
            ShowCompanionMessage(
                responseText,
                _settings.SpeakResponses ? CompanionVisualState.Speaking : CompanionVisualState.Idle,
                _settings.SpeakResponses ? 9000 : 7200,
                CompanionVisualState.Idle
            );
            await SpeakResponseAsync(responseText);
        }

        private static List<string> SaveCodexScreenCaptures(IList<ScreenCaptureInfo> screenCaptures)
        {
            string screenCaptureDirectory = Path.Combine(Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..")), "codex output", "screen captures");
            Directory.CreateDirectory(screenCaptureDirectory);

            string timestamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
            List<string> imagePaths = new List<string>();
            foreach (ScreenCaptureInfo capture in screenCaptures)
            {
                string filePath = Path.Combine(screenCaptureDirectory, string.Format("karl-klammer-codex-screen-{0}-screen{1}.jpg", timestamp, capture.ScreenNumber));
                File.WriteAllBytes(filePath, capture.ImageBytes);
                imagePaths.Add(filePath);
            }

            return imagePaths;
        }

        private static void DeleteFilesQuietly(IList<string> filePaths)
        {
            if (filePaths == null)
            {
                return;
            }

            foreach (string filePath in filePaths)
            {
                TryDeleteFile(filePath);
            }
        }

        private void HandleSpeechPress()
        {
            if (_isTranscribingSpeech)
            {
                return;
            }

            if (_microphoneRecorder.IsRecording)
            {
                return;
            }

            EnsureCompanionOverlay();
            StartSpeechCapture();
        }

        private async Task HandleSpeechReleaseAsync()
        {
            if (_isTranscribingSpeech || !_microphoneRecorder.IsRecording)
            {
                return;
            }

            await StopSpeechCaptureAsync();
        }

        private void StartSpeechCapture()
        {
            try
            {
                _microphoneRecorder.Start();
                SetCompanionState(CompanionVisualState.Listening);
                UpdateSpeechButtonState();
                SetStatus("listening... hold the button or " + _pushToTalkKey.ToString().ToLowerInvariant() + ", then release to transcribe", Color.FromArgb(235, 210, 120));
            }
            catch (Exception exception)
            {
                SetCompanionState(CompanionVisualState.Idle);
                UpdateSpeechButtonState();
                SetStatus("microphone error", Color.FromArgb(235, 120, 120));
                ShowCompanionMessage("microphone error", CompanionVisualState.Idle, 3200, CompanionVisualState.Idle);
                MessageBox.Show(exception.Message, "Karl Klammer microphone", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async Task StopSpeechCaptureAsync()
        {
            string audioFilePath = null;

            try
            {
                _isTranscribingSpeech = true;
                SetCompanionState(CompanionVisualState.Transcribing);
                UpdateSpeechButtonState();

                audioFilePath = _microphoneRecorder.Stop();
                string speechToTextLabel = SpeechToTextClient.GetProviderLabel(_environmentConfiguration);
                SetStatus("transcribing with " + speechToTextLabel + "...", Color.FromArgb(235, 210, 120));

                string transcript = await SpeechToTextClient.TranscribeAsync(_environmentConfiguration, audioFilePath);
                _promptTextBox.Text = transcript;
                _promptTextBox.SelectionStart = _promptTextBox.TextLength;
                _promptTextBox.ScrollToCaret();

                SetCompanionState(CompanionVisualState.Thinking);
                SetStatus(speechToTextLabel + " ready, asking Karl Klammer...", Color.FromArgb(235, 210, 120));
                await RunAskFlowAsync(transcript);
            }
            catch (Exception exception)
            {
                SetCompanionState(CompanionVisualState.Idle);
                SetStatus("speech error", Color.FromArgb(235, 120, 120));
                ShowCompanionMessage("speech error", CompanionVisualState.Idle, 3200, CompanionVisualState.Idle);
                MessageBox.Show(exception.Message, "Karl Klammer speech", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _isTranscribingSpeech = false;
                UpdateSpeechButtonState();
                TryDeleteFile(audioFilePath);
            }
        }

        private void PersistSettings()
        {
            _settings.AssistantProvider = _providerComboBox.SelectedItem == null ? "anthropic" : _providerComboBox.SelectedItem.ToString();
            _settings.AssistantModel = string.IsNullOrWhiteSpace(_modelTextBox.Text)
                ? AppSettings.GetDefaultModelForProvider(_settings.AssistantProvider)
                : _modelTextBox.Text.Trim();
            _settings.ClaudeModel = _settings.AssistantModel;
            _settings.CompanionMode = _modeComboBox.SelectedItem == null ? "companion" : _modeComboBox.SelectedItem.ToString();
            _settings.SuggestAutomations = _suggestAutomationsCheckBox.Checked;
            _settings.UseLocalKnowledge = _useKnowledgeCheckBox.Checked;
            _settings.SpeakResponses = _speakCheckBox.Checked;
            _settings.Save();
        }

        private void ReloadEnvironmentConfiguration()
        {
            _environmentConfiguration = EnvironmentConfiguration.Load();
            string validationError = _environmentConfiguration.Validate(_settings.AssistantProvider);
            ConfigurePushToTalkHotKey();

            if (string.IsNullOrWhiteSpace(validationError))
            {
                _envStatusLabel.Text = ".env loaded successfully";
                _envStatusLabel.ForeColor = Color.FromArgb(93, 212, 136);
            }
            else
            {
                _envStatusLabel.Text = validationError;
                _envStatusLabel.ForeColor = Color.FromArgb(235, 210, 120);
            }

            UpdateProviderVoiceSummary();
        }

        private void RefreshKnowledgeStatus()
        {
            string statusText = KnowledgeBaseService.GetStatusText();
            _knowledgeStatusLabel.Text = statusText + " | folder: " + KnowledgeBaseService.KnowledgeRoot;
            _knowledgeStatusLabel.ForeColor = Color.FromArgb(160, 174, 192);
        }

        private void RefreshKnowledgeManager()
        {
            string selectedKey = (_knowledgeListBox.SelectedItem as KnowledgeDocumentSummary)?.RelativePath;
            _knowledgeListBox.Items.Clear();

            string searchTerm = _knowledgeSearchTextBox == null ? string.Empty : (_knowledgeSearchTextBox.Text ?? string.Empty).Trim();
            List<KnowledgeDocumentSummary> documents = KnowledgeBaseService.SearchDocumentSummaries(searchTerm);
            foreach (KnowledgeDocumentSummary document in documents)
            {
                _knowledgeListBox.Items.Add(document);
            }

            if (_knowledgeListBox.Items.Count == 0)
            {
                _knowledgePreviewTextBox.Text = string.Empty;
                return;
            }

            KnowledgeDocumentSummary selected = documents.FirstOrDefault(document => string.Equals(document.RelativePath, selectedKey, StringComparison.OrdinalIgnoreCase));
            _knowledgeListBox.SelectedItem = selected ?? _knowledgeListBox.Items[0];
        }

        private void ImportKnowledgeDocuments()
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Title = "Import local knowledge documents";
                dialog.Filter = "Knowledge files|*.txt;*.md;*.log;*.json;*.csv;*.pdf;*.docx|All files|*.*";
                dialog.Multiselect = true;
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                int imported = KnowledgeBaseService.ImportFiles(dialog.FileNames);
                int chunkCount = KnowledgeBaseService.Reindex();
                RefreshKnowledgeStatus();
                RefreshKnowledgeManager();
                SetStatus("knowledge updated: " + imported + " file(s), " + chunkCount + " chunks", Color.FromArgb(93, 212, 136));
            }
        }

        private void RemoveSelectedKnowledgeDocument()
        {
            KnowledgeDocumentSummary selected = _knowledgeListBox.SelectedItem as KnowledgeDocumentSummary;
            if (selected == null)
            {
                SetStatus("select a knowledge doc first", Color.FromArgb(235, 210, 120));
                return;
            }

            DialogResult confirmation = MessageBox.Show(
                "Remove this knowledge file?\r\n" + selected.RelativePath,
                "Karl Klammer knowledge manager",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
            if (confirmation != DialogResult.Yes)
            {
                return;
            }

            if (!KnowledgeBaseService.DeleteDocument(selected.RelativePath))
            {
                SetStatus("could not remove knowledge doc", Color.FromArgb(235, 120, 120));
                return;
            }

            KnowledgeBaseService.Reindex();
            RefreshKnowledgeStatus();
            RefreshKnowledgeManager();
            SetStatus("knowledge doc removed", Color.FromArgb(93, 212, 136));
        }

        private void ReindexSelectedKnowledgeDocument()
        {
            KnowledgeDocumentSummary selected = _knowledgeListBox.SelectedItem as KnowledgeDocumentSummary;
            if (selected == null)
            {
                SetStatus("select a knowledge doc first", Color.FromArgb(235, 210, 120));
                return;
            }

            int chunkCount = KnowledgeBaseService.ReindexDocument(selected.RelativePath);
            RefreshKnowledgeStatus();
            RefreshKnowledgeManager();
            SetStatus("reindexed " + selected.Title + " (" + chunkCount + " total chunks)", Color.FromArgb(93, 212, 136));
        }

        private void UpdateRetrievalSourcesLabel(IList<KnowledgeChunk> knowledgeChunks)
        {
            if (knowledgeChunks == null || knowledgeChunks.Count == 0)
            {
                _retrievalSourcesLabel.Text = "sources: none";
                return;
            }

            string summary = string.Join(" | ", knowledgeChunks
                .Select(chunk => string.IsNullOrWhiteSpace(chunk.Title) ? "unknown" : chunk.Title)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(4)
                .ToArray());
            _retrievalSourcesLabel.Text = "sources: " + summary;
        }

        private void UpdateKnowledgePreview()
        {
            KnowledgeDocumentSummary selected = _knowledgeListBox.SelectedItem as KnowledgeDocumentSummary;
            if (selected == null)
            {
                _knowledgePreviewTextBox.Text = string.Empty;
                return;
            }

            _knowledgePreviewTextBox.Text = KnowledgeBaseService.GetPreviewText(selected.RelativePath, 320);
        }

        private void EnsureEnvironmentIsReady()
        {
            ReloadEnvironmentConfiguration();
            string validationError = _environmentConfiguration.Validate(_settings.AssistantProvider);
            if (!string.IsNullOrWhiteSpace(validationError))
            {
                throw new InvalidOperationException(validationError + " in " + EnvironmentConfiguration.EnvFilePath);
            }
        }

        private void ApplyProviderDefaultsIfNeeded()
        {
            string provider = _providerComboBox.SelectedItem == null ? "anthropic" : _providerComboBox.SelectedItem.ToString();
            string currentModel = _modelTextBox.Text == null ? string.Empty : _modelTextBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(currentModel)
                || string.Equals(currentModel, AppSettings.GetDefaultModelForProvider("anthropic"), StringComparison.OrdinalIgnoreCase)
                || string.Equals(currentModel, AppSettings.GetDefaultModelForProvider("openai"), StringComparison.OrdinalIgnoreCase)
                || string.Equals(currentModel, AppSettings.GetDefaultModelForProvider("openai-compatible"), StringComparison.OrdinalIgnoreCase))
            {
                _modelTextBox.Text = AppSettings.GetDefaultModelForProvider(provider);
            }
        }

        private void SaveCurrentPromptAsRecipe()
        {
            PersistSettings();
            string prompt = _promptTextBox.Text == null ? string.Empty : _promptTextBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(prompt))
            {
                SetStatus("type a prompt before saving a recipe", Color.FromArgb(235, 210, 120));
                return;
            }

            string recipeName = BuildRecipeName(prompt);
            bool exists = _automationRecipes.Any(recipe => string.Equals(recipe.Name, recipeName, StringComparison.OrdinalIgnoreCase));
            if (exists)
            {
                recipeName = recipeName + " " + DateTime.Now.ToString("HHmm");
            }

            _automationRecipes.Add(new AutomationRecipe
            {
                Name = recipeName,
                Prompt = prompt,
                CompanionMode = _settings.CompanionMode,
                CreatedAtUtc = DateTime.UtcNow.ToString("o")
            });
            AutomationRecipe.SaveAll(_automationRecipes);
            RefreshRecipeList(recipeName);
            SetStatus("recipe saved", Color.FromArgb(93, 212, 136));
        }

        private async Task RunSelectedRecipeAsync()
        {
            AutomationRecipe selectedRecipe = _recipeComboBox.SelectedItem as AutomationRecipe;
            if (selectedRecipe == null)
            {
                SetStatus("select a recipe first", Color.FromArgb(235, 210, 120));
                return;
            }

            _promptTextBox.Text = selectedRecipe.Prompt;
            _promptTextBox.SelectionStart = _promptTextBox.TextLength;
            _promptTextBox.ScrollToCaret();

            if (!string.IsNullOrWhiteSpace(selectedRecipe.CompanionMode) && _modeComboBox.Items.Contains(selectedRecipe.CompanionMode))
            {
                _modeComboBox.SelectedItem = selectedRecipe.CompanionMode;
            }

            await RunAskFlowAsync(selectedRecipe.Prompt);
        }

        private void RefreshRecipeList(string selectedName)
        {
            _recipeComboBox.Items.Clear();
            _recipeManagerListBox.Items.Clear();
            foreach (AutomationRecipe recipe in _automationRecipes.OrderBy(recipe => recipe.Name))
            {
                _recipeComboBox.Items.Add(recipe);
                _recipeManagerListBox.Items.Add(recipe);
            }

            if (_recipeComboBox.Items.Count == 0)
            {
                _recipePreviewTextBox.Text = string.Empty;
                return;
            }

            AutomationRecipe toSelect = null;
            if (!string.IsNullOrWhiteSpace(selectedName))
            {
                toSelect = _automationRecipes.FirstOrDefault(recipe => string.Equals(recipe.Name, selectedName, StringComparison.OrdinalIgnoreCase));
            }

            _recipeComboBox.SelectedItem = toSelect ?? _recipeComboBox.Items[0];
            _recipeManagerListBox.SelectedItem = toSelect ?? _recipeManagerListBox.Items[0];
        }

        private void RefreshWatchSuggestions(string selectedTitle)
        {
            _watchSuggestions.Clear();
            _watchSuggestions.AddRange(WatchSuggestionEngine.BuildSuggestions(WatchSessionEntry.LoadAll(), ActionHistoryEntry.LoadAll(), _automationRecipes));

            _watchSuggestionComboBox.Items.Clear();
            foreach (WatchSuggestion suggestion in _watchSuggestions)
            {
                _watchSuggestionComboBox.Items.Add(suggestion);
            }

            if (_watchSuggestionComboBox.Items.Count == 0)
            {
                return;
            }

            WatchSuggestion selectedSuggestion = null;
            if (!string.IsNullOrWhiteSpace(selectedTitle))
            {
                selectedSuggestion = _watchSuggestions.FirstOrDefault(suggestion => string.Equals(suggestion.Title, selectedTitle, StringComparison.OrdinalIgnoreCase));
            }

            _watchSuggestionComboBox.SelectedItem = selectedSuggestion ?? _watchSuggestionComboBox.Items[0];
            RefreshProactiveSuggestion();
        }

        private void RefreshActiveAppLabel()
        {
            ActiveWindowInfo activeWindow = ActiveWindowService.GetActiveWindowInfo();
            _lastActiveWindowInfo = activeWindow;
            _activeAppLabel.Text = "active app: " + activeWindow.DisplayName + " (" + (activeWindow.AppKind ?? "unknown") + ")";
            _appActionsLabel.Text = "app actions: " + AppActionAdapter.GetSupportedActionsSummary(activeWindow);
            _activeWindowDetailsLabel.Text = "window: " + (activeWindow.WindowTitle ?? "untitled")
                + " | class: " + (activeWindow.WindowClassName ?? "n/a")
                + " | framework: " + (activeWindow.DesktopFramework ?? "unknown");
        }

        private void RefreshProactiveSuggestion()
        {
            ActiveWindowInfo activeWindow = ActiveWindowService.GetActiveWindowInfo();
            _currentProactiveSuggestion = ProactiveSuggestionEngine.Build(activeWindow, _watchSuggestions);
            if (_currentProactiveSuggestion == null)
            {
                _proactiveSuggestionLabel.Text = "next idea: no strong context suggestion yet";
                _proactiveSuggestionLabel.ForeColor = Color.FromArgb(160, 174, 192);
                return;
            }

            _proactiveSuggestionLabel.Text = "next idea: " + _currentProactiveSuggestion;
            _proactiveSuggestionLabel.ForeColor = string.Equals(_currentProactiveSuggestion.Source, "ritual", StringComparison.OrdinalIgnoreCase)
                ? Color.FromArgb(93, 212, 136)
                : Color.FromArgb(255, 183, 77);
        }

        private void SaveSelectedWatchSuggestionAsRecipe()
        {
            WatchSuggestion selectedSuggestion = _watchSuggestionComboBox.SelectedItem as WatchSuggestion;
            if (selectedSuggestion == null)
            {
                SetStatus("no learned ritual available yet", Color.FromArgb(235, 210, 120));
                return;
            }

            bool exists = _automationRecipes.Any(recipe => string.Equals(recipe.Prompt, selectedSuggestion.Prompt, StringComparison.OrdinalIgnoreCase));
            if (exists)
            {
                SetStatus("ritual already saved as recipe", Color.FromArgb(235, 210, 120));
                return;
            }

            _automationRecipes.Add(new AutomationRecipe
            {
                Name = selectedSuggestion.Title,
                Prompt = selectedSuggestion.Prompt,
                CompanionMode = selectedSuggestion.CompanionMode,
                CreatedAtUtc = DateTime.UtcNow.ToString("o")
            });
            AutomationRecipe.SaveAll(_automationRecipes);
            RefreshRecipeList(selectedSuggestion.Title);
            RefreshWatchSuggestions(null);
            SetStatus("ritual saved as recipe", Color.FromArgb(93, 212, 136));
        }

        private void UseSelectedWatchSuggestion()
        {
            WatchSuggestion selectedSuggestion = _watchSuggestionComboBox.SelectedItem as WatchSuggestion;
            if (selectedSuggestion == null)
            {
                SetStatus("no learned ritual available yet", Color.FromArgb(235, 210, 120));
                return;
            }

            _promptTextBox.Text = selectedSuggestion.Prompt;
            _promptTextBox.SelectionStart = _promptTextBox.TextLength;
            _promptTextBox.ScrollToCaret();

            if (!string.IsNullOrWhiteSpace(selectedSuggestion.CompanionMode) && _modeComboBox.Items.Contains(selectedSuggestion.CompanionMode))
            {
                _modeComboBox.SelectedItem = selectedSuggestion.CompanionMode;
            }

            SetStatus("ritual loaded into prompt", Color.FromArgb(93, 212, 136));
        }

        private void UpdateRecipeManagerPreview()
        {
            AutomationRecipe selectedRecipe = _recipeManagerListBox.SelectedItem as AutomationRecipe;
            if (selectedRecipe == null)
            {
                _recipePreviewTextBox.Text = string.Empty;
                return;
            }

            _recipePreviewTextBox.Text =
                "name: " + selectedRecipe.Name + Environment.NewLine +
                "mode: " + (selectedRecipe.CompanionMode ?? "companion") + Environment.NewLine +
                "created: " + (selectedRecipe.CreatedAtUtc ?? string.Empty) + Environment.NewLine + Environment.NewLine +
                (selectedRecipe.Prompt ?? string.Empty);
            _recipeComboBox.SelectedItem = selectedRecipe;
        }

        private void DeleteSelectedRecipe()
        {
            AutomationRecipe selectedRecipe = _recipeManagerListBox.SelectedItem as AutomationRecipe;
            if (selectedRecipe == null)
            {
                SetStatus("select a ritual first", Color.FromArgb(235, 210, 120));
                return;
            }

            DialogResult confirmation = MessageBox.Show(
                "Delete ritual?\r\n" + selectedRecipe.Name,
                "Karl Klammer ritual manager",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
            if (confirmation != DialogResult.Yes)
            {
                return;
            }

            _automationRecipes.Remove(selectedRecipe);
            AutomationRecipe.SaveAll(_automationRecipes);
            RefreshRecipeList(null);
            RefreshWatchSuggestions(null);
            SetStatus("ritual deleted", Color.FromArgb(93, 212, 136));
        }

        private void ShowRitualManagerDialog()
        {
            using (Form dialog = CreateUtilityDialog("Ritual Manager", 900, 620))
            {
                ListBox ritualListBox = CreateUtilityListBox(16, 16, 240, 520);
                TextBox nameTextBox = CreateUtilityTextBox(272, 42, 260, 30, string.Empty);
                ComboBox modeComboBox = new ComboBox();
                modeComboBox.Location = new Point(548, 42);
                modeComboBox.Size = new Size(160, 30);
                modeComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
                modeComboBox.Items.AddRange(new object[] { "companion", "agent", "automation", "watch" });
                TextBox promptTextBox = CreateUtilityTextBox(272, 102, 596, 360, string.Empty);
                promptTextBox.Multiline = true;
                promptTextBox.ScrollBars = ScrollBars.Vertical;
                Button saveButton = CreateButton("save ritual", 272, 478, 140, 34, Color.FromArgb(33, 150, 243), Color.White);
                Button deleteButton = CreateButton("delete ritual", 422, 478, 140, 34, Color.FromArgb(89, 40, 40), Color.White);
                Button runButton = CreateButton("run ritual", 572, 478, 140, 34, Color.FromArgb(22, 30, 45), Color.White);

                dialog.Controls.Add(CreateLabel("rituals", 16, 0, 100, 18, new Font("Segoe UI Semibold", 9.5f), Color.White));
                dialog.Controls.Add(CreateLabel("name", 272, 18, 100, 18, new Font("Segoe UI Semibold", 9.0f), Color.White));
                dialog.Controls.Add(CreateLabel("mode", 548, 18, 100, 18, new Font("Segoe UI Semibold", 9.0f), Color.White));
                dialog.Controls.Add(CreateLabel("prompt / workflow", 272, 78, 180, 18, new Font("Segoe UI Semibold", 9.0f), Color.White));
                dialog.Controls.Add(ritualListBox);
                dialog.Controls.Add(nameTextBox);
                dialog.Controls.Add(modeComboBox);
                dialog.Controls.Add(promptTextBox);
                dialog.Controls.Add(saveButton);
                dialog.Controls.Add(deleteButton);
                dialog.Controls.Add(runButton);

                Action refreshList = delegate
                {
                    string selectedName = (ritualListBox.SelectedItem as AutomationRecipe)?.Name;
                    ritualListBox.Items.Clear();
                    foreach (AutomationRecipe recipe in _automationRecipes.OrderBy(recipe => recipe.Name))
                    {
                        ritualListBox.Items.Add(recipe);
                    }

                    if (ritualListBox.Items.Count > 0)
                    {
                        AutomationRecipe selected = _automationRecipes.FirstOrDefault(recipe => string.Equals(recipe.Name, selectedName, StringComparison.OrdinalIgnoreCase));
                        ritualListBox.SelectedItem = selected ?? ritualListBox.Items[0];
                    }
                };

                ritualListBox.SelectedIndexChanged += delegate
                {
                    AutomationRecipe selected = ritualListBox.SelectedItem as AutomationRecipe;
                    if (selected == null)
                    {
                        nameTextBox.Text = string.Empty;
                        promptTextBox.Text = string.Empty;
                        modeComboBox.SelectedIndex = -1;
                        return;
                    }

                    nameTextBox.Text = selected.Name ?? string.Empty;
                    promptTextBox.Text = selected.Prompt ?? string.Empty;
                    modeComboBox.SelectedItem = string.IsNullOrWhiteSpace(selected.CompanionMode) ? "automation" : selected.CompanionMode;
                };

                saveButton.Click += delegate
                {
                    AutomationRecipe selected = ritualListBox.SelectedItem as AutomationRecipe;
                    if (selected == null)
                    {
                        return;
                    }

                    selected.Name = (nameTextBox.Text ?? string.Empty).Trim();
                    selected.Prompt = (promptTextBox.Text ?? string.Empty).Trim();
                    selected.CompanionMode = modeComboBox.SelectedItem == null ? "automation" : modeComboBox.SelectedItem.ToString();
                    AutomationRecipe.SaveAll(_automationRecipes);
                    RefreshRecipeList(selected.Name);
                    RefreshWatchSuggestions(null);
                    refreshList();
                    AddDiagnosticItem("ritual", "Updated ritual '" + selected.Name + "'.");
                };

                deleteButton.Click += delegate
                {
                    AutomationRecipe selected = ritualListBox.SelectedItem as AutomationRecipe;
                    if (selected == null)
                    {
                        return;
                    }

                    _automationRecipes.Remove(selected);
                    AutomationRecipe.SaveAll(_automationRecipes);
                    RefreshRecipeList(null);
                    RefreshWatchSuggestions(null);
                    refreshList();
                    AddDiagnosticItem("ritual", "Deleted ritual '" + selected.Name + "'.");
                };

                runButton.Click += async delegate
                {
                    AutomationRecipe selected = ritualListBox.SelectedItem as AutomationRecipe;
                    if (selected == null)
                    {
                        return;
                    }

                    dialog.Hide();
                    _promptTextBox.Text = selected.Prompt ?? string.Empty;
                    if (_modeComboBox.Items.Contains(selected.CompanionMode))
                    {
                        _modeComboBox.SelectedItem = selected.CompanionMode;
                    }

                    await RunAskFlowAsync(selected.Prompt);
                    dialog.Show();
                };

                refreshList();
                dialog.ShowDialog(this);
            }
        }

        private void ShowActionHistoryDialog()
        {
            using (Form dialog = CreateUtilityDialog("Action History", 920, 620))
            {
                List<ActionHistoryEntry> history = ActionHistoryEntry.LoadAll().OrderByDescending(entry => entry.TimestampUtc).ToList();
                TextBox searchTextBox = CreateUtilityTextBox(16, 42, 290, 30, string.Empty);
                ListBox historyListBox = CreateUtilityListBox(16, 16, 290, 520);
                historyListBox.Location = new Point(16, 82);
                historyListBox.Size = new Size(290, 454);
                TextBox detailTextBox = CreateUtilityTextBox(322, 42, 566, 470, string.Empty);
                detailTextBox.Multiline = true;
                detailTextBox.ScrollBars = ScrollBars.Vertical;
                Button loadPromptButton = CreateButton("load prompt", 322, 526, 130, 34, Color.FromArgb(22, 30, 45), Color.White);
                Button replayButton = CreateButton("replay action", 462, 526, 130, 34, Color.FromArgb(39, 88, 68), Color.White);
                Button refreshButton = CreateButton("refresh", 462, 526, 100, 34, Color.FromArgb(22, 30, 45), Color.White);
                refreshButton.Location = new Point(602, 526);

                dialog.Controls.Add(CreateLabel("action history", 16, 0, 140, 18, new Font("Segoe UI Semibold", 9.5f), Color.White));
                dialog.Controls.Add(CreateLabel("search", 16, 18, 100, 18, new Font("Segoe UI Semibold", 9.0f), Color.White));
                dialog.Controls.Add(CreateLabel("details", 322, 18, 100, 18, new Font("Segoe UI Semibold", 9.0f), Color.White));
                dialog.Controls.Add(searchTextBox);
                dialog.Controls.Add(historyListBox);
                dialog.Controls.Add(detailTextBox);
                dialog.Controls.Add(loadPromptButton);
                dialog.Controls.Add(replayButton);
                dialog.Controls.Add(refreshButton);

                Action reload = delegate
                {
                    history = ActionHistoryEntry.LoadAll().OrderByDescending(entry => entry.TimestampUtc).ToList();
                    string filter = (searchTextBox.Text ?? string.Empty).Trim();
                    historyListBox.Items.Clear();
                    foreach (ActionHistoryEntry entry in history)
                    {
                        string haystack = string.Join(" ", new[]
                        {
                            entry.TimestampUtc ?? string.Empty,
                            entry.ActionName ?? string.Empty,
                            entry.ActionArgument ?? string.Empty,
                            entry.TargetLabel ?? string.Empty,
                            entry.ActiveApp ?? string.Empty,
                            entry.SpokenText ?? string.Empty
                        });
                        if (filter.Length > 0 && haystack.IndexOf(filter, StringComparison.OrdinalIgnoreCase) < 0)
                        {
                            continue;
                        }

                        historyListBox.Items.Add((entry.TimestampUtc ?? string.Empty) + " | " + (entry.ActionName ?? string.Empty) + " | " + (entry.ActiveApp ?? string.Empty));
                    }
                    if (historyListBox.Items.Count > 0)
                    {
                        historyListBox.SelectedIndex = 0;
                    }
                };

                historyListBox.SelectedIndexChanged += delegate
                {
                    if (historyListBox.SelectedIndex < 0 || historyListBox.SelectedIndex >= history.Count)
                    {
                        detailTextBox.Text = string.Empty;
                        return;
                    }

                    string selectedLine = historyListBox.SelectedItem == null ? string.Empty : historyListBox.SelectedItem.ToString();
                    ActionHistoryEntry entry = history.FirstOrDefault(candidate =>
                        selectedLine == ((candidate.TimestampUtc ?? string.Empty) + " | " + (candidate.ActionName ?? string.Empty) + " | " + (candidate.ActiveApp ?? string.Empty)));
                    if (entry == null)
                    {
                        detailTextBox.Text = string.Empty;
                        return;
                    }

                    detailTextBox.Text =
                        "timestamp: " + (entry.TimestampUtc ?? string.Empty) + Environment.NewLine +
                        "action: " + (entry.ActionName ?? string.Empty) + Environment.NewLine +
                        "argument: " + (entry.ActionArgument ?? string.Empty) + Environment.NewLine +
                        "target: " + (entry.TargetLabel ?? string.Empty) + Environment.NewLine +
                        "app: " + (entry.ActiveApp ?? string.Empty) + Environment.NewLine + Environment.NewLine +
                        "spoken text:" + Environment.NewLine + (entry.SpokenText ?? string.Empty);
                };

                loadPromptButton.Click += delegate
                {
                    if (historyListBox.SelectedIndex < 0 || historyListBox.SelectedIndex >= history.Count)
                    {
                        return;
                    }

                    string selectedLine = historyListBox.SelectedItem == null ? string.Empty : historyListBox.SelectedItem.ToString();
                    ActionHistoryEntry entry = history.FirstOrDefault(candidate =>
                        selectedLine == ((candidate.TimestampUtc ?? string.Empty) + " | " + (candidate.ActionName ?? string.Empty) + " | " + (candidate.ActiveApp ?? string.Empty)));
                    if (entry == null)
                    {
                        return;
                    }

                    _promptTextBox.Text = entry.SpokenText ?? string.Empty;
                    _promptTextBox.SelectionStart = _promptTextBox.TextLength;
                    _promptTextBox.ScrollToCaret();
                    dialog.Close();
                };

                replayButton.Click += async delegate
                {
                    if (historyListBox.SelectedIndex < 0 || historyListBox.SelectedItem == null)
                    {
                        return;
                    }

                    string selectedLine = historyListBox.SelectedItem.ToString();
                    ActionHistoryEntry entry = history.FirstOrDefault(candidate =>
                        selectedLine == ((candidate.TimestampUtc ?? string.Empty) + " | " + (candidate.ActionName ?? string.Empty) + " | " + (candidate.ActiveApp ?? string.Empty)));
                    if (entry == null || string.IsNullOrWhiteSpace(entry.ActionName))
                    {
                        return;
                    }

                    dialog.Hide();
                    List<ScreenCaptureInfo> screenCaptures = ScreenCaptureService.CaptureAllScreens();
                    ActionTagResult action = new ActionTagResult
                    {
                        CleanText = entry.SpokenText ?? string.Empty,
                        ActionName = entry.ActionName ?? string.Empty,
                        ActionArgument = entry.ActionArgument ?? string.Empty
                    };
                    PointTagResult pointTag = new PointTagResult
                    {
                        SpokenText = entry.SpokenText ?? string.Empty,
                        ElementLabel = entry.TargetLabel ?? "history target"
                    };

                    await ExecuteActionIfConfirmedAsync(action, pointTag, screenCaptures, entry.SpokenText ?? string.Empty);
                    dialog.Show();
                };

                searchTextBox.TextChanged += delegate { reload(); };
                refreshButton.Click += delegate { reload(); };

                reload();
                dialog.ShowDialog(this);
            }
        }

        private void ShowControlInspectorDialog()
        {
            using (Form dialog = CreateUtilityDialog("Window / Control Inspector", 960, 680))
            {
                Label appLabel = CreateLabel(string.Empty, 16, 16, 900, 20, new Font("Segoe UI Semibold", 9.5f), Color.White);
                TextBox customActionTextBox = CreateUtilityTextBox(626, 48, 302, 30, "focus_control:search");
                TextBox resultTextBox = CreateUtilityTextBox(16, 128, 912, 484, string.Empty);
                resultTextBox.Multiline = true;
                resultTextBox.ScrollBars = ScrollBars.Vertical;
                Button listControlsButton = CreateButton("list controls", 16, 46, 120, 32, Color.FromArgb(22, 30, 45), Color.White);
                Button readFormButton = CreateButton("read form", 146, 46, 110, 32, Color.FromArgb(22, 30, 45), Color.White);
                Button readTableButton = CreateButton("read table", 266, 46, 110, 32, Color.FromArgb(22, 30, 45), Color.White);
                Button readDialogButton = CreateButton("read dialog", 386, 46, 110, 32, Color.FromArgb(22, 30, 45), Color.White);
                Button readSelectionButton = CreateButton("selected row", 506, 46, 110, 32, Color.FromArgb(22, 30, 45), Color.White);
                Button runCustomActionButton = CreateButton("run app action", 626, 86, 140, 32, Color.FromArgb(39, 88, 68), Color.White);
                Button refreshButton = CreateButton("refresh app", 778, 86, 110, 32, Color.FromArgb(22, 30, 45), Color.White);

                dialog.Controls.Add(appLabel);
                dialog.Controls.Add(CreateLabel("custom action", 626, 24, 120, 18, new Font("Segoe UI Semibold", 9.0f), Color.White));
                dialog.Controls.Add(customActionTextBox);
                dialog.Controls.Add(resultTextBox);
                dialog.Controls.Add(listControlsButton);
                dialog.Controls.Add(readFormButton);
                dialog.Controls.Add(readTableButton);
                dialog.Controls.Add(readDialogButton);
                dialog.Controls.Add(readSelectionButton);
                dialog.Controls.Add(runCustomActionButton);
                dialog.Controls.Add(refreshButton);

                Action<string> runInspector = delegate(string actionArgument)
                {
                    ActiveWindowInfo activeWindow = ActiveWindowService.GetActiveWindowInfo();
                    appLabel.Text = "active window: " + activeWindow.DisplayName + " | class: " + (activeWindow.WindowClassName ?? string.Empty) + " | framework: " + (activeWindow.DesktopFramework ?? string.Empty);
                    string result = FatClientAutomationService.ExecuteOrInspect(activeWindow, actionArgument);
                    resultTextBox.Text = result;
                    AddDiagnosticItem("inspect", actionArgument + " on " + activeWindow.DisplayName);
                };

                listControlsButton.Click += delegate { runInspector("list_controls"); };
                readFormButton.Click += delegate { runInspector("read_form"); };
                readTableButton.Click += delegate { runInspector("read_table"); };
                readDialogButton.Click += delegate { runInspector("read_dialog"); };
                readSelectionButton.Click += delegate { runInspector("read_selected_row"); };
                runCustomActionButton.Click += delegate
                {
                    string actionArgument = (customActionTextBox.Text ?? string.Empty).Trim();
                    if (actionArgument.Length == 0)
                    {
                        return;
                    }

                    runInspector(actionArgument);
                };
                refreshButton.Click += delegate { runInspector("list_controls"); };

                runInspector("list_controls");
                dialog.ShowDialog(this);
            }
        }

        private void ShowProviderVoiceDialog()
        {
            using (Form dialog = CreateUtilityDialog("Provider / Voice Settings", 760, 420))
            {
                ComboBox providerComboBox = new ComboBox();
                providerComboBox.Location = new Point(16, 42);
                providerComboBox.Size = new Size(180, 30);
                providerComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
                providerComboBox.Items.AddRange(new object[] { "anthropic", "openai", "openai-compatible" });
                providerComboBox.SelectedItem = _providerComboBox.SelectedItem;

                TextBox modelTextBox = CreateUtilityTextBox(212, 42, 250, 30, _modelTextBox.Text ?? string.Empty);
                CheckBox speakCheckBox = new CheckBox();
                speakCheckBox.Text = "speak responses";
                speakCheckBox.Location = new Point(16, 92);
                speakCheckBox.Size = new Size(180, 24);
                speakCheckBox.Checked = _speakCheckBox.Checked;
                speakCheckBox.ForeColor = Color.White;
                speakCheckBox.BackColor = dialog.BackColor;

                TextBox voiceIdTextBox = CreateUtilityTextBox(16, 152, 446, 30, _environmentConfiguration.ElevenLabsVoiceId ?? string.Empty);
                voiceIdTextBox.ReadOnly = true;
                TextBox infoTextBox = CreateUtilityTextBox(16, 214, 706, 120, string.Empty);
                infoTextBox.Multiline = true;
                infoTextBox.ReadOnly = true;
                Button applyButton = CreateButton("apply", 478, 42, 110, 34, Color.FromArgb(33, 150, 243), Color.White);
                Button ttsTestButton = CreateButton("test tts", 598, 42, 110, 34, Color.FromArgb(22, 30, 45), Color.White);
                Button reloadEnvButton = CreateButton("reload env", 478, 86, 110, 34, Color.FromArgb(22, 30, 45), Color.White);
                Button smokeTestButton = CreateButton("smoke test", 598, 86, 110, 34, Color.FromArgb(22, 30, 45), Color.White);

                dialog.Controls.Add(CreateLabel("provider", 16, 18, 100, 18, new Font("Segoe UI Semibold", 9.0f), Color.White));
                dialog.Controls.Add(CreateLabel("model", 212, 18, 100, 18, new Font("Segoe UI Semibold", 9.0f), Color.White));
                dialog.Controls.Add(CreateLabel("elevenlabs voice id (.env)", 16, 128, 220, 18, new Font("Segoe UI Semibold", 9.0f), Color.White));
                dialog.Controls.Add(CreateLabel("health + routing", 16, 190, 140, 18, new Font("Segoe UI Semibold", 9.0f), Color.White));
                dialog.Controls.Add(providerComboBox);
                dialog.Controls.Add(modelTextBox);
                dialog.Controls.Add(speakCheckBox);
                dialog.Controls.Add(voiceIdTextBox);
                dialog.Controls.Add(infoTextBox);
                dialog.Controls.Add(applyButton);
                dialog.Controls.Add(ttsTestButton);
                dialog.Controls.Add(reloadEnvButton);
                dialog.Controls.Add(smokeTestButton);

                Action refreshInfo = delegate
                {
                    string provider = providerComboBox.SelectedItem == null ? "anthropic" : providerComboBox.SelectedItem.ToString();
                    bool providerReady =
                        (provider == "openai" || provider == "openai-compatible")
                            ? !string.IsNullOrWhiteSpace(_environmentConfiguration.OpenAIApiKey)
                            : !string.IsNullOrWhiteSpace(_environmentConfiguration.AnthropicApiKey);
                    infoTextBox.Text =
                        "provider ready: " + providerReady + Environment.NewLine +
                        "stt provider: " + SpeechToTextClient.GetProviderLabel(_environmentConfiguration) + Environment.NewLine +
                        "tts provider: " + (!string.IsNullOrWhiteSpace(_environmentConfiguration.ElevenLabsApiKey) ? "elevenlabs" : "local fallback") + Environment.NewLine +
                        "openai base url: " + (_environmentConfiguration.OpenAIBaseUrl ?? string.Empty) + Environment.NewLine +
                        "env path: " + EnvironmentConfiguration.EnvFilePath;
                };

                applyButton.Click += delegate
                {
                    _providerComboBox.SelectedItem = providerComboBox.SelectedItem;
                    _modelTextBox.Text = modelTextBox.Text;
                    _speakCheckBox.Checked = speakCheckBox.Checked;
                    PersistSettings();
                    UpdateProviderVoiceSummary();
                    AddDiagnosticItem("settings", "Updated provider/voice settings.");
                    refreshInfo();
                };

                ttsTestButton.Click += async delegate
                {
                    PersistSettings();
                    await SpeakResponseAsync("Karl Klammer testet jetzt die aktive Stimme.");
                };

                reloadEnvButton.Click += delegate
                {
                    ReloadEnvironmentConfiguration();
                    voiceIdTextBox.Text = _environmentConfiguration.ElevenLabsVoiceId ?? string.Empty;
                    UpdateProviderVoiceSummary();
                    AddDiagnosticItem("settings", "Reloaded environment configuration.");
                    refreshInfo();
                };

                smokeTestButton.Click += async delegate
                {
                    PersistSettings();
                    await RunWorkerTestAsync();
                    refreshInfo();
                };

                refreshInfo();
                dialog.ShowDialog(this);
            }
        }

        private void ShowDiagnosticsDialog()
        {
            using (Form dialog = CreateUtilityDialog("Diagnostics Center", 920, 620))
            {
                TextBox filterTextBox = CreateUtilityTextBox(16, 42, 300, 30, string.Empty);
                ListBox listBox = CreateUtilityListBox(16, 82, 300, 454);
                TextBox detailTextBox = CreateUtilityTextBox(332, 42, 556, 470, string.Empty);
                detailTextBox.Multiline = true;
                detailTextBox.ScrollBars = ScrollBars.Vertical;
                Button copyButton = CreateButton("copy selected", 332, 526, 130, 34, Color.FromArgb(22, 30, 45), Color.White);
                Button copyAllButton = CreateButton("copy all", 472, 526, 110, 34, Color.FromArgb(22, 30, 45), Color.White);
                Button exportButton = CreateButton("export", 592, 526, 100, 34, Color.FromArgb(22, 30, 45), Color.White);
                Button clearButton = CreateButton("clear feed", 472, 526, 120, 34, Color.FromArgb(89, 40, 40), Color.White);
                clearButton.Location = new Point(702, 526);

                dialog.Controls.Add(CreateLabel("events", 16, 0, 100, 18, new Font("Segoe UI Semibold", 9.5f), Color.White));
                dialog.Controls.Add(CreateLabel("filter", 16, 18, 100, 18, new Font("Segoe UI Semibold", 9.0f), Color.White));
                dialog.Controls.Add(CreateLabel("detail", 332, 18, 100, 18, new Font("Segoe UI Semibold", 9.0f), Color.White));
                dialog.Controls.Add(filterTextBox);
                dialog.Controls.Add(listBox);
                dialog.Controls.Add(detailTextBox);
                dialog.Controls.Add(copyButton);
                dialog.Controls.Add(copyAllButton);
                dialog.Controls.Add(exportButton);
                dialog.Controls.Add(clearButton);

                Action refresh = delegate
                {
                    listBox.Items.Clear();
                    string filter = (filterTextBox.Text ?? string.Empty).Trim();
                    foreach (object item in _diagnosticsListBox.Items)
                    {
                        string line = item == null ? string.Empty : item.ToString();
                        if (filter.Length > 0 && line.IndexOf(filter, StringComparison.OrdinalIgnoreCase) < 0)
                        {
                            continue;
                        }

                        listBox.Items.Add(line);
                    }

                    if (listBox.Items.Count > 0)
                    {
                        listBox.SelectedIndex = 0;
                    }
                };

                listBox.SelectedIndexChanged += delegate
                {
                    detailTextBox.Text = listBox.SelectedItem == null ? string.Empty : listBox.SelectedItem.ToString();
                };

                copyButton.Click += delegate
                {
                    if (listBox.SelectedItem != null)
                    {
                        Clipboard.SetText(listBox.SelectedItem.ToString());
                    }
                };

                copyAllButton.Click += delegate
                {
                    if (listBox.Items.Count == 0)
                    {
                        return;
                    }

                    Clipboard.SetText(string.Join(Environment.NewLine, listBox.Items.Cast<object>().Select(item => item == null ? string.Empty : item.ToString()).ToArray()));
                };

                exportButton.Click += delegate
                {
                    string path = ExportDiagnosticsSnapshot(listBox.Items.Cast<object>().Select(item => item == null ? string.Empty : item.ToString()).ToList());
                    detailTextBox.Text = "exported diagnostics to:" + Environment.NewLine + path;
                };

                clearButton.Click += delegate
                {
                    _diagnosticsListBox.Items.Clear();
                    refresh();
                };

                filterTextBox.TextChanged += delegate { refresh(); };
                refresh();
                dialog.ShowDialog(this);
            }
        }

        private void ShowSetupWizardDialog()
        {
            using (Form dialog = CreateUtilityDialog("Setup Wizard", 760, 520))
            {
                CheckBox useKnowledgeCheckBox = new CheckBox();
                useKnowledgeCheckBox.Text = "use local knowledge";
                useKnowledgeCheckBox.Location = new Point(16, 118);
                useKnowledgeCheckBox.Size = new Size(180, 24);
                useKnowledgeCheckBox.Checked = _useKnowledgeCheckBox.Checked;
                useKnowledgeCheckBox.ForeColor = Color.White;
                useKnowledgeCheckBox.BackColor = dialog.BackColor;

                CheckBox speakCheckBox = new CheckBox();
                speakCheckBox.Text = "speak responses";
                speakCheckBox.Location = new Point(208, 118);
                speakCheckBox.Size = new Size(180, 24);
                speakCheckBox.Checked = _speakCheckBox.Checked;
                speakCheckBox.ForeColor = Color.White;
                speakCheckBox.BackColor = dialog.BackColor;

                ComboBox providerComboBox = new ComboBox();
                providerComboBox.Location = new Point(16, 44);
                providerComboBox.Size = new Size(170, 30);
                providerComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
                providerComboBox.Items.AddRange(new object[] { "anthropic", "openai", "openai-compatible" });
                providerComboBox.SelectedItem = _providerComboBox.SelectedItem;

                ComboBox modeComboBox = new ComboBox();
                modeComboBox.Location = new Point(198, 44);
                modeComboBox.Size = new Size(170, 30);
                modeComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
                modeComboBox.Items.AddRange(new object[] { "companion", "agent", "automation", "watch" });
                modeComboBox.SelectedItem = _modeComboBox.SelectedItem;

                TextBox modelTextBox = CreateUtilityTextBox(380, 44, 330, 30, _modelTextBox.Text ?? string.Empty);
                TextBox summaryTextBox = CreateUtilityTextBox(16, 182, 694, 246, string.Empty);
                summaryTextBox.Multiline = true;
                summaryTextBox.ReadOnly = true;
                summaryTextBox.ScrollBars = ScrollBars.Vertical;
                Button reloadButton = CreateButton("reload env", 16, 438, 120, 34, Color.FromArgb(22, 30, 45), Color.White);
                Button testButton = CreateButton("test apis", 146, 438, 120, 34, Color.FromArgb(22, 30, 45), Color.White);
                Button applyButton = CreateButton("apply setup", 590, 438, 120, 34, Color.FromArgb(33, 150, 243), Color.White);

                dialog.Controls.Add(CreateLabel("provider", 16, 18, 100, 18, new Font("Segoe UI Semibold", 9.0f), Color.White));
                dialog.Controls.Add(CreateLabel("mode", 198, 18, 100, 18, new Font("Segoe UI Semibold", 9.0f), Color.White));
                dialog.Controls.Add(CreateLabel("model", 380, 18, 100, 18, new Font("Segoe UI Semibold", 9.0f), Color.White));
                dialog.Controls.Add(CreateLabel("health + next steps", 16, 156, 180, 18, new Font("Segoe UI Semibold", 9.0f), Color.White));
                dialog.Controls.Add(providerComboBox);
                dialog.Controls.Add(modeComboBox);
                dialog.Controls.Add(modelTextBox);
                dialog.Controls.Add(useKnowledgeCheckBox);
                dialog.Controls.Add(speakCheckBox);
                dialog.Controls.Add(summaryTextBox);
                dialog.Controls.Add(reloadButton);
                dialog.Controls.Add(testButton);
                dialog.Controls.Add(applyButton);

                Action refreshSummary = delegate
                {
                    string provider = providerComboBox.SelectedItem == null ? "anthropic" : providerComboBox.SelectedItem.ToString();
                    bool providerReady =
                        (provider == "openai" || provider == "openai-compatible")
                            ? !string.IsNullOrWhiteSpace(_environmentConfiguration.OpenAIApiKey)
                            : !string.IsNullOrWhiteSpace(_environmentConfiguration.AnthropicApiKey);
                    summaryTextBox.Text =
                        "env path: " + EnvironmentConfiguration.EnvFilePath + Environment.NewLine +
                        "provider ready: " + providerReady + Environment.NewLine +
                        "tts ready: " + (!string.IsNullOrWhiteSpace(_environmentConfiguration.ElevenLabsApiKey)) + Environment.NewLine +
                        "stt: " + SpeechToTextClient.GetProviderLabel(_environmentConfiguration) + Environment.NewLine +
                        "knowledge docs: " + KnowledgeBaseService.GetDocumentSummaries().Count + Environment.NewLine +
                        "knowledge enabled: " + useKnowledgeCheckBox.Checked + Environment.NewLine +
                        "speak enabled: " + speakCheckBox.Checked + Environment.NewLine + Environment.NewLine +
                        "recommended next steps:" + Environment.NewLine +
                        "- reload env after editing keys" + Environment.NewLine +
                        "- run test apis" + Environment.NewLine +
                        "- import docs into knowledge manager if you want local RAG" + Environment.NewLine +
                        "- use window inspector for fat clients";
                };

                reloadButton.Click += delegate
                {
                    ReloadEnvironmentConfiguration();
                    refreshSummary();
                    AddDiagnosticItem("setup", "Reloaded env from setup wizard.");
                };

                testButton.Click += async delegate
                {
                    PersistSettings();
                    await RunWorkerTestAsync();
                    refreshSummary();
                };

                applyButton.Click += delegate
                {
                    _providerComboBox.SelectedItem = providerComboBox.SelectedItem;
                    _modeComboBox.SelectedItem = modeComboBox.SelectedItem;
                    _modelTextBox.Text = modelTextBox.Text;
                    _useKnowledgeCheckBox.Checked = useKnowledgeCheckBox.Checked;
                    _speakCheckBox.Checked = speakCheckBox.Checked;
                    PersistSettings();
                    UpdateModeBadge();
                    UpdateProviderVoiceSummary();
                    RefreshKnowledgeStatus();
                    AddDiagnosticItem("setup", "Applied setup wizard settings.");
                    refreshSummary();
                };

                refreshSummary();
                dialog.ShowDialog(this);
            }
        }

        private static Form CreateUtilityDialog(string title, int width, int height)
        {
            return new Form
            {
                Text = title,
                Width = width,
                Height = height,
                StartPosition = FormStartPosition.CenterParent,
                BackColor = Color.FromArgb(11, 17, 27),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9.0f)
            };
        }

        private static ListBox CreateUtilityListBox(int x, int y, int width, int height)
        {
            return new ListBox
            {
                Location = new Point(x, y),
                Size = new Size(width, height),
                Font = new Font("Segoe UI", 9.0f),
                BackColor = Color.FromArgb(19, 29, 43),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };
        }

        private static TextBox CreateUtilityTextBox(int x, int y, int width, int height, string text)
        {
            TextBox textBox = CreateTextBox(x, y, width, height, text);
            textBox.BackColor = Color.FromArgb(19, 29, 43);
            textBox.ForeColor = Color.White;
            return textBox;
        }

        private void UpdateProviderVoiceSummary()
        {
            string provider = _providerComboBox != null && _providerComboBox.SelectedItem != null
                ? _providerComboBox.SelectedItem.ToString()
                : (_settings.AssistantProvider ?? "anthropic");
            bool providerReady =
                (string.Equals(provider, "openai", StringComparison.OrdinalIgnoreCase)
                 || string.Equals(provider, "openai-compatible", StringComparison.OrdinalIgnoreCase))
                    ? !string.IsNullOrWhiteSpace(_environmentConfiguration.OpenAIApiKey)
                    : !string.IsNullOrWhiteSpace(_environmentConfiguration.AnthropicApiKey);

            _providerHealthLabel.Text = "provider: " + provider + " | model: " + (_modelTextBox.Text ?? string.Empty).Trim();
            _providerHealthLabel.ForeColor = providerReady ? Color.FromArgb(173, 231, 190) : Color.FromArgb(255, 200, 140);

            string ttsProvider = !string.IsNullOrWhiteSpace(_environmentConfiguration.ElevenLabsApiKey)
                ? "elevenlabs"
                : "local fallback";
            string sttProvider = SpeechToTextClient.GetProviderLabel(_environmentConfiguration);
            _voiceStatusLabel.Text = "voice: tts " + ttsProvider + " | stt " + sttProvider + " | speak " + (_settings.SpeakResponses ? "on" : "off");
            _voiceStatusLabel.ForeColor = Color.FromArgb(173, 191, 214);
        }

        private void UpdateActionInspector(string title, string body)
        {
            if (_actionInspectorTextBox == null)
            {
                return;
            }

            _actionInspectorTextBox.Text = string.IsNullOrWhiteSpace(body)
                ? title
                : title + Environment.NewLine + Environment.NewLine + body;
        }

        private void AddDiagnosticItem(string kind, string text)
        {
            if (_diagnosticsListBox == null)
            {
                return;
            }

            string line = DateTime.Now.ToString("HH:mm:ss") + " [" + NormalizeBadgeText(kind, 10) + "] " + NormalizeBadgeText(text, 64);
            _diagnosticsListBox.Items.Insert(0, line);
            while (_diagnosticsListBox.Items.Count > 18)
            {
                _diagnosticsListBox.Items.RemoveAt(_diagnosticsListBox.Items.Count - 1);
            }
        }

        private static string ExportDiagnosticsSnapshot(IList<string> lines)
        {
            Directory.CreateDirectory(AppSettings.StorageRoot);
            string fileName = "diagnostics-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".log";
            string path = Path.Combine(AppSettings.StorageRoot, fileName);
            File.WriteAllLines(path, lines ?? new List<string>(), Encoding.UTF8);
            return path;
        }

        private void ApplyProactiveSuggestionToPrompt()
        {
            if (_currentProactiveSuggestion == null)
            {
                SetStatus("no context idea available yet", Color.FromArgb(235, 210, 120));
                return;
            }

            _promptTextBox.Text = _currentProactiveSuggestion.Prompt ?? string.Empty;
            _promptTextBox.SelectionStart = _promptTextBox.TextLength;
            _promptTextBox.ScrollToCaret();

            if (!string.IsNullOrWhiteSpace(_currentProactiveSuggestion.CompanionMode) && _modeComboBox.Items.Contains(_currentProactiveSuggestion.CompanionMode))
            {
                _modeComboBox.SelectedItem = _currentProactiveSuggestion.CompanionMode;
            }

            SetStatus("context idea loaded into prompt", Color.FromArgb(93, 212, 136));
        }

        private async Task RunProactiveSuggestionAsync()
        {
            if (_currentProactiveSuggestion == null)
            {
                SetStatus("no context idea available yet", Color.FromArgb(235, 210, 120));
                return;
            }

            if (_currentProactiveSuggestion.ReplaySteps != null && _currentProactiveSuggestion.ReplaySteps.Count > 0)
            {
                List<ScreenCaptureInfo> screenCaptures = ScreenCaptureService.CaptureAllScreens();
                PointTagResult pointTag = new PointTagResult
                {
                    SpokenText = _currentProactiveSuggestion.Prompt,
                    ElementLabel = "context target"
                };

                ActionPlanResult actionPlan = new ActionPlanResult
                {
                    CleanText = _currentProactiveSuggestion.Prompt,
                    Steps = _currentProactiveSuggestion.ReplaySteps
                };

                await ExecuteActionPlanIfConfirmedAsync(actionPlan, pointTag, screenCaptures, _currentProactiveSuggestion.Prompt);
                return;
            }

            await RunAskFlowAsync(_currentProactiveSuggestion.Prompt);
        }

        private async Task ReplaySelectedWatchSuggestionAsync()
        {
            WatchSuggestion selectedSuggestion = _watchSuggestionComboBox.SelectedItem as WatchSuggestion;
            if (selectedSuggestion == null)
            {
                SetStatus("no learned ritual available yet", Color.FromArgb(235, 210, 120));
                return;
            }

            if (selectedSuggestion.ReplaySteps == null || selectedSuggestion.ReplaySteps.Count == 0)
            {
                SetStatus("this ritual has no replayable action chain yet", Color.FromArgb(235, 210, 120));
                return;
            }

            List<ScreenCaptureInfo> screenCaptures = ScreenCaptureService.CaptureAllScreens();
            PointTagResult pointTag = new PointTagResult
            {
                SpokenText = selectedSuggestion.Prompt,
                ElementLabel = "ritual target"
            };

            ActionPlanResult actionPlan = new ActionPlanResult
            {
                CleanText = selectedSuggestion.Prompt,
                Steps = selectedSuggestion.ReplaySteps
            };

            await ExecuteActionPlanIfConfirmedAsync(actionPlan, pointTag, screenCaptures, selectedSuggestion.Prompt);
        }

        private static string BuildRecipeName(string prompt)
        {
            string compact = Regex.Replace(prompt, @"\s+", " ").Trim();
            if (compact.Length == 0)
            {
                return "recipe";
            }

            if (compact.Length > 36)
            {
                compact = compact.Substring(0, 36).Trim();
            }

            return compact;
        }

        private void AddConversationTurn(string userPrompt, string responseText)
        {
            _conversationHistory.Add(new ConversationTurn
            {
                UserTranscript = userPrompt,
                AssistantResponse = responseText
            });

            while (_conversationHistory.Count > _settings.MaxConversationTurns)
            {
                _conversationHistory.RemoveAt(0);
            }
        }

        private Point? ConvertPointTagToScreenPoint(PointTagResult pointTag, IList<ScreenCaptureInfo> screenCaptures)
        {
            if (!pointTag.Coordinate.HasValue)
            {
                return null;
            }

            ScreenCaptureInfo targetCapture;
            if (pointTag.ScreenNumber.HasValue)
            {
                targetCapture = screenCaptures.FirstOrDefault(capture => capture.ScreenNumber == pointTag.ScreenNumber.Value);
            }
            else
            {
                targetCapture = screenCaptures.FirstOrDefault(capture => capture.IsCursorScreen);
            }

            if (targetCapture == null)
            {
                return null;
            }

            double clampedX = Math.Max(0, Math.Min(pointTag.Coordinate.Value.X, targetCapture.ScreenshotWidth));
            double clampedY = Math.Max(0, Math.Min(pointTag.Coordinate.Value.Y, targetCapture.ScreenshotHeight));

            double displayX = clampedX * (targetCapture.DisplayBounds.Width / (double)targetCapture.ScreenshotWidth);
            double displayY = clampedY * (targetCapture.DisplayBounds.Height / (double)targetCapture.ScreenshotHeight);

            return new Point(
                targetCapture.DisplayBounds.Left + (int)Math.Round(displayX),
                targetCapture.DisplayBounds.Top + (int)Math.Round(displayY)
            );
        }

        private async Task SpeakResponseAsync(string text)
        {
            if (!_settings.SpeakResponses || string.IsNullOrWhiteSpace(text))
            {
                return;
            }

            try
            {
                byte[] audioBytes = await DirectApiClient.SynthesizeSpeechAsync(_environmentConfiguration, text);
                StopAudioPlayback();
                Directory.CreateDirectory(AppSettings.StorageRoot);
                _currentAudioFilePath = Path.Combine(AppSettings.StorageRoot, "clicky-response-" + Guid.NewGuid().ToString("N") + ".mp3");
                File.WriteAllBytes(_currentAudioFilePath, audioBytes);
                PlayWithWindowsMediaPlayer(_currentAudioFilePath);
            }
            catch
            {
                SpeakLocalFallback(text);
            }
        }

        private void StopAudioPlayback()
        {
            try
            {
                if (_audioPlayer != null)
                {
                    _audioPlayer.controls.stop();
                    _audioPlayer.close();
                    _audioPlayer = null;
                }
            }
            catch
            {
                _audioPlayer = null;
            }

            try
            {
                if (!string.IsNullOrWhiteSpace(_currentAudioFilePath) && File.Exists(_currentAudioFilePath))
                {
                    File.Delete(_currentAudioFilePath);
                }
            }
            catch
            {
            }

            _currentAudioFilePath = null;
        }

        private void PlayWithWindowsMediaPlayer(string filePath)
        {
            Type mediaPlayerType = Type.GetTypeFromProgID("WMPlayer.OCX");
            if (mediaPlayerType == null)
            {
                throw new InvalidOperationException("Windows Media Player COM interface is unavailable.");
            }

            dynamic mediaPlayer = Activator.CreateInstance(mediaPlayerType);
            mediaPlayer.settings.autoStart = false;
            mediaPlayer.URL = filePath;
            mediaPlayer.controls.play();
            _audioPlayer = mediaPlayer;
        }

        private void SpeakLocalFallback(string text)
        {
            try
            {
                Type voiceType = Type.GetTypeFromProgID("SAPI.SpVoice");
                if (voiceType == null)
                {
                    return;
                }

                dynamic voice = Activator.CreateInstance(voiceType);
                voice.Speak(text, 1);
            }
            catch
            {
            }
        }

        private void ShowClickyWindow()
        {
            Show();
            WindowState = FormWindowState.Normal;
            BringToFront();
            Activate();
            _promptTextBox.Focus();
        }

        private void SetStatus(string text, Color color)
        {
            _statusLabel.Text = text;
            _statusLabel.ForeColor = color;
            UpdateStatusBadge(text, color);
            AddActivityItem("status", text);
            AddDiagnosticItem("status", text);
            Application.DoEvents();
        }

        private void BringInteractiveControlsToFront()
        {
            foreach (Control control in Controls)
            {
                if (!(control is Panel))
                {
                    control.BringToFront();
                }
            }
        }

        private void ApplyAdaptiveLayout()
        {
            if (_headerPanel == null || _controlPanel == null || _contextPanel == null || _workspacePanel == null)
            {
                return;
            }

            int margin = 18;
            int gutter = 12;
            int headerHeight = 72;
            int top = 14;
            int contentTop = top + headerHeight + 12;
            int availableWidth = Math.Max(920, ClientSize.Width - (margin * 2));
            int availableHeight = Math.Max(820, ClientSize.Height - contentTop - margin);
            int requiredControlHeight = 540;
            int requiredContextHeight = 790;
            int compactThreshold = 1180;
            _isCompactLayout = availableWidth < compactThreshold;

            if (_isCompactLayout)
            {
                int controlHeight = requiredControlHeight;
                int contextHeight = requiredContextHeight;
                int workspaceHeight = Math.Max(460, availableHeight);
                int totalHeight = contentTop + controlHeight + gutter + contextHeight + gutter + workspaceHeight + margin;

                AutoScrollMinSize = new Size(availableWidth + (margin * 2), totalHeight);
                _headerPanel.SetBounds(margin, top, availableWidth, headerHeight);
                _controlPanel.SetBounds(margin, contentTop, availableWidth, controlHeight);
                _contextPanel.SetBounds(margin, contentTop + controlHeight + gutter, availableWidth, contextHeight);
                _workspacePanel.SetBounds(margin, _contextPanel.Bottom + gutter, availableWidth, workspaceHeight);
            }
            else
            {
                int contextWidth = ClampRange((int)Math.Round(availableWidth * 0.39d), 360, 460);
                int leftWidth = Math.Max(availableWidth - contextWidth - gutter, 560);
                int workspaceHeight = ClampRange((int)Math.Round(availableHeight * 0.34d), 300, 400);
                int topRowHeight = Math.Max(Math.Max(requiredControlHeight, requiredContextHeight), availableHeight - workspaceHeight - gutter);
                int totalHeight = contentTop + topRowHeight + gutter + workspaceHeight + margin;

                AutoScrollMinSize = new Size(availableWidth + (margin * 2), totalHeight);
                _headerPanel.SetBounds(margin, top, availableWidth, headerHeight);
                _controlPanel.SetBounds(margin, contentTop, leftWidth, topRowHeight);
                _contextPanel.SetBounds(margin + leftWidth + gutter, contentTop, contextWidth, topRowHeight);
                _workspacePanel.SetBounds(margin, contentTop + topRowHeight + gutter, availableWidth, workspaceHeight);
            }

            _titleLabel.Location = new Point(margin + 8, top + 8);
            _subtitleLabel.Location = new Point(margin + 10, top + 38);
            _subtitleLabel.Width = Math.Max(320, availableWidth - 360);

            LayoutHeaderBadges(margin + availableWidth - 12, top + 10);

            _controlSectionLabel.Location = new Point(_controlPanel.Left + 16, _controlPanel.Top + 10);
            _contextSectionLabel.Location = new Point(_contextPanel.Left + 16, _contextPanel.Top + 10);
            _conversationSectionLabel.Location = new Point(_workspacePanel.Left + 16, _workspacePanel.Top + 10);

            LayoutControlDeck();
            LayoutContextDeck();
            LayoutWorkspace();
            BringInteractiveControlsToFront();
        }

        private void LayoutHeaderBadges(int rightEdge, int top)
        {
            Control[] primaryBadges = new Control[]
            {
                _riskBadgeLabel,
                _statusBadgeLabel,
                _modeBadgeLabel
            };

            int x = rightEdge;
            foreach (Control badge in primaryBadges)
            {
                x -= badge.Width;
                badge.Location = new Point(x, top);
                x -= 8;
            }

            Control[] secondaryBadges = new Control[]
            {
                _safeActionsChip,
                _ritualMemoryChip,
                _screenAwareChip
            };

            int secondaryTop = top + 34;
            int secondaryRight = rightEdge;
            foreach (Control badge in secondaryBadges)
            {
                secondaryRight -= badge.Width;
                badge.Location = new Point(secondaryRight, secondaryTop);
                secondaryRight -= 8;
            }
        }

        private void LayoutControlDeck()
        {
            int left = _controlPanel.Left + 16;
            int top = _controlPanel.Top + 42;
            int width = _controlPanel.Width - 32;
            int innerWidth = width - 20;
            int contentLeft = left + 12;
            int sectionGap = 12;
            int envHeight = 108;
            int modeHeight = 138;
            int actionHeight = 206;
            int envTop = top;
            int modeTop = envTop + envHeight + sectionGap;
            int actionTop = modeTop + modeHeight + sectionGap;

            _controlEnvPanel.SetBounds(left, envTop, width, envHeight);
            _controlEnvSectionLabel.Location = new Point(contentLeft, envTop + 6);
            _controlModePanel.SetBounds(left, modeTop, width, modeHeight);
            _controlModeSectionLabel.Location = new Point(contentLeft, modeTop + 6);
            _controlActionPanel.SetBounds(left, actionTop, width, actionHeight);
            _controlActionSectionLabel.Location = new Point(contentLeft, actionTop + 6);

            _envFileLabel.Location = new Point(contentLeft, envTop + 30);
            _envPathTextBox.SetBounds(contentLeft, envTop + 52, innerWidth, 28);
            _envStatusLabel.SetBounds(contentLeft, envTop + 84, innerWidth, 24);

            int rightColumnLeft = contentLeft + ClampRange(innerWidth / 2, 190, 260);
            int rightColumnWidth = Math.Max(180, innerWidth - (rightColumnLeft - contentLeft));
            int toggleTop = modeTop + 36;
            _speakCheckBox.Location = new Point(contentLeft, toggleTop);
            _useKnowledgeCheckBox.Location = new Point(contentLeft, toggleTop + 34);

            _modeLabel.Location = new Point(rightColumnLeft, toggleTop - 2);
            _modeComboBox.SetBounds(rightColumnLeft, toggleTop + 22, ClampRange(rightColumnWidth, 160, 230), 28);
            _suggestAutomationsCheckBox.Location = new Point(rightColumnLeft, toggleTop + 58);
            _suggestAutomationsCheckBox.Width = Math.Max(160, rightColumnWidth);

            _statusLabel.SetBounds(contentLeft, actionTop + 34, innerWidth, 24);
            _hotkeyLabel.SetBounds(contentLeft, actionTop + 60, innerWidth, 24);

            int buttonTop = actionTop + 90;
            LayoutButtonRow(new[]
            {
                _saveButton,
                _reloadEnvButton,
                _testApiButton,
                _clearHistoryButton,
                _reindexKnowledgeButton
            }, contentLeft, buttonTop, innerWidth, 10, 34);

            LayoutButtonRow(new[]
            {
                _useContextIdeaButton,
                _runContextIdeaButton,
                _saveRecipeButton,
                _runRecipeButton
            }, contentLeft, buttonTop + 44, innerWidth, 10, 34);

            LayoutButtonRow(new[]
            {
                _openRitualManagerButton,
                _openHistoryViewerButton,
                _openControlInspectorButton,
                _openProviderVoiceButton,
                _openDiagnosticsButton,
                _openSetupWizardButton
            }, contentLeft, buttonTop + 88, innerWidth, 8, 30);
        }

        private void LayoutContextDeck()
        {
            int left = _contextPanel.Left + 16;
            int top = _contextPanel.Top + 42;
            int width = _contextPanel.Width - 32;
            int innerWidth = width - 24;
            int contentLeft = left + 12;

            _contextModelPanel.SetBounds(left, top, width, 94);
            _contextModelSectionLabel.Location = new Point(contentLeft, top + 6);
            _providerLabel.Location = new Point(contentLeft, top + 30);
            _providerComboBox.SetBounds(contentLeft, top + 54, ClampRange(innerWidth / 3, 130, 170), 28);

            _modelLabel.Location = new Point(_providerComboBox.Right + 12, top + 30);
            _modelTextBox.SetBounds(_providerComboBox.Right + 12, top + 54, innerWidth - _providerComboBox.Width - 12, 28);

            int ritualTop = top + 102;
            _contextRitualPanel.SetBounds(left, ritualTop, width, 128);
            _contextRitualSectionLabel.Location = new Point(contentLeft, ritualTop + 6);
            _recipeLabel.Location = new Point(contentLeft, ritualTop + 30);
            _recipeComboBox.SetBounds(contentLeft, ritualTop + 54, innerWidth, 28);
            _learnedRitualsLabel.Location = new Point(contentLeft, ritualTop + 88);
            _watchSuggestionComboBox.SetBounds(contentLeft, ritualTop + 112, innerWidth, 28);
            LayoutButtonRow(new[] { _saveWatchIdeaButton, _useWatchIdeaButton, _replayRitualButton }, contentLeft, ritualTop + 146, innerWidth, 10, 34);

            int knowledgeTop = ritualTop + 192;
            _contextKnowledgePanel.SetBounds(left, knowledgeTop, width, 184);
            _contextKnowledgeSectionLabel.Location = new Point(contentLeft, knowledgeTop + 6);
            _knowledgeDocsLabel.Location = new Point(contentLeft, knowledgeTop + 30);
            _knowledgeSearchTextBox.SetBounds(contentLeft, knowledgeTop + 54, innerWidth, 28);
            _knowledgeListBox.SetBounds(contentLeft, knowledgeTop + 88, innerWidth, 88);
            LayoutButtonRow(new[] { _importKnowledgeButton, _refreshKnowledgeButton, _reindexSelectedKnowledgeButton, _removeKnowledgeButton }, contentLeft, knowledgeTop + 180, innerWidth, 8, 28);

            int insightTop = knowledgeTop + 220;
            _contextInsightPanel.SetBounds(left, insightTop, width, 214);
            _contextInsightSectionLabel.Location = new Point(contentLeft, insightTop + 6);
            _contextInfoPanel1.SetBounds(contentLeft, insightTop + 30, innerWidth, 40);
            _activeAppLabel.SetBounds(contentLeft + 10, insightTop + 38, innerWidth - 20, 24);
            _contextInfoPanel2.SetBounds(contentLeft, insightTop + 74, innerWidth, 40);
            _appActionsLabel.SetBounds(contentLeft + 10, insightTop + 82, innerWidth - 20, 24);
            _contextInfoPanel3.SetBounds(contentLeft, insightTop + 118, innerWidth, 40);
            _knowledgeStatusLabel.SetBounds(contentLeft + 10, insightTop + 126, innerWidth - 20, 24);
            _contextInfoPanel4.SetBounds(contentLeft, insightTop + 162, innerWidth, 44);
            _proactiveSuggestionLabel.SetBounds(contentLeft + 10, insightTop + 168, innerWidth - 20, 34);

            _previewLabel.Location = new Point(contentLeft, _contextPanel.Bottom - 146);
            _knowledgePreviewTextBox.SetBounds(contentLeft, _contextPanel.Bottom - 122, innerWidth, 106);
        }

        private void LayoutWorkspace()
        {
            int left = _workspacePanel.Left + 16;
            int top = _workspacePanel.Top + 42;
            int width = _workspacePanel.Width - 32;
            int railWidth = ClampRange((int)Math.Round(width * 0.29d), 250, 330);
            int mainWidth = width - railWidth - 12;

            _promptLabel.Location = new Point(left, top);
            _promptTextBox.SetBounds(left, top + 28, mainWidth, 106);

            int actionTop = top + 146;
            _askButton.SetBounds(left, actionTop, 196, 42);
            _dictationButton.SetBounds(_askButton.Right + 12, actionTop, 152, 42);
            _copyResponseButton.SetBounds(_dictationButton.Right + 12, actionTop, 144, 42);

            _retrievalSourcesLabel.SetBounds(left, actionTop + 50, mainWidth, 22);
            _responseLabel.Location = new Point(left, actionTop + 74);
            _responseTextBox.SetBounds(left, actionTop + 102, mainWidth, Math.Max(72, _workspacePanel.Bottom - (actionTop + 118) - 16));

            int railLeft = _workspacePanel.Right - railWidth - 16;
            _activityPanel.SetBounds(railLeft - 8, top + 12, railWidth + 16, _workspacePanel.Height - 70);
            _activityRailLabel.Location = new Point(railLeft, top + 20);
            int diagnosticsHeight = 76;
            int inspectorHeight = 88;
            int recipeHeight = Math.Max(96, _activityPanel.Height - diagnosticsHeight - inspectorHeight - 112);

            _activityListBox.SetBounds(railLeft, top + 46, railWidth, 74);
            _actionInspectorLabel.Location = new Point(railLeft, top + 128);
            _actionInspectorTextBox.SetBounds(railLeft, top + 152, railWidth, inspectorHeight);
            _recipeManagerLabel.Location = new Point(railLeft, top + 248);
            _recipeManagerListBox.SetBounds(railLeft, top + 272, railWidth, recipeHeight);
            _recipePreviewTextBox.SetBounds(railLeft, top + 276 + recipeHeight, railWidth, 78);
            _deleteRecipeButton.SetBounds(railLeft, top + 360 + recipeHeight, railWidth, 28);
            _diagnosticsLabel.Location = new Point(railLeft, top + 396 + recipeHeight);
            _diagnosticsListBox.SetBounds(railLeft, top + 420 + recipeHeight, railWidth, diagnosticsHeight);

            _providerHealthLabel.Location = new Point(left, top + 194);
            _voiceStatusLabel.Location = new Point(left, top + 216);
            _activeWindowDetailsLabel.SetBounds(left, top + 238, mainWidth, 34);
        }

        private void LayoutButtonRow(IList<Button> buttons, int left, int top, int width, int spacing, int height)
        {
            int visibleCount = buttons.Count(button => button != null);
            if (visibleCount == 0)
            {
                return;
            }

            int buttonWidth = Math.Max(66, (width - ((visibleCount - 1) * spacing)) / visibleCount);
            int currentLeft = left;
            foreach (Button button in buttons)
            {
                if (button == null)
                {
                    continue;
                }

                button.SetBounds(currentLeft, top, buttonWidth, height);
                currentLeft += buttonWidth + spacing;
            }
        }

        private void UpdateStatusBadge(string text, Color color)
        {
            if (_statusBadgeLabel == null)
            {
                return;
            }

            _statusBadgeLabel.Text = "status: " + NormalizeBadgeText(text, 18);
            _statusBadgeLabel.BackColor = Color.FromArgb(Clamp(color.R / 2 + 18), Clamp(color.G / 2 + 18), Clamp(color.B / 2 + 18));
            _statusBadgeLabel.ForeColor = Color.White;
        }

        private void UpdateRiskBadge(string level)
        {
            if (_riskBadgeLabel == null)
            {
                return;
            }

            string normalized = string.IsNullOrWhiteSpace(level) ? "low" : level.Trim().ToLowerInvariant();
            _riskBadgeLabel.Text = "risk: " + normalized;
            if (normalized == "high")
            {
                _riskBadgeLabel.BackColor = Color.FromArgb(106, 41, 41);
                _riskBadgeLabel.ForeColor = Color.FromArgb(255, 220, 220);
                return;
            }

            if (normalized == "medium")
            {
                _riskBadgeLabel.BackColor = Color.FromArgb(110, 82, 30);
                _riskBadgeLabel.ForeColor = Color.FromArgb(255, 232, 184);
                return;
            }

            _riskBadgeLabel.BackColor = Color.FromArgb(53, 73, 34);
            _riskBadgeLabel.ForeColor = Color.FromArgb(219, 244, 173);
        }

        private void UpdateModeBadge()
        {
            if (_modeBadgeLabel == null)
            {
                return;
            }

            string mode = _modeComboBox == null || _modeComboBox.SelectedItem == null
                ? _settings.CompanionMode
                : _modeComboBox.SelectedItem.ToString();
            _modeBadgeLabel.Text = "mode: " + NormalizeBadgeText(mode, 12);
        }

        private void AddActivityItem(string kind, string text)
        {
            if (_activityListBox == null)
            {
                return;
            }

            string timestamp = DateTime.Now.ToString("HH:mm");
            string line = timestamp + "  " + NormalizeBadgeText(kind, 8).PadRight(8) + "  " + NormalizeBadgeText(text, 46);
            _activityListBox.Items.Insert(0, line);
            while (_activityListBox.Items.Count > 12)
            {
                _activityListBox.Items.RemoveAt(_activityListBox.Items.Count - 1);
            }
        }

        private static string NormalizeBadgeText(string value, int maxLength)
        {
            string normalized = string.IsNullOrWhiteSpace(value) ? string.Empty : value.Replace("\r", " ").Replace("\n", " ").Trim();
            if (normalized.Length <= maxLength)
            {
                return normalized;
            }

            return normalized.Substring(0, Math.Max(0, maxLength - 1)).TrimEnd() + "…";
        }

        private static int Clamp(int value)
        {
            if (value < 0)
            {
                return 0;
            }

            if (value > 255)
            {
                return 255;
            }

            return value;
        }

        private static int ClampRange(int value, int minimum, int maximum)
        {
            if (value < minimum)
            {
                return minimum;
            }

            if (value > maximum)
            {
                return maximum;
            }

            return value;
        }

        private void EnsureCompanionOverlay()
        {
            if (_companionOverlay != null && !_companionOverlay.IsDisposed)
            {
                if (!_companionOverlay.Visible)
                {
                    _companionOverlay.Show();
                }

                return;
            }

            _companionOverlay = new CompanionOverlayForm();
            _companionOverlay.Show();
        }

        private void SetCompanionState(CompanionVisualState state)
        {
            EnsureCompanionOverlay();
            _companionOverlay.SetState(state);
        }

        private void ShowCompanionMessage(string message, CompanionVisualState state, int durationMs, CompanionVisualState stateAfterBubble)
        {
            EnsureCompanionOverlay();
            _companionOverlay.ShowMessage(message, state, durationMs, stateAfterBubble);
        }

        private void NavigateCompanionTo(Point screenPoint, string message, CompanionVisualState state, int durationMs, CompanionVisualState stateAfterNavigation)
        {
            EnsureCompanionOverlay();
            _companionOverlay.NavigateTo(screenPoint, message, state, durationMs, stateAfterNavigation);
        }

        private void UpdateSpeechButtonState()
        {
            _dictationButton.Enabled = !_isTranscribingSpeech;
            _dictationButton.Capture = _microphoneRecorder.IsRecording;

            if (_isTranscribingSpeech)
            {
                _dictationButton.Text = "transcribing...";
                _dictationButton.BackColor = Color.FromArgb(22, 30, 45);
                _dictationButton.ForeColor = Color.White;
                return;
            }

            if (_microphoneRecorder.IsRecording)
            {
                _dictationButton.Text = "release to send";
                _dictationButton.BackColor = Color.FromArgb(235, 120, 120);
                _dictationButton.ForeColor = Color.White;
                return;
            }

            _dictationButton.Text = "hold to talk";
            _dictationButton.BackColor = Color.FromArgb(22, 30, 45);
            _dictationButton.ForeColor = Color.White;
        }

        private void ConfigurePushToTalkHotKey()
        {
            Keys configuredKey = ParsePushToTalkKey(_environmentConfiguration.PushToTalkKey);

            if (_hotKeyListener != null && configuredKey == _pushToTalkKey && _hotKeyListener.IsRegistered)
            {
                _hotkeyLabel.Text = "push to talk hotkey: hold " + _pushToTalkKey.ToString().ToLowerInvariant();
                _hotkeyLabel.ForeColor = Color.FromArgb(160, 174, 192);
                return;
            }

            if (_hotKeyListener != null)
            {
                _hotKeyListener.Dispose();
                _hotKeyListener = null;
            }

            _pushToTalkKey = configuredKey;
            _hotKeyListener = new PushToTalkHotKeyListener(_pushToTalkKey);

            if (_hotKeyListener.IsRegistered)
            {
                _hotKeyListener.HotKeyPressed += delegate
                {
                    BeginInvoke((Action)delegate { HandleSpeechPress(); });
                };
                _hotKeyListener.HotKeyReleased += delegate
                {
                    BeginInvoke((Action)async delegate { await HandleSpeechReleaseAsync(); });
                };

                _hotkeyLabel.Text = "push to talk hotkey: hold " + _pushToTalkKey.ToString().ToLowerInvariant();
                _hotkeyLabel.ForeColor = Color.FromArgb(160, 174, 192);
            }
            else
            {
                _hotkeyLabel.Text = "push to talk hotkey unavailable for " + _pushToTalkKey.ToString().ToLowerInvariant();
                _hotkeyLabel.ForeColor = Color.FromArgb(235, 120, 120);
            }
        }

        private static Keys ParsePushToTalkKey(string configuredKey)
        {
            if (string.IsNullOrWhiteSpace(configuredKey))
            {
                return Keys.F8;
            }

            Keys parsedKey;
            if (Enum.TryParse(configuredKey, true, out parsedKey))
            {
                return parsedKey;
            }

            return Keys.F8;
        }

        private void OnMainFormClosing(object sender, FormClosingEventArgs e)
        {
            if (!_quitRequested)
            {
                e.Cancel = true;
                Hide();
                return;
            }

            StopAudioPlayback();
            _microphoneRecorder.Dispose();
            if (_companionOverlay != null)
            {
                _companionOverlay.Close();
                _companionOverlay.Dispose();
            }
            if (_hotKeyListener != null)
            {
                _hotKeyListener.Dispose();
            }
            _notifyIcon.Visible = false;
            _notifyIcon.Dispose();
            _trayMenu.Dispose();
        }

        private static void TryDeleteFile(string filePath)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(filePath) && File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
            }
            catch
            {
            }
        }

        private static Label CreateLabel(string text, int x, int y, int width, int height, Font font = null, Color? color = null)
        {
            return new Label
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(width, height),
                Font = font ?? new Font("Segoe UI", 10.0f),
                ForeColor = color ?? Color.White,
                BackColor = Color.Transparent
            };
        }

        private static Label CreateSectionLabel(string text, int x, int y, int width)
        {
            return new Label
            {
                Text = text.ToUpperInvariant(),
                Location = new Point(x, y),
                Size = new Size(width, 20),
                Font = new Font("Segoe UI Semibold", 8.5f),
                ForeColor = Color.FromArgb(116, 151, 188),
                BackColor = Color.Transparent
            };
        }

        private static Label CreateChip(string text, int x, int y, int width, Color backColor, Color foreColor)
        {
            return new Label
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(width, 24),
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI Semibold", 8.5f),
                ForeColor = foreColor,
                BackColor = backColor
            };
        }

        private static TextBox CreateTextBox(int x, int y, int width, int height, string text)
        {
            return new TextBox
            {
                Location = new Point(x, y),
                Size = new Size(width, height),
                Text = text,
                Font = new Font("Segoe UI", 10.0f),
                BackColor = Color.FromArgb(20, 29, 44),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };
        }

        private static Button CreateButton(string text, int x, int y, int width, int height, Color backColor, Color foreColor)
        {
            Button button = new Button
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(width, height),
                BackColor = backColor,
                ForeColor = foreColor,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI Semibold", 9.5f),
                Cursor = Cursors.Hand
            };
            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.MouseDownBackColor = ControlPaint.Dark(backColor, 0.12f);
            button.FlatAppearance.MouseOverBackColor = ControlPaint.Light(backColor, 0.08f);
            return button;
        }

        private static Panel CreateCardPanel(int x, int y, int width, int height, Color backColor)
        {
            Panel panel = new Panel
            {
                Location = new Point(x, y),
                Size = new Size(width, height),
                BackColor = backColor
            };
            panel.Paint += delegate(object sender, PaintEventArgs e)
            {
                Rectangle borderRectangle = new Rectangle(0, 0, panel.Width - 1, panel.Height - 1);
                using (Pen borderPen = new Pen(Color.FromArgb(46, 80, 118), 1.0f))
                {
                    e.Graphics.DrawRectangle(borderPen, borderRectangle);
                }
            };
            return panel;
        }
    }
}
