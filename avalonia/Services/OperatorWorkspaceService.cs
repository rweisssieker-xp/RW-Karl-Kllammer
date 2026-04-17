using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ClippyRWAvalonia.Models;

namespace ClippyRWAvalonia.Services;

public sealed class OperatorWorkspaceService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true
    };

    private static readonly string[] SupportedKnowledgeExtensions = [".txt", ".md", ".log", ".json", ".csv", ".pdf", ".docx"];

    public OperatorWorkspaceService()
    {
        RepoRoot = FindRepoRoot();
        WindowsRoot = Path.Combine(RepoRoot, "windows");
        DataRoot = Path.Combine(WindowsRoot, "data");
        KnowledgeRoot = Path.Combine(DataRoot, "knowledge");
        EnvPath = ResolvePreferredEnvPath();
        SettingsPath = Path.Combine(DataRoot, "settings.json");
        RecipePath = Path.Combine(DataRoot, "automation-recipes.json");
        WatchSessionsPath = Path.Combine(DataRoot, "watch-sessions.json");
        ActionHistoryPath = Path.Combine(DataRoot, "action-history.json");
        KnowledgeIndexPath = Path.Combine(DataRoot, "knowledge-index.json");
    }

    public string RepoRoot { get; }
    public string WindowsRoot { get; }
    public string DataRoot { get; }
    public string KnowledgeRoot { get; }
    public string EnvPath { get; }
    public string SettingsPath { get; }
    public string RecipePath { get; }
    public string WatchSessionsPath { get; }
    public string ActionHistoryPath { get; }
    public string KnowledgeIndexPath { get; }

    public OperatorWorkspaceSnapshot Load()
    {
        Directory.CreateDirectory(DataRoot);
        Directory.CreateDirectory(KnowledgeRoot);

        var settings = ReadStringDictionary(SettingsPath);
        var envValues = ReadEnvFile();
        var documents = GetKnowledgeDocuments();
        var recipes = ReadRecipes(RecipePath);
        var watchSessions = ReadJsonList<WatchSessionEntry>(WatchSessionsPath)
            .OrderByDescending(entry => entry.TimestampUtc)
            .ToList();
        var history = ReadJsonList<ActionHistoryEntry>(ActionHistoryPath)
            .OrderByDescending(entry => entry.TimestampUtc)
            .ToList();
        var diagnostics = GetDiagnosticsEntries();

        return new OperatorWorkspaceSnapshot
        {
            RepoRoot = RepoRoot,
            WindowsRoot = WindowsRoot,
            DataRoot = DataRoot,
            EnvPath = EnvPath,
            EnvExists = File.Exists(EnvPath),
            Provider = settings.TryGetValue("AssistantProvider", out var provider) ? provider : "anthropic",
            Model = settings.TryGetValue("AssistantModel", out var model) ? model : string.Empty,
            Mode = settings.TryGetValue("CompanionMode", out var mode) ? mode : "companion",
            SpeakResponses = ReadBool(settings, "SpeakResponses"),
            UseLocalKnowledge = ReadBool(settings, "UseLocalKnowledge"),
            SuggestAutomations = ReadBool(settings, "SuggestAutomations"),
            RuntimeSummary = BuildRuntimeSummary(File.Exists(EnvPath), documents.Count, history.Count, watchSessions.Count, diagnostics.Count),
            KnowledgeStatus = GetKnowledgeStatusText(documents.Count),
            KnowledgeDocuments = documents,
            Recipes = recipes,
            WatchSessions = watchSessions,
            ActionHistory = history,
            Diagnostics = diagnostics,
            Settings = settings,
            EnvValues = envValues
        };
    }

    public void SaveSettings(string provider, string model, string mode, bool speakResponses, bool useLocalKnowledge, bool suggestAutomations)
    {
        Directory.CreateDirectory(DataRoot);
        var settings = ReadStringDictionary(SettingsPath);
        settings["AssistantProvider"] = provider;
        settings["AssistantModel"] = model;
        settings["ClaudeModel"] = model;
        settings["CompanionMode"] = mode;
        settings["SpeakResponses"] = speakResponses ? "true" : "false";
        settings["UseLocalKnowledge"] = useLocalKnowledge ? "true" : "false";
        settings["SuggestAutomations"] = suggestAutomations ? "true" : "false";
        File.WriteAllText(SettingsPath, JsonSerializer.Serialize(settings, JsonOptions), Encoding.UTF8);
    }

    public IReadOnlyList<KnowledgeDocumentSummary> GetKnowledgeDocuments(string searchTerm = "")
    {
        Directory.CreateDirectory(KnowledgeRoot);
        var index = LoadKnowledgeIndex();
        var chunkCounts = (index?.Chunks ?? [])
            .Where(chunk => !string.IsNullOrWhiteSpace(chunk.SourcePath))
            .GroupBy(chunk => chunk.SourcePath, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.OrdinalIgnoreCase);

        var docs = Directory.GetFiles(KnowledgeRoot, "*.*", SearchOption.AllDirectories)
            .Where(path => SupportedKnowledgeExtensions.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase))
            .Select(path => new FileInfo(path))
            .OrderByDescending(file => file.LastWriteTimeUtc)
            .Select(file => new KnowledgeDocumentSummary
            {
                Title = file.Name,
                RelativePath = Path.GetRelativePath(KnowledgeRoot, file.FullName),
                ChunkCount = chunkCounts.TryGetValue(file.FullName, out var count) ? count : 0,
                LastWriteUtc = file.LastWriteTimeUtc.ToString("o"),
                Extension = file.Extension
            })
            .ToList();

        if (string.IsNullOrWhiteSpace(searchTerm))
        {
            return docs;
        }

        var normalizedSearch = searchTerm.Trim();
        var matchingPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (index?.Chunks != null)
        {
            var tokens = Tokenize(normalizedSearch);
            foreach (var chunk in index.Chunks)
            {
                var haystack = $"{chunk.Title} {chunk.Text}".ToLowerInvariant();
                if (haystack.Contains(normalizedSearch.ToLowerInvariant()) ||
                    tokens.Any(token => token.Length >= 3 && haystack.Contains(token)))
                {
                    matchingPaths.Add(chunk.SourcePath);
                }
            }
        }

        return docs
            .Where(doc =>
                doc.Title.Contains(normalizedSearch, StringComparison.OrdinalIgnoreCase) ||
                doc.RelativePath.Contains(normalizedSearch, StringComparison.OrdinalIgnoreCase) ||
                matchingPaths.Contains(Path.GetFullPath(Path.Combine(KnowledgeRoot, doc.RelativePath))))
            .ToList();
    }

    public string GetKnowledgePreview(string relativePath, int maxLength = 5000)
    {
        var fullPath = ResolveKnowledgePath(relativePath);
        if (fullPath == null)
        {
            return string.Empty;
        }

        var text = ReadKnowledgeText(fullPath);
        if (string.IsNullOrWhiteSpace(text))
        {
            return "No readable preview available for this document yet.";
        }

        var compact = Regex.Replace(text, @"\s+", " ").Trim();
        if (compact.Length <= maxLength)
        {
            return compact;
        }

        return compact[..maxLength].Trim() + "...";
    }

    public int ImportKnowledgeFiles(IEnumerable<string> sourcePaths)
    {
        Directory.CreateDirectory(KnowledgeRoot);
        var imported = 0;
        foreach (var sourcePath in sourcePaths)
        {
            if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
            {
                continue;
            }

            var extension = Path.GetExtension(sourcePath);
            if (!SupportedKnowledgeExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
            {
                continue;
            }

            var targetPath = EnsureUniqueTargetPath(Path.Combine(KnowledgeRoot, Path.GetFileName(sourcePath)));
            File.Copy(sourcePath, targetPath, overwrite: false);
            imported++;
        }

        return imported;
    }

    public bool DeleteKnowledgeDocument(string relativePath)
    {
        var fullPath = ResolveKnowledgePath(relativePath);
        if (fullPath == null || !File.Exists(fullPath))
        {
            return false;
        }

        File.Delete(fullPath);
        return true;
    }

    public int ReindexKnowledge()
    {
        Directory.CreateDirectory(KnowledgeRoot);
        Directory.CreateDirectory(DataRoot);
        var chunks = new List<KnowledgeChunk>();

        foreach (var filePath in Directory.GetFiles(KnowledgeRoot, "*.*", SearchOption.AllDirectories))
        {
            if (!SupportedKnowledgeExtensions.Contains(Path.GetExtension(filePath), StringComparer.OrdinalIgnoreCase))
            {
                continue;
            }

            var text = ReadKnowledgeText(filePath);
            if (string.IsNullOrWhiteSpace(text))
            {
                continue;
            }

            foreach (var chunkText in ChunkText(text, 900))
            {
                chunks.Add(new KnowledgeChunk
                {
                    SourcePath = filePath,
                    Title = Path.GetFileName(filePath),
                    Text = chunkText
                });
            }
        }

        var index = new KnowledgeIndex
        {
            Chunks = chunks,
            IndexedAtUtc = DateTime.UtcNow.ToString("o")
        };
        File.WriteAllText(KnowledgeIndexPath, JsonSerializer.Serialize(index, JsonOptions), Encoding.UTF8);
        return chunks.Count;
    }

    public IReadOnlyList<KnowledgeChunk> GetRelevantKnowledgeChunks(string prompt, int maxChunks = 3)
    {
        var index = LoadKnowledgeIndex();
        if (index?.Chunks == null || index.Chunks.Count == 0 || string.IsNullOrWhiteSpace(prompt))
        {
            return [];
        }

        var normalizedPrompt = prompt.Trim().ToLowerInvariant();
        var tokens = Tokenize(normalizedPrompt);

        return index.Chunks
            .Select(chunk => new
            {
                Chunk = chunk,
                Score = ScoreChunk(chunk, normalizedPrompt, tokens)
            })
            .Where(entry => entry.Score > 0)
            .OrderByDescending(entry => entry.Score)
            .ThenBy(entry => entry.Chunk.Title, StringComparer.OrdinalIgnoreCase)
            .Take(Math.Max(1, maxChunks))
            .Select(entry => entry.Chunk)
            .ToList();
    }

    public void SaveRecipes(IEnumerable<AutomationRecipe> recipes)
    {
        Directory.CreateDirectory(DataRoot);
        File.WriteAllText(RecipePath, JsonSerializer.Serialize(recipes, JsonOptions), Encoding.UTF8);
    }

    public void SaveActionHistory(IEnumerable<ActionHistoryEntry> entries)
    {
        Directory.CreateDirectory(DataRoot);
        File.WriteAllText(ActionHistoryPath, JsonSerializer.Serialize(entries, JsonOptions), Encoding.UTF8);
    }

    public IReadOnlyList<DiagnosticEntry> GetDiagnosticsEntries()
    {
        Directory.CreateDirectory(DataRoot);
        var entries = new List<DiagnosticEntry>();
        foreach (var file in Directory.GetFiles(DataRoot, "diagnostics-*.log", SearchOption.TopDirectoryOnly).OrderByDescending(path => path))
        {
            foreach (var line in File.ReadAllLines(file, Encoding.UTF8))
            {
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                entries.Add(new DiagnosticEntry
                {
                    SourceFile = Path.GetFileName(file),
                    Line = line.Trim(),
                    TimestampHint = TryExtractTimestamp(line)
                });
            }
        }

        return entries;
    }

    public string ExportDiagnostics(IEnumerable<DiagnosticEntry> entries)
    {
        Directory.CreateDirectory(DataRoot);
        var path = Path.Combine(DataRoot, $"diagnostics-{DateTime.UtcNow:yyyyMMdd-HHmmss}.log");
        var lines = entries.Select(entry => entry.Line).ToArray();
        File.WriteAllLines(path, lines, Encoding.UTF8);
        return path;
    }

    public void ClearDiagnosticsLogs()
    {
        Directory.CreateDirectory(DataRoot);
        foreach (var file in Directory.GetFiles(DataRoot, "diagnostics-*.log", SearchOption.TopDirectoryOnly))
        {
            File.Delete(file);
        }
    }

    public Dictionary<string, string> ReadEnvFile()
    {
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (!File.Exists(EnvPath))
        {
            return values;
        }

        foreach (var rawLine in File.ReadAllLines(EnvPath, Encoding.UTF8))
        {
            var line = rawLine.Trim();
            if (line.Length == 0 || line.StartsWith("#", StringComparison.Ordinal))
            {
                continue;
            }

            var separatorIndex = line.IndexOf('=');
            if (separatorIndex <= 0)
            {
                continue;
            }

            var key = line[..separatorIndex].Trim();
            var value = line[(separatorIndex + 1)..].Trim().Trim('"');
            values[key] = value;
        }

        return values;
    }

    public ActiveWindowInfo GetActiveWindow()
    {
        if (!OperatingSystem.IsWindows())
        {
            return new ActiveWindowInfo
            {
                ProcessName = "unsupported",
                WindowTitle = "active window inspection is only available on Windows",
                AppKind = "unsupported",
                DesktopFramework = "unsupported"
            };
        }

        try
        {
            var handle = GetForegroundWindow();
            if (handle == IntPtr.Zero)
            {
                return new ActiveWindowInfo();
            }

            GetWindowThreadProcessId(handle, out var processId);
            var processName = "unknown app";
            if (processId != 0)
            {
                using var process = System.Diagnostics.Process.GetProcessById((int)processId);
                processName = process.ProcessName;
            }

            var titleBuilder = new StringBuilder(512);
            var classBuilder = new StringBuilder(256);
            GetWindowText(handle, titleBuilder, titleBuilder.Capacity);
            GetClassName(handle, classBuilder, classBuilder.Capacity);
            GetWindowRect(handle, out var rect);

            return new ActiveWindowInfo
            {
                ProcessName = processName,
                WindowTitle = titleBuilder.ToString().Trim(),
                WindowClassName = classBuilder.ToString().Trim(),
                AppKind = DetectAppKind(processName),
                DesktopFramework = DetectDesktopFramework(classBuilder.ToString().Trim(), processName),
                Left = rect.Left,
                Top = rect.Top,
                Width = Math.Max(0, rect.Right - rect.Left),
                Height = Math.Max(0, rect.Bottom - rect.Top)
            };
        }
        catch
        {
            return new ActiveWindowInfo();
        }
    }

    private Dictionary<string, string> ReadStringDictionary(string path)
    {
        try
        {
            if (!File.Exists(path))
            {
                return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }

            using var document = JsonDocument.Parse(File.ReadAllText(path, Encoding.UTF8));
            return document.RootElement.EnumerateObject().ToDictionary(
                property => property.Name,
                property => property.Value.ValueKind switch
                {
                    JsonValueKind.String => property.Value.GetString() ?? string.Empty,
                    JsonValueKind.True => "true",
                    JsonValueKind.False => "false",
                    _ => property.Value.ToString()
                },
                StringComparer.OrdinalIgnoreCase);
        }
        catch
        {
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }
    }

    private static bool ReadBool(Dictionary<string, string> settings, string key)
    {
        return settings.TryGetValue(key, out var value) && bool.TryParse(value, out var parsed) && parsed;
    }

    private List<T> ReadJsonList<T>(string path)
    {
        try
        {
            if (!File.Exists(path))
            {
                return [];
            }

            return JsonSerializer.Deserialize<List<T>>(File.ReadAllText(path, Encoding.UTF8), JsonOptions) ?? [];
        }
        catch
        {
            return [];
        }
    }

    private List<AutomationRecipe> ReadRecipes(string path)
    {
        try
        {
            if (!File.Exists(path))
            {
                return [];
            }

            var json = File.ReadAllText(path, Encoding.UTF8);
            var recipes = JsonSerializer.Deserialize<List<AutomationRecipe>>(json, JsonOptions) ?? [];
            foreach (var recipe in recipes)
            {
                MigrateRecipe(recipe);
            }

            return recipes;
        }
        catch
        {
            return [];
        }
    }

    private static void MigrateRecipe(AutomationRecipe recipe)
    {
        recipe.Id = string.IsNullOrWhiteSpace(recipe.Id) ? Guid.NewGuid().ToString("n") : recipe.Id;
        recipe.Description ??= string.Empty;
        recipe.Category = string.IsNullOrWhiteSpace(recipe.Category) ? InferRecipeCategory(recipe.Prompt) : recipe.Category;
        recipe.SourceType = string.IsNullOrWhiteSpace(recipe.SourceType) ? "manual" : recipe.SourceType;
        recipe.RiskLevel = string.IsNullOrWhiteSpace(recipe.RiskLevel) ? InferRecipeRisk(recipe) : recipe.RiskLevel;
        recipe.KnowledgeSources ??= [];
        recipe.Tags ??= [];
        recipe.Parameters ??= [];
        recipe.Steps ??= [];
        if (recipe.Steps.Count == 0)
        {
            recipe.Steps = ParseLegacyPromptToSteps(recipe.Prompt);
        }

        if (!recipe.Tags.Contains(recipe.Category, StringComparer.OrdinalIgnoreCase))
        {
            recipe.Tags.Add(recipe.Category);
        }

        if (recipe.Category.Equals("ax", StringComparison.OrdinalIgnoreCase) && string.IsNullOrWhiteSpace(recipe.GuardApp))
        {
            recipe.GuardApp = "ax";
        }
    }

    private static List<RitualStep> ParseLegacyPromptToSteps(string prompt)
    {
        if (string.IsNullOrWhiteSpace(prompt))
        {
            return [];
        }

        var lines = prompt.Replace("\r\n", "\n").Split('\n', StringSplitOptions.RemoveEmptyEntries)
            .Select(line => line.Trim())
            .Where(line => line.Length > 0)
            .ToList();

        var directSteps = new List<RitualStep>();
        foreach (var line in lines)
        {
            if (line.StartsWith("category:", StringComparison.OrdinalIgnoreCase) ||
                line.StartsWith("guard app:", StringComparison.OrdinalIgnoreCase) ||
                line.StartsWith("guard form:", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (line.StartsWith("ax.", StringComparison.OrdinalIgnoreCase))
            {
                directSteps.Add(new RitualStep
                {
                    ActionType = "ax",
                    ActionArgument = line,
                    RiskLevel = InferStepRisk(line)
                });
                continue;
            }

            if (line.StartsWith("app|", StringComparison.OrdinalIgnoreCase))
            {
                var plan = AssistantActionPlan.Parse("[ACTIONS:" + line + "]");
                directSteps.AddRange(plan.Steps.Select(step => new RitualStep
                {
                    ActionType = step.ActionName,
                    ActionArgument = step.ActionArgument,
                    WaitMs = step.WaitMilliseconds,
                    RetryCount = step.RetryCount,
                    IfApp = step.RequiredAppContains,
                    IfForm = step.RequiredFormContains,
                    IfDialog = step.RequiredDialogContains,
                    IfTab = step.RequiredTabContains,
                    OnFail = step.OnFail,
                    RiskLevel = InferStepRisk(step.ActionArgument)
                }));
            }
        }

        if (directSteps.Count > 0)
        {
            return directSteps;
        }

        return prompt.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(part => part.Length > 0)
            .Select(part => new RitualStep
            {
                ActionType = part.StartsWith("ax.", StringComparison.OrdinalIgnoreCase) ? "ax" : "app",
                ActionArgument = part,
                RiskLevel = InferStepRisk(part)
            })
            .ToList();
    }

    private static string InferRecipeCategory(string prompt)
    {
        return prompt.Contains("ax.", StringComparison.OrdinalIgnoreCase) ? "ax" : "general";
    }

    private static string InferRecipeRisk(AutomationRecipe recipe)
    {
        if (recipe.Steps.Any(step => InferStepRisk(step.ActionArgument) == "high"))
        {
            return "high";
        }

        if (recipe.Steps.Any(step => InferStepRisk(step.ActionArgument) == "medium"))
        {
            return "medium";
        }

        return "low";
    }

    private static string InferStepRisk(string action)
    {
        var normalized = (action ?? string.Empty).Trim().ToLowerInvariant();
        if (normalized.Contains("post", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("delete", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("book", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("save", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("confirm", StringComparison.OrdinalIgnoreCase))
        {
            return "high";
        }

        if (normalized.Contains("set_field", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("type_control", StringComparison.OrdinalIgnoreCase) ||
            normalized.Contains("click_action", StringComparison.OrdinalIgnoreCase))
        {
            return "medium";
        }

        return "low";
    }

    private KnowledgeIndex? LoadKnowledgeIndex()
    {
        try
        {
            if (!File.Exists(KnowledgeIndexPath))
            {
                return null;
            }

            return JsonSerializer.Deserialize<KnowledgeIndex>(File.ReadAllText(KnowledgeIndexPath, Encoding.UTF8), JsonOptions);
        }
        catch
        {
            return null;
        }
    }

    private string? ResolveKnowledgePath(string relativePath)
    {
        if (string.IsNullOrWhiteSpace(relativePath))
        {
            return null;
        }

        var fullPath = Path.GetFullPath(Path.Combine(KnowledgeRoot, relativePath));
        var rootPath = Path.GetFullPath(KnowledgeRoot);
        return fullPath.StartsWith(rootPath, StringComparison.OrdinalIgnoreCase) ? fullPath : null;
    }

    private static IEnumerable<string> ChunkText(string text, int maxChunkLength)
    {
        var normalized = Regex.Replace(text ?? string.Empty, @"\r\n?", "\n").Trim();
        while (normalized.Length > 0)
        {
            var length = Math.Min(maxChunkLength, normalized.Length);
            var splitIndex = normalized.LastIndexOf('\n', length - 1, length);
            if (splitIndex < maxChunkLength / 3)
            {
                splitIndex = length;
            }

            var chunk = normalized[..splitIndex].Trim();
            if (!string.IsNullOrWhiteSpace(chunk))
            {
                yield return chunk;
            }

            normalized = splitIndex >= normalized.Length ? string.Empty : normalized[splitIndex..].Trim();
        }
    }

    private static string BuildRuntimeSummary(bool envExists, int docs, int actionHistory, int watchSessions, int diagnostics)
    {
        var parts = new List<string>
        {
            envExists ? "env ready" : "env missing",
            $"{docs} docs",
            $"{actionHistory} actions",
            $"{watchSessions} watch sessions",
            $"{diagnostics} diagnostics"
        };
        return string.Join(" • ", parts);
    }

    private string GetKnowledgeStatusText(int docs)
    {
        if (!File.Exists(KnowledgeIndexPath))
        {
            return "knowledge: no index yet";
        }

        var index = LoadKnowledgeIndex();
        var count = index?.Chunks?.Count ?? 0;
        return $"knowledge: {docs} documents • {count} chunks indexed";
    }

    private string ReadKnowledgeText(string path)
    {
        var extension = Path.GetExtension(path).ToLowerInvariant();
        return extension switch
        {
            ".txt" or ".md" or ".log" or ".json" or ".csv" => SafeReadText(path),
            ".docx" => ReadDocxText(path),
            ".pdf" => $"PDF document: {Path.GetFileName(path)}",
            _ => string.Empty
        };
    }

    private static string SafeReadText(string path)
    {
        try
        {
            return File.ReadAllText(path, Encoding.UTF8);
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string ReadDocxText(string path)
    {
        try
        {
            using var archive = ZipFile.OpenRead(path);
            var entry = archive.GetEntry("word/document.xml");
            if (entry == null)
            {
                return string.Empty;
            }

            using var stream = entry.Open();
            using var reader = new StreamReader(stream, Encoding.UTF8);
            var xml = reader.ReadToEnd();
            var document = XDocument.Parse(xml);
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            return string.Join(" ", document.Descendants(w + "t").Select(element => element.Value));
        }
        catch
        {
            return string.Empty;
        }
    }

    private string EnsureUniqueTargetPath(string targetPath)
    {
        if (!File.Exists(targetPath))
        {
            return targetPath;
        }

        var directory = Path.GetDirectoryName(targetPath) ?? KnowledgeRoot;
        var fileName = Path.GetFileNameWithoutExtension(targetPath);
        var extension = Path.GetExtension(targetPath);
        var counter = 2;
        while (true)
        {
            var candidate = Path.Combine(directory, $"{fileName}-{counter}{extension}");
            if (!File.Exists(candidate))
            {
                return candidate;
            }

            counter++;
        }
    }

    private string ResolvePreferredEnvPath()
    {
        var preferred = Path.Combine(WindowsRoot, ".env");
        if (File.Exists(preferred))
        {
            return preferred;
        }

        var publishCandidate = Path.Combine(WindowsRoot, "bin", "Release", "net10.0-windows", "win-arm64", ".env");
        return File.Exists(publishCandidate) ? publishCandidate : preferred;
    }

    private static string[] Tokenize(string text)
    {
        return Regex.Split(text.ToLowerInvariant(), @"[^a-z0-9äöüß]+")
            .Where(token => token.Length >= 3)
            .Distinct()
            .ToArray();
    }

    private static int ScoreChunk(KnowledgeChunk chunk, string normalizedPrompt, string[] tokens)
    {
        var haystack = $"{chunk.Title} {chunk.Text}".ToLowerInvariant();
        var score = 0;
        if (haystack.Contains(normalizedPrompt, StringComparison.Ordinal))
        {
            score += 12;
        }

        foreach (var token in tokens)
        {
            if (haystack.Contains(token, StringComparison.Ordinal))
            {
                score += token.Length >= 6 ? 4 : 2;
            }
        }

        return score;
    }

    private static string TryExtractTimestamp(string line)
    {
        var match = Regex.Match(line, @"\d{4}-\d{2}-\d{2}[T ][0-9:\.\-+Z]+");
        return match.Success ? match.Value : string.Empty;
    }

    private static string FindRepoRoot()
    {
        var current = AppContext.BaseDirectory;
        for (var i = 0; i < 8 && !string.IsNullOrWhiteSpace(current); i++)
        {
            var candidate = Path.GetFullPath(Path.Combine(current, string.Concat(Enumerable.Repeat("..\\", i))));
            if (Directory.Exists(Path.Combine(candidate, "windows")))
            {
                return candidate;
            }
        }

        return Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
    }

    private static string DetectAppKind(string processName)
    {
        return (processName ?? string.Empty).Trim().ToLowerInvariant() switch
        {
            "ax32" => "ax",
            "chrome" or "msedge" or "firefox" or "brave" or "opera" => "browser",
            "explorer" => "explorer",
            "code" or "devenv" or "rider64" or "idea64" or "pycharm64" => "ide",
            "outlook" or "olk" => "mail",
            "slack" or "teams" or "discord" or "telegram" => "messenger",
            _ => "generic"
        };
    }

    private static string DetectDesktopFramework(string windowClassName, string processName)
    {
        var normalizedClass = (windowClassName ?? string.Empty).Trim();
        var normalizedProcess = (processName ?? string.Empty).Trim().ToLowerInvariant();

        if (normalizedClass.StartsWith("WindowsForms10", StringComparison.OrdinalIgnoreCase))
        {
            return "winforms";
        }

        if (normalizedClass.StartsWith("HwndWrapper", StringComparison.OrdinalIgnoreCase))
        {
            return "wpf";
        }

        if (normalizedClass.StartsWith("SunAwt", StringComparison.OrdinalIgnoreCase) || normalizedProcess.Contains("java", StringComparison.OrdinalIgnoreCase))
        {
            return "java";
        }

        if (normalizedClass.StartsWith("Qt", StringComparison.OrdinalIgnoreCase))
        {
            return "qt";
        }

        return "classic";
    }

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder text, int count);

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }
}
