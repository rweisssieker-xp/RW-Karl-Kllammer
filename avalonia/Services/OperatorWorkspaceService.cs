using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using CarolusNexus.Core;
using CarolusNexus.Platform.Windows;
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
        var layout = WorkspacePathResolver.CreateFromDefaultBase();
        RepoRoot = layout.RepoRoot;
        WindowsRoot = layout.WindowsRoot;
        DataRoot = layout.DataRoot;
        KnowledgeRoot = layout.KnowledgeRoot;
        EnvPath = layout.EnvPath;
        SettingsPath = layout.SettingsPath;
        RecipePath = layout.RecipePath;
        WatchSessionsPath = layout.WatchSessionsPath;
        ActionHistoryPath = layout.ActionHistoryPath;
        KnowledgeIndexPath = layout.KnowledgeIndexPath;
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
            AutoRouteLocalAgents = ReadBoolWithDefault(settings, "AutoRouteLocalAgents", defaultValue: true),
            SpeakAfterAsk = ReadBoolWithDefault(settings, "SpeakAfterAsk", defaultValue: false),
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

    public void SaveSettings(
        string provider,
        string model,
        string mode,
        bool speakResponses,
        bool useLocalKnowledge,
        bool suggestAutomations,
        bool autoRouteLocalAgents,
        bool speakAfterAsk)
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
        settings["AutoRouteLocalAgents"] = autoRouteLocalAgents ? "true" : "false";
        settings["SpeakAfterAsk"] = speakAfterAsk ? "true" : "false";
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

    /// <summary>ZIP with redacted context for support: settings, env key names only, filtered diagnostics, runtime info. No .env values.</summary>
    public string ExportSupportBundle(IReadOnlyList<DiagnosticEntry> diagnosticsForExport, int maxDiagnosticLines = 500)
    {
        Directory.CreateDirectory(DataRoot);
        var zipPath = Path.Combine(DataRoot, $"support-bundle-{DateTime.UtcNow:yyyyMMdd-HHmmss}.zip");
        var envKeys = ReadEnvFile().Keys.OrderBy(k => k, StringComparer.OrdinalIgnoreCase).ToList();
        var asm = Assembly.GetExecutingAssembly().GetName();
        var appInfo =
            $"Carolus Nexus operator export{Environment.NewLine}" +
            $"timestamp_utc: {DateTime.UtcNow:o}{Environment.NewLine}" +
            $"assembly: {asm.Name} {asm.Version}{Environment.NewLine}" +
            $"os: {Environment.OSVersion}{Environment.NewLine}" +
            $"repo_root: {RepoRoot}{Environment.NewLine}" +
            $"env_path_exists: {File.Exists(EnvPath)}{Environment.NewLine}";

        var diagLines = diagnosticsForExport
            .Take(Math.Max(0, maxDiagnosticLines))
            .Select(e => e.ToString());

        using (var fs = File.Create(zipPath))
        using (var zip = new ZipArchive(fs, ZipArchiveMode.Create))
        {
            static void AddUtf8(ZipArchive z, string name, string text)
            {
                var entry = z.CreateEntry(name);
                using var w = new StreamWriter(entry.Open(), new UTF8Encoding(false));
                w.Write(text);
            }

            AddUtf8(zip, "readme.txt",
                "Support bundle for Carolus Nexus.\n" +
                "- settings.json: local operator settings (no API secrets).\n" +
                "- env-variable-names.txt: keys present in .env only; values are NOT included.\n" +
                "- diagnostics-export.txt: lines from the current diagnostics list in the UI.\n" +
                "- runtime.txt: build and OS summary.\n");
            AddUtf8(zip, "runtime.txt", appInfo);
            AddUtf8(zip, "env-variable-names.txt", string.Join(Environment.NewLine, envKeys));
            AddUtf8(zip, "diagnostics-export.txt", string.Join(Environment.NewLine, diagLines));

            if (File.Exists(SettingsPath))
            {
                AddUtf8(zip, "settings.json", File.ReadAllText(SettingsPath, Encoding.UTF8));
            }
            else
            {
                AddUtf8(zip, "settings.json", "{}\n");
            }
        }

        return zipPath;
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

    public ActiveWindowInfo GetActiveWindow() => WindowsForegroundWindow.GetActiveWindow();

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

    private static bool ReadBoolWithDefault(Dictionary<string, string> settings, string key, bool defaultValue)
    {
        if (!settings.TryGetValue(key, out var value) || string.IsNullOrWhiteSpace(value))
        {
            return defaultValue;
        }

        return bool.TryParse(value, out var parsed) && parsed;
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

}
