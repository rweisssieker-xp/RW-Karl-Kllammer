namespace CarolusNexus.Core;

public static class WorkspacePathResolver
{
    public static WorkspaceLayout CreateFromDefaultBase()
    {
        var repoRoot = ResolveRepoRoot();
        return Create(repoRoot);
    }

    public static WorkspaceLayout Create(string repoRoot)
    {
        var windowsRoot = Path.Combine(repoRoot, "windows");
        var dataRoot = Path.Combine(windowsRoot, "data");
        var knowledgeRoot = Path.Combine(dataRoot, "knowledge");
        var envPath = ResolvePreferredEnvPath(windowsRoot);
        return new WorkspaceLayout(
            RepoRoot: repoRoot,
            WindowsRoot: windowsRoot,
            DataRoot: dataRoot,
            KnowledgeRoot: knowledgeRoot,
            EnvPath: envPath,
            SettingsPath: Path.Combine(dataRoot, "settings.json"),
            RecipePath: Path.Combine(dataRoot, "automation-recipes.json"),
            WatchSessionsPath: Path.Combine(dataRoot, "watch-sessions.json"),
            ActionHistoryPath: Path.Combine(dataRoot, "action-history.json"),
            KnowledgeIndexPath: Path.Combine(dataRoot, "knowledge-index.json"));
    }

    public static string ResolveRepoRoot() =>
        TryFindRepoRoot(AppContext.BaseDirectory)
        ?? Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));

    private static string? TryFindRepoRoot(string current)
    {
        for (var i = 0; i < 8 && !string.IsNullOrWhiteSpace(current); i++)
        {
            var candidate = Path.GetFullPath(Path.Combine(current, string.Concat(Enumerable.Repeat("..\\", i))));
            if (Directory.Exists(Path.Combine(candidate, "windows")))
            {
                return candidate;
            }
        }

        return null;
    }

    private static string ResolvePreferredEnvPath(string windowsRoot)
    {
        var preferred = Path.Combine(windowsRoot, ".env");
        if (File.Exists(preferred))
        {
            return preferred;
        }

        var publishCandidate = Path.Combine(windowsRoot, "bin", "Release", "net10.0-windows", "win-arm64", ".env");
        return File.Exists(publishCandidate) ? publishCandidate : preferred;
    }
}
