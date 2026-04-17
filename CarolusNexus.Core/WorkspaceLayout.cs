namespace CarolusNexus.Core;

public sealed record WorkspaceLayout(
    string RepoRoot,
    string WindowsRoot,
    string DataRoot,
    string KnowledgeRoot,
    string EnvPath,
    string SettingsPath,
    string RecipePath,
    string WatchSessionsPath,
    string ActionHistoryPath,
    string KnowledgeIndexPath);
