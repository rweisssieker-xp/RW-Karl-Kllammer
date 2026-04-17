using System.Text.RegularExpressions;

namespace CarolusNexus.Core;

/// <summary>Detects German CLI handoff phrases and IDE-style auto-routing (legacy WinForms parity).</summary>
public static class AgentHandoffTriggers
{
    private static readonly Regex CodexTriggerRegex = new(
        @"\b(?:nimm|nim|nehm|nehm|mit)\s+(?:den\s+)?(?:codex|kodex|kodes|codecs|kodexx)\b[\s,:-]*",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly Regex CodexWithScreenTriggerRegex = new(
        @"\b(?:nimm|nim|nehm|mit)\s+(?:den\s+)?(?:codex|kodex|kodes)\s+(?:mit|mids?|plus)\s+(?:screen|screenshot|bild|main\s*screen|hauptbildschirm|hauptscreen)\b[\s,:-]*",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly Regex ClaudeCodeTriggerRegex = new(
        @"\b(?:nimm|nim|nehm|mit)\s+(?:den\s+)?(?:claude|cloud|clod|klod|klode)\s+code\b[\s,:-]*",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly Regex OpenClawTriggerRegex = new(
        @"\b(?:nimm|nim|nehm|mit)\s+(?:den\s+)?(?:(?:open|oben|orpen|onpen|oppen)\s*cl(?:aw|au)|openclaw|klaus|claus|claws)\b[\s,:-]*",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    public static bool IsOpenClawTriggered(string prompt) =>
        !string.IsNullOrWhiteSpace(prompt) && OpenClawTriggerRegex.IsMatch(NormalizePrompt(prompt));

    public static bool IsClaudeCodeTriggered(string prompt) =>
        !string.IsNullOrWhiteSpace(prompt) && ClaudeCodeTriggerRegex.IsMatch(NormalizeClaudeCodePrompt(prompt));

    public static bool IsCodexTriggered(string prompt)
    {
        if (string.IsNullOrWhiteSpace(prompt))
        {
            return false;
        }

        var normalized = NormalizePrompt(prompt);
        return CodexWithScreenTriggerRegex.IsMatch(normalized) || CodexTriggerRegex.IsMatch(normalized);
    }

    public static bool ShouldAttachCodexScreens(string prompt) =>
        !string.IsNullOrWhiteSpace(prompt) && CodexWithScreenTriggerRegex.IsMatch(NormalizePrompt(prompt));

    public static string RemoveOpenClawPrompt(string prompt) =>
        string.IsNullOrWhiteSpace(prompt) ? string.Empty : OpenClawTriggerRegex.Replace(NormalizePrompt(prompt), string.Empty, 1).Trim();

    public static string RemoveClaudeCodePrompt(string prompt) =>
        string.IsNullOrWhiteSpace(prompt) ? string.Empty : ClaudeCodeTriggerRegex.Replace(NormalizeClaudeCodePrompt(prompt), string.Empty, 1).Trim();

    public static string RemoveCodexPrompt(string prompt)
    {
        if (string.IsNullOrWhiteSpace(prompt))
        {
            return string.Empty;
        }

        var normalized = NormalizePrompt(prompt);
        normalized = CodexWithScreenTriggerRegex.Replace(normalized, string.Empty, 1).Trim();
        return CodexTriggerRegex.Replace(normalized, string.Empty, 1).Trim();
    }

    /// <summary>Returns "codex", "openclaw", or empty — same heuristics as legacy WinForms IntentRouter.</summary>
    public static string DetectIntentRoute(string prompt, ActiveWindowInfo activeWindow)
    {
        var normalizedPrompt = (prompt ?? string.Empty).Trim().ToLowerInvariant();
        if (normalizedPrompt.Length == 0)
        {
            return string.Empty;
        }

        var appKind = activeWindow.AppKind ?? "generic";
        var looksLikeCodingTask =
            string.Equals(appKind, "ide", StringComparison.OrdinalIgnoreCase)
            || normalizedPrompt.Contains("fix ")
            || normalizedPrompt.Contains("debug")
            || normalizedPrompt.Contains("refactor")
            || normalizedPrompt.Contains("implement")
            || normalizedPrompt.Contains("write code")
            || normalizedPrompt.Contains("change the code")
            || normalizedPrompt.Contains("run tests")
            || normalizedPrompt.Contains("build failed")
            || normalizedPrompt.Contains("compiler error");

        var looksLikeAgentTask =
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

    private static string NormalizePrompt(string prompt)
    {
        var normalized = prompt.ToLowerInvariant();
        normalized = normalized.Replace("kodex", "codex");
        normalized = normalized.Replace("kodes", "codex");
        normalized = normalized.Replace("codecs", "codex");
        normalized = normalized.Replace("codexx", "codex");
        normalized = normalized.Replace("nehm", "nimm");
        normalized = Regex.Replace(normalized, @"\s+", " ").Trim();
        return normalized;
    }

    private static string NormalizeClaudeCodePrompt(string prompt)
    {
        var normalized = prompt.ToLowerInvariant();
        normalized = normalized.Replace("cloud code", "claude code");
        normalized = normalized.Replace("clod code", "claude code");
        normalized = normalized.Replace("klod code", "claude code");
        normalized = normalized.Replace("klode code", "claude code");
        normalized = normalized.Replace("nehm", "nimm");
        normalized = Regex.Replace(normalized, @"\s+", " ").Trim();
        return normalized;
    }
}
