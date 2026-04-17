using System.Diagnostics;
using System.Text;
using ClippyRWAvalonia.Models;

namespace ClippyRWAvalonia.Services;

public sealed class LocalAgentRunService
{
    private readonly OperatorWorkspaceService _workspaceService;

    public LocalAgentRunService(OperatorWorkspaceService workspaceService)
    {
        _workspaceService = workspaceService;
    }

    public async Task<AgentRunResult> RunAsync(string agent, string prompt)
    {
        var env = _workspaceService.ReadEnvFile();
        var normalizedAgent = (agent ?? string.Empty).Trim().ToLowerInvariant();
        var repoRoot = _workspaceService.RepoRoot;
        var outputDirectory = Path.Combine(repoRoot, "codex output");
        Directory.CreateDirectory(outputDirectory);

        var workingDirectory = ResolveWorkingDirectory(env, repoRoot);
        Directory.CreateDirectory(workingDirectory);

        var timestamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
        var outputFilePath = Path.Combine(outputDirectory, $"karl-klammer-{normalizedAgent}-{timestamp}.txt");

        var startInfo = BuildStartInfo(normalizedAgent, env, prompt, outputFilePath, workingDirectory);
        using var process = new Process { StartInfo = startInfo };
        if (!process.Start())
        {
            throw new InvalidOperationException($"Failed to start {normalizedAgent}.");
        }

        if (normalizedAgent is "codex" or "claude-code")
        {
            var promptBytes = new UTF8Encoding(false).GetBytes(prompt ?? string.Empty);
            await process.StandardInput.BaseStream.WriteAsync(promptBytes, 0, promptBytes.Length).ConfigureAwait(false);
            await process.StandardInput.BaseStream.FlushAsync().ConfigureAwait(false);
            process.StandardInput.Close();
        }

        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();
        var timeoutSeconds = ResolveTimeoutSeconds(normalizedAgent, env);
        var exited = await Task.Run(() => process.WaitForExit(timeoutSeconds * 1000)).ConfigureAwait(false);
        if (!exited)
        {
            try { process.Kill(); } catch { }
            throw new InvalidOperationException($"{normalizedAgent} timed out before finishing.");
        }

        var stdout = await stdoutTask.ConfigureAwait(false);
        var stderr = await stderrTask.ConfigureAwait(false);
        WriteOutputFile(outputFilePath, normalizedAgent, workingDirectory, prompt ?? string.Empty, process.ExitCode, stdout, stderr);

        if (process.ExitCode != 0)
        {
            throw new InvalidOperationException(string.IsNullOrWhiteSpace(stderr)
                ? $"{normalizedAgent} failed. Check the output file: {outputFilePath}"
                : stderr.Trim());
        }

        var responseText = string.IsNullOrWhiteSpace(stdout)
            ? $"{normalizedAgent} session finished."
            : stdout.Trim();

        return new AgentRunResult
        {
            Agent = normalizedAgent,
            Prompt = prompt ?? string.Empty,
            WorkingDirectory = workingDirectory,
            OutputFilePath = outputFilePath,
            ExitCode = process.ExitCode,
            StandardOutput = stdout,
            StandardError = stderr,
            ResponseText = responseText
        };
    }

    private ProcessStartInfo BuildStartInfo(string agent, Dictionary<string, string> env, string prompt, string outputFilePath, string workingDirectory)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = "cmd.exe",
            WorkingDirectory = agent == "openclaw" ? _workspaceService.RepoRoot : workingDirectory,
            UseShellExecute = false,
            RedirectStandardInput = agent is "codex" or "claude-code",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        AddArgument(startInfo, "/c");
        switch (agent)
        {
            case "codex":
                AddArgument(startInfo, GetValueOrDefault(env, "CODEX_COMMAND", "codex.cmd"));
                AddArgument(startInfo, "exec");
                AddArgument(startInfo, "--full-auto");
                AddArgument(startInfo, "--skip-git-repo-check");
                AddArgument(startInfo, "-C");
                AddArgument(startInfo, workingDirectory);
                AddArgument(startInfo, "-o");
                AddArgument(startInfo, outputFilePath);
                AddArgument(startInfo, "-");
                break;
            case "claude-code":
                AddArgument(startInfo, GetValueOrDefault(env, "CLAUDE_CODE_COMMAND", "claude"));
                AddArgument(startInfo, "-p");
                AddArgument(startInfo, "--permission-mode");
                AddArgument(startInfo, "bypassPermissions");
                break;
            case "openclaw":
                AddArgument(startInfo, GetValueOrDefault(env, "OPENCLAW_COMMAND", "openclaw"));
                AddArgument(startInfo, "agent");
                AddArgument(startInfo, "--agent");
                AddArgument(startInfo, ResolveOpenClawAgentId(env));
                AddArgument(startInfo, "--message");
                AddArgument(startInfo, prompt ?? string.Empty);
                AddArgument(startInfo, "--timeout");
                AddArgument(startInfo, ResolveTimeoutSeconds(agent, env).ToString());
                break;
            default:
                throw new InvalidOperationException($"Unsupported agent '{agent}'.");
        }

        return startInfo;
    }

    private string ResolveWorkingDirectory(Dictionary<string, string> env, string repoRoot)
    {
        var configured = GetValueOrDefault(env, "CODEX_WORKDIR", Path.Combine(repoRoot, "playground"));
        return configured;
    }

    private int ResolveTimeoutSeconds(string agent, Dictionary<string, string> env)
    {
        var key = agent == "openclaw" ? "OPENCLAW_TIMEOUT_SECONDS" : "CODEX_TIMEOUT_SECONDS";
        return int.TryParse(GetValueOrDefault(env, key, agent == "openclaw" ? "120" : "900"), out var parsed) && parsed > 0
            ? parsed
            : (agent == "openclaw" ? 120 : 900);
    }

    private string ResolveOpenClawAgentId(Dictionary<string, string> env)
    {
        var configured = GetValueOrDefault(env, "OPENCLAW_SESSION_KEY", "main").Trim();
        var parts = configured.Split(':', StringSplitOptions.RemoveEmptyEntries);
        return parts.Length >= 2 && string.Equals(parts[0], "agent", StringComparison.OrdinalIgnoreCase)
            ? parts[1]
            : configured;
    }

    private static string GetValueOrDefault(Dictionary<string, string> values, string key, string defaultValue)
    {
        return values.TryGetValue(key, out var value) && !string.IsNullOrWhiteSpace(value) ? value : defaultValue;
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

    private static void WriteOutputFile(string outputFilePath, string agent, string workingDirectory, string prompt, int exitCode, string standardOutput, string standardError)
    {
        var builder = new StringBuilder();
        builder.AppendLine($"Karl Klammer {agent} run");
        builder.AppendLine("timestamp: " + DateTime.Now.ToString("O"));
        builder.AppendLine("working_directory: " + workingDirectory);
        builder.AppendLine("exit_code: " + exitCode);
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
