using ClippyRWAvalonia.Models;

namespace ClippyRWAvalonia.Services;

public sealed class RitualRuntimeService
{
    private readonly OperatorWorkspaceService _workspaceService;
    private readonly DesktopInspectorService _desktopInspectorService;
    private readonly AxClientAutomationService _axClientAutomationService;

    public RitualRuntimeService(
        OperatorWorkspaceService workspaceService,
        DesktopInspectorService desktopInspectorService,
        AxClientAutomationService axClientAutomationService)
    {
        _workspaceService = workspaceService;
        _desktopInspectorService = desktopInspectorService;
        _axClientAutomationService = axClientAutomationService;
    }

    public async Task<RitualRunResult> DryRunAsync(AutomationRecipe recipe, int startIndex = 0, int count = int.MaxValue)
    {
        var steps = SliceSteps(recipe, startIndex, count);
        var log = steps.Select((step, index) => $"step {startIndex + index + 1}: {step}").ToList();
        return await Task.FromResult(new RitualRunResult
        {
            Status = "dry-run",
            Summary = steps.Count == 0 ? "no ritual steps to simulate" : $"simulated {steps.Count} ritual step(s)",
            LogLines = log,
            NextStepIndex = Math.Min(recipe.Steps.Count, startIndex + steps.Count)
        }).ConfigureAwait(false);
    }

    public async Task<RitualRunResult> RunAsync(
        AutomationRecipe recipe,
        Dictionary<string, string>? parameterValues = null,
        int startIndex = 0,
        int count = int.MaxValue,
        CancellationToken cancellationToken = default)
    {
        var values = MergeParameterValues(recipe, parameterValues);
        var active = _workspaceService.GetActiveWindow();
        var axContext = _axClientAutomationService.CaptureActiveContext();
        var steps = SliceSteps(recipe, startIndex, count);
        var log = new List<string>();
        var nextIndex = startIndex;
        var status = "completed";
        var lastResult = string.Empty;

        if (!ValidateRecipeGuards(recipe, active, axContext, out var guardFailure))
        {
            return new RitualRunResult
            {
                Status = "blocked",
                Summary = guardFailure,
                LogLines = [guardFailure],
                NextStepIndex = startIndex
            };
        }

        foreach (var indexed in steps.Select((step, offset) => new { Step = step, Index = startIndex + offset }))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var step = indexed.Step;
            nextIndex = indexed.Index;

            if (IsSensitiveBlocked(active, step, recipe, out var sensitiveFailure))
            {
                log.Add($"step {indexed.Index + 1}: blocked {step.ActionArgument} ({sensitiveFailure})");
                return new RitualRunResult
                {
                    Status = "blocked",
                    Summary = sensitiveFailure,
                    LogLines = log,
                    NextStepIndex = indexed.Index,
                    LastResult = lastResult
                };
            }

            if (!ValidateStepGuards(step, active, axContext, out var stepGuardFailure))
            {
                log.Add($"step {indexed.Index + 1}: skip {step.ActionArgument} ({stepGuardFailure})");
                if (string.Equals(step.OnFail, "stop", StringComparison.OrdinalIgnoreCase))
                {
                    status = "blocked";
                    return new RitualRunResult
                    {
                        Status = status,
                        Summary = stepGuardFailure,
                        LogLines = log,
                        NextStepIndex = indexed.Index,
                        LastResult = lastResult
                    };
                }

                nextIndex = indexed.Index + 1;
                continue;
            }

            if (step.WaitMs > 0)
            {
                log.Add($"step {indexed.Index + 1}: wait {step.WaitMs}ms");
                await Task.Delay(step.WaitMs, cancellationToken).ConfigureAwait(false);
            }

            var resolvedAction = ResolveAction(step, values);
            var attempts = Math.Max(1, step.RetryCount + 1);
            var success = false;
            for (var attempt = 1; attempt <= attempts; attempt++)
            {
                lastResult = step.ActionType.Equals("ax", StringComparison.OrdinalIgnoreCase) ||
                             resolvedAction.StartsWith("ax.", StringComparison.OrdinalIgnoreCase)
                    ? _axClientAutomationService.Execute(resolvedAction, _axClientAutomationService.CaptureActiveContext())
                    : _desktopInspectorService.Execute(resolvedAction);

                if (IsSuccessfulActionResult(lastResult))
                {
                    success = true;
                    break;
                }

                if (attempt < attempts)
                {
                    await Task.Delay(220, cancellationToken).ConfigureAwait(false);
                }
            }

            log.Add($"step {indexed.Index + 1}: {resolvedAction} -> {(success ? lastResult : "failed")}");
            nextIndex = indexed.Index + 1;
            if (!success)
            {
                status = string.Equals(step.OnFail, "skip", StringComparison.OrdinalIgnoreCase) ? "partial" : "error";
                if (status == "error")
                {
                    return new RitualRunResult
                    {
                        Status = status,
                        Summary = BuildFailureSummary(step, lastResult),
                        LogLines = log,
                        NextStepIndex = indexed.Index,
                        LastResult = lastResult
                    };
                }
            }
        }

        return new RitualRunResult
        {
            Status = status,
            Summary = log.Count == 0 ? "no ritual steps executed" : $"executed {log.Count} ritual step(s)",
            LogLines = log,
            NextStepIndex = nextIndex,
            LastResult = lastResult
        };
    }

    private static List<RitualStep> SliceSteps(AutomationRecipe recipe, int startIndex, int count)
    {
        if (recipe.Steps.Count == 0 || startIndex >= recipe.Steps.Count)
        {
            return [];
        }

        return recipe.Steps.Skip(Math.Max(0, startIndex)).Take(Math.Max(1, count)).ToList();
    }

    private static Dictionary<string, string> MergeParameterValues(AutomationRecipe recipe, Dictionary<string, string>? parameterValues)
    {
        var merged = recipe.Parameters
            .Where(parameter => !string.IsNullOrWhiteSpace(parameter.Name))
            .ToDictionary(parameter => parameter.Name, parameter => parameter.DefaultValue ?? string.Empty, StringComparer.OrdinalIgnoreCase);

        if (parameterValues == null)
        {
            return merged;
        }

        foreach (var pair in parameterValues)
        {
            merged[pair.Key] = pair.Value;
        }

        return merged;
    }

    private static bool ValidateRecipeGuards(AutomationRecipe recipe, ActiveWindowInfo active, AxContextSnapshot axContext, out string failure)
    {
        if (!string.IsNullOrWhiteSpace(recipe.GuardApp) &&
            !active.AppKind.Contains(recipe.GuardApp, StringComparison.OrdinalIgnoreCase))
        {
            failure = $"ritual guard app mismatch: expected {recipe.GuardApp}";
            return false;
        }

        if (!string.IsNullOrWhiteSpace(recipe.GuardForm) &&
            !axContext.FormName.Contains(recipe.GuardForm, StringComparison.OrdinalIgnoreCase))
        {
            failure = $"ritual guard form mismatch: expected {recipe.GuardForm}";
            return false;
        }

        if (!string.IsNullOrWhiteSpace(recipe.GuardDialog) &&
            !axContext.DialogTitle.Contains(recipe.GuardDialog, StringComparison.OrdinalIgnoreCase))
        {
            failure = $"ritual guard dialog mismatch: expected {recipe.GuardDialog}";
            return false;
        }

        if (!string.IsNullOrWhiteSpace(recipe.GuardTab) &&
            !axContext.ActiveTab.Contains(recipe.GuardTab, StringComparison.OrdinalIgnoreCase))
        {
            failure = $"ritual guard tab mismatch: expected {recipe.GuardTab}";
            return false;
        }

        failure = string.Empty;
        return true;
    }

    private static bool ValidateStepGuards(RitualStep step, ActiveWindowInfo active, AxContextSnapshot axContext, out string failure)
    {
        if (!string.IsNullOrWhiteSpace(step.IfApp) &&
            !active.AppKind.Contains(step.IfApp, StringComparison.OrdinalIgnoreCase) &&
            !active.DisplayName.Contains(step.IfApp, StringComparison.OrdinalIgnoreCase))
        {
            failure = $"ifapp mismatch: {step.IfApp}";
            return false;
        }

        if (!string.IsNullOrWhiteSpace(step.IfForm) &&
            !axContext.FormName.Contains(step.IfForm, StringComparison.OrdinalIgnoreCase))
        {
            failure = $"if_form mismatch: {step.IfForm}";
            return false;
        }

        if (!string.IsNullOrWhiteSpace(step.IfDialog) &&
            !axContext.DialogTitle.Contains(step.IfDialog, StringComparison.OrdinalIgnoreCase))
        {
            failure = $"if_dialog mismatch: {step.IfDialog}";
            return false;
        }

        if (!string.IsNullOrWhiteSpace(step.IfTab) &&
            !axContext.ActiveTab.Contains(step.IfTab, StringComparison.OrdinalIgnoreCase))
        {
            failure = $"if_tab mismatch: {step.IfTab}";
            return false;
        }

        failure = string.Empty;
        return true;
    }

    private static bool IsSensitiveBlocked(ActiveWindowInfo active, RitualStep step, AutomationRecipe recipe, out string failure)
    {
        var action = (step.ActionArgument ?? string.Empty).Trim().ToLowerInvariant();
        var app = (active.AppKind ?? string.Empty).Trim().ToLowerInvariant();
        var riskySend = action.Contains("send", StringComparison.OrdinalIgnoreCase) ||
                        action.Contains("post", StringComparison.OrdinalIgnoreCase) ||
                        action.Contains("book", StringComparison.OrdinalIgnoreCase);
        if (riskySend && (app == "mail" || app == "messenger"))
        {
            failure = $"safety zone blocked sensitive action in {app}";
            return true;
        }

        if (riskySend && string.Equals(recipe.RiskLevel, "high", StringComparison.OrdinalIgnoreCase))
        {
            failure = "high-risk posting/sending action requires manual confirmation";
            return true;
        }

        failure = string.Empty;
        return false;
    }

    private static string ResolveAction(RitualStep step, Dictionary<string, string> parameterValues)
    {
        var action = step.ActionArgument ?? string.Empty;
        foreach (var pair in parameterValues)
        {
            action = action.Replace("{{" + pair.Key + "}}", pair.Value ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        }

        foreach (var binding in step.ParameterBindings)
        {
            if (parameterValues.TryGetValue(binding.Key, out var value))
            {
                action = action.Replace(binding.Value, value ?? string.Empty, StringComparison.OrdinalIgnoreCase);
            }
        }

        return action;
    }

    private static bool IsSuccessfulActionResult(string? result)
    {
        if (string.IsNullOrWhiteSpace(result))
        {
            return false;
        }

        return !result.Contains("not found", StringComparison.OrdinalIgnoreCase) &&
               !result.Contains("failed", StringComparison.OrdinalIgnoreCase) &&
               !result.Contains("unsupported", StringComparison.OrdinalIgnoreCase) &&
               !result.Contains("mismatch", StringComparison.OrdinalIgnoreCase) &&
               !result.Contains("not an AX client", StringComparison.OrdinalIgnoreCase);
    }

    private static string BuildFailureSummary(RitualStep step, string result)
    {
        if (step.ActionArgument.StartsWith("ax.select_grid_row:", StringComparison.OrdinalIgnoreCase))
        {
            return "grid row not found";
        }

        if (step.ActionArgument.StartsWith("ax.set_field:", StringComparison.OrdinalIgnoreCase))
        {
            return "field not found or not writable";
        }

        if (step.ActionArgument.StartsWith("ax.open_lookup:", StringComparison.OrdinalIgnoreCase))
        {
            return "lookup unresolved";
        }

        return string.IsNullOrWhiteSpace(result) ? "ritual step failed" : result;
    }
}
