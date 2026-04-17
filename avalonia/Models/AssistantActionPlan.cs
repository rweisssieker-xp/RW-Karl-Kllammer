using System.Text.RegularExpressions;

namespace ClippyRWAvalonia.Models;

public sealed class AssistantActionPlan
{
    private static readonly Regex ActionPlanRegex = new(@"\s*\[ACTIONS:([^\]]+)\]\s*", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex PointRegex = new(@"\s*\[POINT:(\d+)\|(\d{1,3})\|(\d{1,3})\|([^\]]+)\]\s*", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public string CleanText { get; set; } = string.Empty;
    public List<AssistantActionStep> Steps { get; set; } = [];
    public AssistantPointTag? PointTag { get; set; }

    public static AssistantActionPlan Parse(string responseText)
    {
        var result = new AssistantActionPlan();
        var intermediate = ParsePointTag(responseText ?? string.Empty, result);
        result.CleanText = ParseActionSteps(intermediate, result).Trim();
        return result;
    }

    private static string ParsePointTag(string text, AssistantActionPlan result)
    {
        return PointRegex.Replace(text, match =>
        {
            if (result.PointTag != null)
            {
                return " ";
            }

            if (!int.TryParse(match.Groups[1].Value, out var screenIndex) ||
                !int.TryParse(match.Groups[2].Value, out var xPercent) ||
                !int.TryParse(match.Groups[3].Value, out var yPercent))
            {
                return " ";
            }

            result.PointTag = new AssistantPointTag
            {
                ScreenIndex = Math.Max(1, screenIndex),
                XPercent = Math.Clamp(xPercent, 0, 100),
                YPercent = Math.Clamp(yPercent, 0, 100),
                Label = match.Groups[4].Value.Trim()
            };
            return " ";
        });
    }

    private static string ParseActionSteps(string text, AssistantActionPlan result)
    {
        return ActionPlanRegex.Replace(text, match =>
        {
            var payload = match.Groups[1].Value;
            foreach (var rawStep in payload.Split([';'], StringSplitOptions.RemoveEmptyEntries))
            {
                var trimmedStep = rawStep.Trim();
                if (trimmedStep.Length == 0)
                {
                    continue;
                }

                var tokens = trimmedStep.Split(['|'], StringSplitOptions.None);
                var actionName = string.Empty;
                var actionArgument = string.Empty;
                var waitMilliseconds = 0;
                var retryCount = 0;
                var requiredAppContains = string.Empty;
                var requiredFormContains = string.Empty;
                var requiredDialogContains = string.Empty;
                var requiredTabContains = string.Empty;
                var onFail = "stop";
                var argumentTokens = new List<string>();

                for (var i = 0; i < tokens.Length; i++)
                {
                    var token = (tokens[i] ?? string.Empty).Trim();
                    if (token.Length == 0)
                    {
                        continue;
                    }

                    if (token.StartsWith("wait=", StringComparison.OrdinalIgnoreCase))
                    {
                        if (int.TryParse(token[5..], out var parsedWait) && parsedWait > 0)
                        {
                            waitMilliseconds = parsedWait;
                        }

                        continue;
                    }

                    if (token.StartsWith("ifapp=", StringComparison.OrdinalIgnoreCase))
                    {
                        requiredAppContains = token[6..].Trim();
                        continue;
                    }

                    if (token.StartsWith("if_form=", StringComparison.OrdinalIgnoreCase))
                    {
                        requiredFormContains = token["if_form=".Length..].Trim();
                        continue;
                    }

                    if (token.StartsWith("if_dialog=", StringComparison.OrdinalIgnoreCase))
                    {
                        requiredDialogContains = token["if_dialog=".Length..].Trim();
                        continue;
                    }

                    if (token.StartsWith("if_tab=", StringComparison.OrdinalIgnoreCase))
                    {
                        requiredTabContains = token["if_tab=".Length..].Trim();
                        continue;
                    }

                    if (token.StartsWith("retry=", StringComparison.OrdinalIgnoreCase))
                    {
                        if (int.TryParse(token[6..], out var parsedRetry) && parsedRetry > 0)
                        {
                            retryCount = Math.Min(parsedRetry, 5);
                        }

                        continue;
                    }

                    if (token.StartsWith("on_fail=", StringComparison.OrdinalIgnoreCase))
                    {
                        var requested = token["on_fail=".Length..].Trim().ToLowerInvariant();
                        onFail = requested is "skip" or "continue" ? "skip" : "stop";
                        continue;
                    }

                    if (actionName.Length == 0)
                    {
                        actionName = token.ToLowerInvariant();
                        continue;
                    }

                    argumentTokens.Add(token);
                }

                actionArgument = string.Join("|", argumentTokens).Trim();

                if (actionName is "app")
                {
                    result.Steps.Add(new AssistantActionStep
                    {
                        ActionName = actionName,
                        ActionArgument = actionArgument,
                        WaitMilliseconds = waitMilliseconds,
                        RequiredAppContains = requiredAppContains,
                        RequiredFormContains = requiredFormContains,
                        RequiredDialogContains = requiredDialogContains,
                        RequiredTabContains = requiredTabContains,
                        RetryCount = retryCount,
                        OnFail = onFail
                    });
                }
            }

            return " ";
        });
    }
}

public sealed class AssistantActionStep
{
    public string ActionName { get; set; } = string.Empty;
    public string ActionArgument { get; set; } = string.Empty;
    public int WaitMilliseconds { get; set; }
    public int RetryCount { get; set; }
    public string RequiredAppContains { get; set; } = string.Empty;
    public string RequiredFormContains { get; set; } = string.Empty;
    public string RequiredDialogContains { get; set; } = string.Empty;
    public string RequiredTabContains { get; set; } = string.Empty;
    public string OnFail { get; set; } = "stop";
}

public sealed class AssistantPointTag
{
    public int ScreenIndex { get; set; }
    public int XPercent { get; set; }
    public int YPercent { get; set; }
    public string Label { get; set; } = string.Empty;
}
