namespace CarolusNexus.Core;

/// <summary>
/// Maps Windows process names to operator <see cref="ActiveWindowInfo.AppKind"/> labels for rituals, guards, and Live Context.
/// Includes Dynamics AX, PTC Creo, Babtec B4/BCT line, CATIA, and generic fallbacks.
/// </summary>
public static class AppKindDetector
{
    public static string FromProcessName(string? processName)
    {
        var p = (processName ?? string.Empty).Trim().ToLowerInvariant();
        if (p.Length == 0)
        {
            return "generic";
        }

        return p switch
        {
            "ax32" => "ax",
            "chrome" or "msedge" or "firefox" or "brave" or "opera" => "browser",
            "explorer" => "explorer",
            "code" or "devenv" or "rider64" or "idea64" or "pycharm64" => "ide",
            "outlook" or "olk" => "mail",
            "slack" or "teams" or "discord" or "telegram" => "messenger",
            _ => FromProcessNameHeuristic(p)
        };
    }

    private static string FromProcessNameHeuristic(string p)
    {
        // PTC Creo / Pro/E family (common host processes)
        if (p.Contains("parametric", StringComparison.Ordinal) ||
            p is "xtop" or "xtp" ||
            p.Contains("creoagent", StringComparison.Ordinal) ||
            p.Contains("creopma", StringComparison.Ordinal) ||
            (p.Contains("creo", StringComparison.Ordinal) && !p.Contains("creator", StringComparison.Ordinal)))
        {
            return "creo";
        }

        // Babtec Software (B4, BCT clients — names vary by product/version)
        if (p.Contains("babtec", StringComparison.Ordinal) ||
            p.StartsWith("bct", StringComparison.Ordinal) ||
            p.Contains("b4win", StringComparison.Ordinal) ||
            p.Contains("bab4", StringComparison.Ordinal) ||
            p.Contains("b4client", StringComparison.Ordinal))
        {
            return "babtec";
        }

        // Dassault CATIA (typical process CNEXT)
        if (p is "cnext" or "cnext3d" ||
            p.Contains("catia", StringComparison.Ordinal) ||
            p.Contains("cnext", StringComparison.Ordinal))
        {
            return "catia";
        }

        // Siemens NX (optional — common host)
        if (p is "ugraf" || p.Contains("nxlauncher", StringComparison.Ordinal))
        {
            return "nx";
        }

        return "generic";
    }
}
