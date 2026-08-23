using Office.Automation.Runtime;

namespace Office.Automation.Com;

/// <summary>
/// Maps an app name to its COM backend (proposal §8.2: one sidecar process
/// per Office application).
/// </summary>
public static class ComBackendFactory
{
    /// <summary>Apps this M1 COM MVP drives.</summary>
    public static readonly string[] SupportedApps = { "powerpoint", "word", "excel" };

    public static bool IsSupported(string app) =>
        SupportedApps.Contains(app, StringComparer.OrdinalIgnoreCase);

    /// <summary>Registry-only availability probe; never launches Office.</summary>
    public static bool IsInstalled(string app)
    {
        if (!IsSupported(app))
        {
            return false;
        }
        string progId = app.ToLowerInvariant() switch
        {
            "powerpoint" => "PowerPoint.Application",
            "word" => "Word.Application",
            "excel" => "Excel.Application",
            _ => throw new InvalidOperationException("supported app has no ProgID"),
        };
        try
        {
            return Type.GetTypeFromProgID(progId, throwOnError: false) is not null;
        }
        catch
        {
            return false;
        }
    }

    public static OfficeComBackend Create(
        string app,
        StaDispatcher sta,
        TimeSpan? requestTimeout = null,
        int timeoutStreakForRecovery = 2) => app.ToLowerInvariant() switch
        {
            "powerpoint" => new PowerPointBackend(sta, requestTimeout, timeoutStreakForRecovery),
            "word" => new WordBackend(sta, requestTimeout, timeoutStreakForRecovery),
            "excel" => new ExcelBackend(sta, requestTimeout, timeoutStreakForRecovery),
            _ => throw new OfficeComException(OfficeErrorCode.OfficeCapabilityUnsupported,
                $"no COM backend for app '{app}'"),
        };
}
