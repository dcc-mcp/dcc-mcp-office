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

    public static OfficeComBackend Create(string app, StaDispatcher sta) => app.ToLowerInvariant() switch
    {
        "powerpoint" => new PowerPointBackend(sta),
        "word" => new WordBackend(sta),
        "excel" => new ExcelBackend(sta),
        _ => throw new OfficeComException(OfficeErrorCode.OfficeCapabilityUnsupported,
            $"no COM backend for app '{app}'"),
    };
}
