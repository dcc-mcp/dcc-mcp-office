using System.Reflection;

namespace Office.Automation.Host;

/// <summary>Build and wire versions reported by the host.</summary>
internal static class HostBuildInfo
{
    internal const string ProtocolVersion = "office-rpc/1";

    internal static string Version { get; } = ReadAssemblyVersion();

    private static string ReadAssemblyVersion()
    {
        string? informationalVersion = typeof(HostBuildInfo).Assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?
            .InformationalVersion;
        if (string.IsNullOrWhiteSpace(informationalVersion))
        {
            throw new InvalidOperationException(
                "dcc-office-host is missing AssemblyInformationalVersion");
        }

        int buildMetadata = informationalVersion.IndexOf('+');
        return buildMetadata >= 0
            ? informationalVersion[..buildMetadata]
            : informationalVersion;
    }
}
