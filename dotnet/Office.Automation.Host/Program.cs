using Office.Automation.Runtime;

namespace Office.Automation.Host;

/// <summary>
/// dcc-office-host — per-application Office sidecar entry point.
///
/// M0 skeleton: argument parsing + STA dispatcher smoke test.
/// M1 wires (proposal §8/§9/§12):
///   - named pipe server \\.\pipe\dcc-mcp-office-{app}-{user_sid}-{session_id},
///   - office-rpc/1 handshake + capability manifest + heartbeat,
///   - attach-or-create the Office application (OfficeInstanceResolver),
///   - command execution through the STA dispatcher,
///   - graceful shutdown + orphan-process policy (never force-kill user Office).
///
/// One process per application: --app=powerpoint|word|excel|outlook-classic|
/// visio|project|access. Publish aliases (dcc-office-powerpoint-host.exe, ...)
/// can point at the same binary (proposal §22.1).
/// </summary>
public static class Program
{
    private static readonly string[] SupportedApps =
    {
        "powerpoint", "word", "excel", "outlook-classic", "visio", "project", "access",
    };

    public static int Main(string[] args)
    {
        string? app = null;
        foreach (var arg in args)
        {
            if (arg.StartsWith("--app=", StringComparison.Ordinal))
            {
                app = arg["--app=".Length..];
            }
        }

        if (app is null || !SupportedApps.Contains(app))
        {
            Console.Error.WriteLine(
                "usage: dcc-office-host --app=<powerpoint|word|excel|outlook-classic|visio|project|access>");
            return 2;
        }

        Console.WriteLine($"dcc-office-host: app={app} (office-rpc/1, M0 skeleton — pipe server lands in M1)");

        // M0: prove the STA queue machinery starts and stops cleanly.
        using (var sta = new StaDispatcher())
        {
            var threadId = sta.Post(() => Environment.CurrentManagedThreadId);
            Console.WriteLine($"sta dispatch ok: submitted on {Environment.CurrentManagedThreadId}, ran on {threadId}");
        }

        return 0;
    }
}
