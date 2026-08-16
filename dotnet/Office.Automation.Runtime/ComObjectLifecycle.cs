namespace Office.Automation.Runtime;

/// <summary>
/// COM object lifetime rules (proposal §9.3). M0 placeholder; M1 implements
/// stable-ID re-resolution and request-bound release.
///
/// Rules this class will enforce:
///   - COM references never cross the process boundary.
///   - Never cache large numbers of leaf objects (Range, Shape, TextRange).
///   - Cache only stable document handles + application-level objects.
///   - Re-resolve leaf objects per request via stable IDs.
///   - Never FinalReleaseComObject a shared Application RCW.
///   - Release short-lived references at request boundaries.
///   - Sidecar process exit is the final isolation/cleanup mechanism.
///   - Copy event payloads to plain DTOs before queueing (no COM refs in async queues).
/// </summary>
public static class ComObjectLifecycle
{
    /// <summary>Releases a request-scoped COM reference, if release is safe.</summary>
    /// <remarks>M1: use Marshal.FinalReleaseComObject only for objects the
    /// request created; shared Application RCWs are exempt.</remarks>
    public static void ReleaseRequestScoped(object? rcw)
    {
        // Intentionally empty in M0.
    }
}
