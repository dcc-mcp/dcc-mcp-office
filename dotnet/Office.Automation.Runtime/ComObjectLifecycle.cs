using System.Runtime.InteropServices;

namespace Office.Automation.Runtime;

/// <summary>
/// COM object lifetime rules (proposal §9.3), enforced for every COM call:
///   - COM references never cross the process boundary (host returns DTOs only).
///   - Leaf objects (Range, Shape, TextRange, ...) are request-scoped and
///     released at request boundaries via <see cref="ComLease"/>.
///   - Shared application/document RCWs are never released here — the owning
///     backend manages them explicitly.
///   - Event payloads are copied to plain DTOs before queueing (no COM refs
///     in async queues).
/// </summary>
public static class ComObjectLifecycle
{
    /// <summary>
    /// Releases a request-scoped COM reference, if release is safe.
    /// Shared Application RCWs are exempt by convention: only objects that
    /// look like runtime callable wrappers are released, and
    /// Marshal.FinalReleaseComObject is never used on a shared reference.
    /// </summary>
    public static void ReleaseRequestScoped(object? rcw)
    {
        if (rcw is null || !Marshal.IsComObject(rcw))
        {
            return;
        }
        try
        {
            Marshal.ReleaseComObject(rcw);
        }
        catch (Exception)
        {
            // Releasing twice is an error, not a policy failure; the process
            // boundary remains the final cleanup mechanism.
        }
    }
}

/// <summary>
/// Request-scoped COM lease: releases its target exactly once when disposed.
/// Wrap every leaf COM object (Shape, TextRange, Range, Worksheet, ...) in a
/// lease so request boundaries clean up deterministically (proposal §9.3).
/// </summary>
public sealed class ComLease : IDisposable
{
    private object? _target;
    private bool _release;

    public ComLease(object? target, bool releaseOnDispose = true)
    {
        _target = target;
        _release = releaseOnDispose;
    }

    public object? Target => _target;

    /// <summary>Detaches without releasing (ownership transfer).</summary>
    public object? Detach()
    {
        _release = false;
        return _target;
    }

    public void Dispose()
    {
        if (_release)
        {
            ComObjectLifecycle.ReleaseRequestScoped(_target);
        }
        _target = null;
        _release = false;
    }
}
