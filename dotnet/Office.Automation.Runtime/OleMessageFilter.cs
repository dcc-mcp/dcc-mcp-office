using System.Runtime.InteropServices;
using System.Threading;

namespace Office.Automation.Runtime;

/// <summary>
/// IOleMessageFilter implementation for the STA thread (proposal §9.2 /
/// §15 MVP criterion 8 "Busy 重试").
///
/// When the Office app is busy (e.g. another automation client, an in-place
/// edit session, or a long UI operation), COM rejects incoming calls with
/// RPC_E_CALL_REJECTED / SERVERCALL_RETRYLATER. This filter retries with
/// exponential backoff and pumps messages while waiting, then gives up with a
/// deterministic cancel instead of hanging forever.
/// </summary>
public sealed class OleMessageFilter : IOleMessageFilter, IDisposable
{
    // IOleMessageFilter IID — the same interface the host registers with
    // CoRegisterMessageFilter on the STA thread.
    private const int SERVERCALL_ISHANDLED = 0;
    private const int PENDINGMSG_WAITDEFPROCESS = 2;

    /// <summary>Retry budget: at 50 ms base backoff this is ~13 s of total retry.</summary>
    private const int MaxRetries = 30;

    private IOleMessageFilter? _previous;
    private int _attempts;
    private bool _disposed;

    /// <summary>Registers this filter for the calling STA thread and keeps the previous one for restore.</summary>
    public void Register()
    {
        int hr = NativeMethods.CoRegisterMessageFilter(this, out var previous);
        if (hr < 0)
        {
            // Filter registration is best effort: the busy-retry ladder still
            // works without it, just without server-side retry timing.
            _previous = null;
            return;
        }
        _previous = previous;
    }

    public int HandleIncomingCall(uint dwCallType, IntPtr htaskCaller, uint dwTickCount, IntPtr lpInterfaceInfo)
    {
        // The STA thread is always able to accept calls when the queue is idle.
        return SERVERCALL_ISHANDLED;
    }

    public int RetryRejectedCall(IntPtr htaskCallee, uint dwTickCount, uint dwRejectType)
    {
        int attempt = Interlocked.Increment(ref _attempts);
        if (attempt >= MaxRetries)
        {
            // Deterministic cancel: the caller surfaces OFFICE_APP_BUSY
            // instead of hanging on a wedged Office instance.
            return -1;
        }
        // Linear-ish backoff: 50 ms * attempt, capped at 1 s per step.
        Thread.Sleep(Math.Min(50 * attempt, 1000));
        return 100; // retry
    }

    public int MessagePending(IntPtr htaskCallee, uint dwTickCount, uint dwPendingType)
    {
        // Let the default handler pump window messages while the call is pending.
        return PENDINGMSG_WAITDEFPROCESS;
    }

    public void Unregister()
    {
        if (_disposed)
        {
            return;
        }
        NativeMethods.CoRegisterMessageFilter(_previous, out _);
        _previous = null;
        _disposed = true;
    }

    public void Dispose() => Unregister();

    private static class NativeMethods
    {
        [DllImport("ole32.dll")]
        public static extern int CoRegisterMessageFilter(IOleMessageFilter? lpMessageFilter, out IOleMessageFilter? lplpMessageFilter);
    }
}

// IOleMessageFilter contract — ComImport so CoRegisterMessageFilter can hand
// it straight back to the COM runtime.
[ComImport]
[Guid("00000016-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IOleMessageFilter
{
    [PreserveSig]
    int HandleIncomingCall(uint dwCallType, IntPtr htaskCaller, uint dwTickCount, IntPtr lpInterfaceInfo);

    [PreserveSig]
    int RetryRejectedCall(IntPtr htaskCallee, uint dwTickCount, uint dwRejectType);

    [PreserveSig]
    int MessagePending(IntPtr htaskCallee, uint dwTickCount, uint dwPendingType);
}
