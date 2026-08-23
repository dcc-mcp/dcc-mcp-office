using System.Collections.Concurrent;
using System.Runtime.InteropServices;

namespace Office.Automation.Runtime;

/// <summary>
/// Owns the dedicated STA thread every Office COM call is marshalled through.
/// One instance per sidecar process (proposal §8.2 / §9.2).
///
/// M1 implementation of the mandatory pieces from proposal §9.2:
///   - single-writer queue (all COM writes of one sidecar serialize here),
///   - COM is initialised on the STA thread and a Windows message pump runs
///     whenever the queue is idle (cross-apartment calls and Office events
///     keep flowing — proposal §9.2 "Message Pump"),
///   - <see cref="OleMessageFilter"/> handles Busy / retry-later with
///     backoff (RPC_E_CALL_REJECTED ladder),
///   - soft per-request timeout: the caller gets a timeout error while the
///     STA work still runs to completion in the background — COM calls are
///     never aborted mid-flight (proposal §9.2 "取消令牌只在安全边界生效"),
///   - cancellation tokens are honoured only between queue items.
/// </summary>
public sealed class StaDispatcher : IDisposable
{
    private static readonly TimeSpan IdlePumpQuantum = TimeSpan.FromMilliseconds(16);

    private readonly Thread _thread;
    private readonly BlockingCollection<Action> _queue = new();
    private readonly OleMessageFilter _messageFilter;
    private volatile bool _disposed;
    private int _threadId = -1;
    private int _timedOutWorkCount;

    public StaDispatcher(int busyRetryCount = 30)
    {
        _messageFilter = new OleMessageFilter(busyRetryCount);
        _thread = new Thread(Pump)
        {
            IsBackground = true,
            Name = "office-sta",
        };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
    }

    /// <summary>Managed id of the STA thread (available once the pump starts).</summary>
    public int ThreadId
    {
        get
        {
            if (Volatile.Read(ref _threadId) == -1)
            {
                SpinWait.SpinUntil(() => Volatile.Read(ref _threadId) != -1, TimeSpan.FromSeconds(5));
            }
            return Volatile.Read(ref _threadId);
        }
    }

    /// <summary>
    /// Submits work to the STA thread and blocks until it completes.
    /// All COM writes of one sidecar serialize through this queue.
    /// </summary>
    public T Post<T>(Func<T> work) => Post(work, TimeSpan.FromSeconds(120));

    /// <summary>Submits void work and blocks until it completes.</summary>
    public void Post(Action work) => Post(work, TimeSpan.FromSeconds(120));

    /// <summary>
    /// Submits work with a soft timeout. On timeout a
    /// <see cref="StaSoftTimeoutException"/> is thrown, but the STA thread
    /// keeps running the work to completion — COM calls are never interrupted
    /// in an arbitrary intermediate state (proposal §9.2).
    /// </summary>
    public T Post<T>(Func<T> work, TimeSpan softTimeout)
    {
        ThrowIfDisposed();
        ThrowIfTimedOutWorkIsStillRunning();

        var completion = new TaskCompletionSource<T>(
            TaskCreationOptions.RunContinuationsAsynchronously);
        var timeoutState = 0;

        _queue.Add(() =>
        {
            try
            {
                completion.TrySetResult(work());
            }
            catch (Exception ex)
            {
                completion.TrySetException(ex);
            }
            finally
            {
                if (Interlocked.CompareExchange(ref timeoutState, 2, 1) == 1)
                {
                    Interlocked.Decrement(ref _timedOutWorkCount);
                }
            }
        });

        if (!completion.Task.IsCompleted
            && Task.WaitAny(new Task[] { completion.Task }, softTimeout) < 0)
        {
            MarkTimedOutWork(completion.Task, ref timeoutState);
            throw new StaSoftTimeoutException(
                $"STA request exceeded soft timeout {softTimeout.TotalSeconds:F0}s; " +
                "the operation continues in the background and the sidecar stays consistent.");
        }

        return completion.Task.GetAwaiter().GetResult();
    }

    public void Post(Action work, TimeSpan softTimeout)
    {
        Post(() =>
        {
            work();
            return true;
        }, softTimeout);
    }

    private void MarkTimedOutWork(Task completion, ref int timeoutState)
    {
        if (Interlocked.CompareExchange(ref timeoutState, 1, 0) != 0)
        {
            return;
        }

        Interlocked.Increment(ref _timedOutWorkCount);
        if (completion.IsCompleted
            && Interlocked.CompareExchange(ref timeoutState, 2, 1) == 1)
        {
            Interlocked.Decrement(ref _timedOutWorkCount);
        }
    }

    private void ThrowIfTimedOutWorkIsStillRunning()
    {
        if (Volatile.Read(ref _timedOutWorkCount) > 0)
        {
            throw new StaDispatcherBusyException(
                "A previously timed-out STA request is still running; retry after it reaches a safe boundary.");
        }
    }

    private void Pump()
    {
        Volatile.Write(ref _threadId, Environment.CurrentManagedThreadId);

        // STA thread: explicit COM initialisation, then the message filter that
        // implements the Busy retry ladder (§9.2: IOleMessageFilter +
        // RPC_E_CALL_REJECTED backoff).
        NativeMethods.CoInitializeEx(IntPtr.Zero, NativeMethods.COINIT_APARTMENTTHREADED);
        _messageFilter.Register();
        try
        {
            while (!_disposed)
            {
                if (_queue.TryTake(out var work, IdlePumpQuantum))
                {
                    work();
                    continue;
                }

                // Idle: pump window messages so cross-apartment COM calls and
                // Office events keep flowing while nothing is queued.
                NativeMethods.PumpPendingMessages();
            }
        }
        finally
        {
            _messageFilter.Unregister();
            NativeMethods.CoUninitialize();
        }
    }

    private void ThrowIfDisposed()
    {
        if (_disposed)
        {
            throw new ObjectDisposedException(nameof(StaDispatcher));
        }
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }
        _disposed = true;
        _queue.CompleteAdding();
        // A hung COM call can keep the STA thread alive beyond the join
        // timeout; the sidecar process exit remains the final isolation and
        // cleanup mechanism (proposal §9.3).
        _thread.Join(TimeSpan.FromSeconds(10));
    }

    private static class NativeMethods
    {
        public const uint COINIT_APARTMENTTHREADED = 0x2;
        public const uint PM_REMOVE = 0x0001;

        [DllImport("ole32.dll")]
        public static extern int CoInitializeEx(IntPtr pvReserved, uint dwCoInit);

        [DllImport("ole32.dll")]
        public static extern void CoUninitialize();

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool PeekMessage(out Msg lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax, uint wRemoveMsg);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool TranslateMessage([In] ref Msg lpMsg);

        [DllImport("user32.dll")]
        private static extern IntPtr DispatchMessage([In] ref Msg lpMsg);

        [StructLayout(LayoutKind.Sequential)]
        private struct Msg
        {
            public IntPtr hwnd;
            public uint message;
            public IntPtr wParam;
            public IntPtr lParam;
            public uint time;
            public System.Drawing.Point pt;
        }

        public static void PumpPendingMessages()
        {
            while (PeekMessage(out var msg, IntPtr.Zero, 0, 0, PM_REMOVE))
            {
                if (msg.message == 0x0012 /* WM_QUIT */)
                {
                    return;
                }
                TranslateMessage(ref msg);
                DispatchMessage(ref msg);
            }
        }
    }
}

/// <summary>Raised when an STA request exceeds its soft timeout (proposal §9.2).</summary>
public sealed class StaSoftTimeoutException : TimeoutException
{
    public StaSoftTimeoutException(string message)
        : base(message)
    {
    }
}

/// <summary>Raised while a previously timed-out STA request is still running.</summary>
public sealed class StaDispatcherBusyException : InvalidOperationException
{
    public StaDispatcherBusyException(string message)
        : base(message)
    {
    }
}
