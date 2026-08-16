using System.Collections.Concurrent;
using System.Runtime.ExceptionServices;

namespace Office.Automation.Runtime;

/// <summary>
/// Owns the dedicated STA thread every Office COM call is marshalled through.
/// One instance per sidecar process (proposal §8.2 / §9.2).
///
/// M0 skeleton: queue + pump only. M1 adds the mandatory pieces from
/// proposal §9.2:
///   - single-writer queue (already here),
///   - Windows message pump (pump currently runs bare queue items),
///   - IOleMessageFilter with RPC_E_CALL_REJECTED backoff/retry,
///   - soft per-request timeout + sidecar hard timeout,
///   - cancellation tokens honoured only at safe boundaries.
/// </summary>
public sealed class StaDispatcher : IDisposable
{
    private readonly Thread _thread;
    private readonly BlockingCollection<Func<object?>> _queue = new();

    public StaDispatcher()
    {
        _thread = new Thread(Pump)
        {
            IsBackground = true,
            Name = "office-sta",
        };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
    }

    /// <summary>
    /// Submits work to the STA thread and blocks until it completes.
    /// All COM writes of one sidecar serialize through this queue.
    /// </summary>
    public T Post<T>(Func<T> work)
    {
        object? result = null;
        Exception? error = null;
        using var done = new ManualResetEventSlim(false);

        _queue.Add(() =>
        {
            try
            {
                result = work();
            }
            catch (Exception ex)
            {
                error = ex;
            }
            finally
            {
                done.Set();
            }

            return null;
        });

        done.Wait();

        if (error is not null)
        {
            ExceptionDispatchInfo.Capture(error).Throw();
        }

        return (T)result!;
    }

    private void Pump()
    {
        foreach (var work in _queue.GetConsumingEnumerable())
        {
            work();
        }
    }

    public void Dispose()
    {
        _queue.CompleteAdding();
    }
}
