using System.Diagnostics;
using System.ComponentModel;

namespace Office.Automation.Host;

public static class ParentProcessMonitor
{
    public static CancellationTokenSource Watch(int processId, TimeSpan interval)
    {
        if (processId <= 0 || processId == Environment.ProcessId)
        {
            throw new ArgumentOutOfRangeException(
                nameof(processId),
                "parent process id must identify another positive process");
        }
        if (interval <= TimeSpan.Zero)
        {
            throw new ArgumentOutOfRangeException(nameof(interval));
        }

        var lifetime = new CancellationTokenSource();
        _ = Task.Run(async () =>
        {
            while (!lifetime.IsCancellationRequested)
            {
                if (HasExited(processId))
                {
                    lifetime.Cancel();
                    return;
                }
                try
                {
                    await Task.Delay(interval, lifetime.Token).ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    return;
                }
            }
        });
        return lifetime;
    }

    private static bool HasExited(int processId)
    {
        try
        {
            using Process process = Process.GetProcessById(processId);
            return process.HasExited;
        }
        catch (ArgumentException)
        {
            return true;
        }
        catch (Exception error) when (error is InvalidOperationException or Win32Exception)
        {
            return true;
        }
    }
}
