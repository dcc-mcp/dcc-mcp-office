using System.Collections.Concurrent;
using System.Text.Json;

namespace Office.Automation.Host;

/// <summary>
/// Bounded producer queue for JSON-RPC notifications. The pipe owns delivery;
/// command and job code only publish typed envelopes here.
/// </summary>
internal sealed class HostNotificationQueue
{
    private const int MaxBufferedNotifications = 1_024;
    private readonly string _app;
    private readonly Func<string> _hostId;
    private readonly ConcurrentQueue<string> _messages = new();
    private int _count;

    internal HostNotificationQueue(string app, Func<string> hostId)
    {
        _app = app;
        _hostId = hostId;
    }

    internal void PublishJobProgress(
        string jobId,
        string stage,
        int completed,
        int total) =>
        Publish("office.job.progress", new
        {
            job_id = jobId,
            stage,
            completed,
            total,
        });

    internal void PublishEvent(
        string method,
        string correlationId,
        object? context = null,
        string? documentId = null,
        ulong? revision = null,
        object? selection = null) =>
        Publish(method, new
        {
            @event = method,
            provider = CapabilityCatalog.Current.Provider,
            application_instance = _hostId(),
            application = _app,
            document_id = documentId,
            revision,
            timestamp = DateTimeOffset.UtcNow.ToString("O"),
            correlation_id = correlationId,
            selection,
            context = context ?? new { },
        });

    internal IReadOnlyList<string> Drain()
    {
        var drained = new List<string>();
        while (_messages.TryDequeue(out string? message))
        {
            Interlocked.Decrement(ref _count);
            drained.Add(message);
        }
        return drained;
    }

    private void Publish(string method, object parameters)
    {
        string message = JsonSerializer.Serialize(new
        {
            jsonrpc = "2.0",
            method,
            @params = parameters,
        });
        _messages.Enqueue(message);
        int count = Interlocked.Increment(ref _count);
        while (count > MaxBufferedNotifications && _messages.TryDequeue(out _))
        {
            count = Interlocked.Decrement(ref _count);
        }
    }
}
