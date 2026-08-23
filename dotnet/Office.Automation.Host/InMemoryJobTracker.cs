using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using Office.Automation.Com;

namespace Office.Automation.Host;

internal sealed class InMemoryJobTracker : IDisposable
{
    private readonly object _gate = new();
    private readonly Dictionary<string, JobRecord> _records = new(StringComparer.Ordinal);
    private readonly Queue<string> _terminalOrder = new();
    private readonly SemaphoreSlim _singleWorker = new(1, 1);
    private readonly HostNotificationQueue _notifications;
    private readonly int _maxRetained;
    private readonly int _maxPending;
    private bool _disposed;

    internal InMemoryJobTracker(
        HostNotificationQueue notifications,
        int maxRetained = 128,
        int maxPending = 32)
    {
        if (maxRetained < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(maxRetained));
        }
        if (maxPending < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(maxPending));
        }
        _notifications = notifications;
        _maxRetained = maxRetained;
        _maxPending = maxPending;
    }

    internal bool IsBusy
    {
        get
        {
            lock (_gate)
            {
                return _records.Values.Any(record => !record.Snapshot().Terminal);
            }
        }
    }

    internal JobSnapshot Submit(
        string capability,
        int total,
        Func<JobExecutionContext, object> work)
    {
        JobRecord record;
        lock (_gate)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            int pending = _records.Values.Count(record => !record.Snapshot().Terminal);
            if (pending >= _maxPending)
            {
                throw new OfficeComException(
                    OfficeErrorCode.OfficeAppBusy,
                    $"office job queue is full ({_maxPending} pending jobs)");
            }
            string jobId = $"job:{Guid.NewGuid():N}";
            record = new JobRecord(jobId, capability, Math.Max(0, total));
            _records.Add(jobId, record);
        }

        _notifications.PublishJobProgress(record.JobId, "queued", 0, record.Total);
        record.Task = Task.Run(() => Run(record, work));
        return record.Snapshot();
    }

    internal JobSnapshot Get(string jobId)
    {
        lock (_gate)
        {
            return Find(jobId).Snapshot();
        }
    }

    internal JobCancelResult Cancel(string jobId)
    {
        lock (_gate)
        {
            JobRecord record = Find(jobId);
            JobSnapshot snapshot = record.Snapshot();
            if (snapshot.Terminal)
            {
                return new JobCancelResult(jobId, Accepted: false, snapshot.Phase);
            }
            record.RequestCancellation();
            return new JobCancelResult(jobId, Accepted: true, record.Snapshot().Phase);
        }
    }

    private void Run(JobRecord record, Func<JobExecutionContext, object> work)
    {
        bool ownsWorker = false;
        try
        {
            _singleWorker.Wait(record.CancellationToken);
            ownsWorker = true;
            record.SetPhase("running");
            _notifications.PublishEvent(
                "office.application.busy",
                record.JobId,
                new { busy = true, job_id = record.JobId });
            _notifications.PublishJobProgress(
                record.JobId,
                "running",
                record.Completed,
                record.Total);

            var context = new JobExecutionContext(record, _notifications);
            object result = work(context);
            JsonNode resultNode = JsonSerializer.SerializeToNode(result)
                ?? throw new InvalidOperationException("job result serialized to null");
            string phase = context.CancellationObserved
                ? "cancelled"
                : DeriveTerminalPhase(resultNode);
            record.Complete(phase, resultNode, error: null);
        }
        catch (OperationCanceledException) when (record.CancellationRequested)
        {
            record.Complete("cancelled", result: null, error: null);
        }
        catch (OfficeComException ex)
        {
            record.Complete(
                "failed",
                result: null,
                new JobError(ex.Code.ToWireName(), ex.Message, ex.Indeterminate));
        }
        catch (Exception ex)
        {
            record.Complete(
                "failed",
                result: null,
                new JobError(
                    OfficeErrorCode.OfficeBackendUnavailable.ToWireName(),
                    ex.Message,
                    Indeterminate: false));
        }
        finally
        {
            if (ownsWorker)
            {
                _singleWorker.Release();
            }
            JobSnapshot terminal = record.Snapshot();
            _notifications.PublishEvent(
                "office.job.completed",
                record.JobId,
                new
                {
                    job_id = record.JobId,
                    capability = record.Capability,
                    phase = terminal.Phase,
                    completed = terminal.Completed,
                    total = terminal.Total,
                });
            RetainTerminal(record.JobId);
            _notifications.PublishEvent(
                "office.application.busy",
                record.JobId,
                new { busy = IsBusy, job_id = record.JobId });
        }
    }

    private static string DeriveTerminalPhase(JsonNode result)
    {
        int succeeded = result["changed"]?["succeeded"]?.GetValue<int>() ?? 0;
        int failed = result["changed"]?["failed"]?.GetValue<int>() ?? 0;
        if (failed == 0)
        {
            return "succeeded";
        }
        return succeeded > 0 ? "partially_succeeded" : "failed";
    }

    private JobRecord Find(string jobId) =>
        _records.GetValueOrDefault(jobId)
        ?? throw new OfficeArgumentException($"unknown job_id: {jobId}");

    private void RetainTerminal(string jobId)
    {
        lock (_gate)
        {
            _terminalOrder.Enqueue(jobId);
            while (_terminalOrder.Count > _maxRetained)
            {
                string expiredJobId = _terminalOrder.Dequeue();
                if (_records.Remove(expiredJobId, out JobRecord? expired))
                {
                    expired.Dispose();
                }
            }
        }
    }

    public void Dispose()
    {
        JobRecord[] records;
        lock (_gate)
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            records = _records.Values.ToArray();
            foreach (JobRecord record in records)
            {
                if (!record.Snapshot().Terminal)
                {
                    record.RequestCancellation();
                }
            }
        }

        Task[] tasks = records
            .Select(record => record.Task)
            .Where(task => task is not null)
            .Cast<Task>()
            .ToArray();
        bool completed = tasks.Length == 0
            || Task.WaitAll(tasks, TimeSpan.FromSeconds(30));
        if (completed)
        {
            foreach (JobRecord record in records)
            {
                record.Dispose();
            }
            _singleWorker.Dispose();
        }
    }

    internal sealed class JobRecord : IDisposable
    {
        private readonly object _gate = new();
        private readonly CancellationTokenSource _cancellation = new();
        private string _phase = "queued";
        private string _stage = "queued";
        private int _completed;
        private int _total;
        private bool _cancelRequested;
        private JsonNode? _result;
        private JobError? _error;
        private readonly string _createdAt = DateTimeOffset.UtcNow.ToString("O");
        private string _updatedAt = DateTimeOffset.UtcNow.ToString("O");

        internal JobRecord(string jobId, string capability, int total)
        {
            JobId = jobId;
            Capability = capability;
            _total = total;
        }

        internal string JobId { get; }

        internal string Capability { get; }

        internal Task? Task { get; set; }

        internal CancellationToken CancellationToken => _cancellation.Token;

        internal bool CancellationRequested => _cancellation.IsCancellationRequested;

        internal int Completed
        {
            get { lock (_gate) { return _completed; } }
        }

        internal int Total
        {
            get { lock (_gate) { return _total; } }
        }

        internal void RequestCancellation()
        {
            lock (_gate)
            {
                _cancelRequested = true;
                _updatedAt = DateTimeOffset.UtcNow.ToString("O");
            }
            _cancellation.Cancel();
        }

        internal void SetPhase(string phase)
        {
            lock (_gate)
            {
                _phase = phase;
                _stage = phase;
                _updatedAt = DateTimeOffset.UtcNow.ToString("O");
            }
        }

        internal void SetTotal(int total)
        {
            lock (_gate)
            {
                _total = Math.Max(0, total);
                _updatedAt = DateTimeOffset.UtcNow.ToString("O");
            }
        }

        internal void Report(string stage, int completed)
        {
            lock (_gate)
            {
                _stage = stage;
                _completed = Math.Clamp(completed, 0, _total);
                _updatedAt = DateTimeOffset.UtcNow.ToString("O");
            }
        }

        internal void Complete(string phase, JsonNode? result, JobError? error)
        {
            lock (_gate)
            {
                _phase = phase;
                _stage = phase;
                _result = result?.DeepClone();
                _error = error;
                _updatedAt = DateTimeOffset.UtcNow.ToString("O");
            }
        }

        internal JobSnapshot Snapshot()
        {
            lock (_gate)
            {
                return new JobSnapshot(
                    JobId,
                    Capability,
                    _phase,
                    _stage,
                    _completed,
                    _total,
                    _cancelRequested,
                    _createdAt,
                    _updatedAt,
                    _result?.DeepClone(),
                    _error);
            }
        }

        public void Dispose() => _cancellation.Dispose();
    }

    internal sealed class JobExecutionContext
    {
        private readonly JobRecord _record;
        private readonly HostNotificationQueue _notifications;

        internal JobExecutionContext(
            JobRecord record,
            HostNotificationQueue notifications)
        {
            _record = record;
            _notifications = notifications;
        }

        internal bool CancellationObserved { get; private set; }

        internal string JobId => _record.JobId;

        internal void SetTotal(int total) => _record.SetTotal(total);

        internal bool StopBeforeNextItem()
        {
            if (!_record.CancellationRequested)
            {
                return false;
            }
            CancellationObserved = true;
            return true;
        }

        internal void Report(string stage, int completed)
        {
            _record.Report(stage, completed);
            _notifications.PublishJobProgress(
                _record.JobId,
                stage,
                _record.Completed,
                _record.Total);
        }
    }
}

internal sealed record JobSnapshot(
    [property: JsonPropertyName("job_id")] string JobId,
    [property: JsonPropertyName("capability")] string Capability,
    [property: JsonPropertyName("phase")] string Phase,
    [property: JsonPropertyName("stage")] string Stage,
    [property: JsonPropertyName("completed")] int Completed,
    [property: JsonPropertyName("total")] int Total,
    [property: JsonPropertyName("cancel_requested")] bool CancelRequested,
    [property: JsonPropertyName("created_at")] string CreatedAt,
    [property: JsonPropertyName("updated_at")] string UpdatedAt,
    [property: JsonPropertyName("result")] JsonNode? Result,
    [property: JsonPropertyName("error")] JobError? Error)
{
    [JsonIgnore]
    internal bool Terminal =>
        Phase is "succeeded" or "partially_succeeded" or "failed" or "cancelled";
}

internal sealed record JobError(
    [property: JsonPropertyName("code")] string Code,
    [property: JsonPropertyName("message")] string Message,
    [property: JsonPropertyName("indeterminate")] bool Indeterminate);

internal sealed record JobCancelResult(
    [property: JsonPropertyName("job_id")] string JobId,
    [property: JsonPropertyName("accepted")] bool Accepted,
    [property: JsonPropertyName("phase")] string Phase);
