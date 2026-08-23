using System.Text.Json;

namespace Office.Automation.Host;

public enum HostLogLevel
{
    Debug,
    Info,
    Warning,
    Error,
}

/// <summary>Bounded JSON-lines diagnostics with correlation but no payloads.</summary>
public sealed class HostLogger : IDisposable
{
    private readonly string _app;
    private readonly TextWriter _standardError;
    private readonly HostLogLevel _minimumLevel;
    private StreamWriter? _file;
    private readonly object _gate = new();

    public HostLogger(
        string app,
        TextWriter standardError,
        HostLogLevel minimumLevel,
        string? filePath)
    {
        _app = app;
        _standardError = standardError;
        _minimumLevel = minimumLevel;
        if (!string.IsNullOrWhiteSpace(filePath))
        {
            try
            {
                string fullPath = Path.GetFullPath(filePath);
                Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
                _file = new StreamWriter(
                    new FileStream(fullPath, FileMode.Append, FileAccess.Write, FileShare.Read))
                {
                    AutoFlush = true,
                };
            }
            catch (Exception error) when (IsFileSinkError(error))
            {
                Write(
                    HostLogLevel.Error,
                    "log.file_unavailable",
                    new Dictionary<string, object?> { ["message"] = error.Message });
            }
        }
    }

    public void RequestCompleted(
        string correlationId,
        string? capability,
        string? operationId,
        long durationMs,
        string outcomeCode) => Write(
            HostLogLevel.Info,
            "request.completed",
            new Dictionary<string, object?>
            {
                ["correlation_id"] = correlationId,
                ["capability"] = capability,
                ["operation_id"] = operationId,
                ["duration_ms"] = durationMs,
                ["outcome_code"] = outcomeCode,
            });

    public void Warning(string eventName, string message) => Write(
        HostLogLevel.Warning,
        eventName,
        new Dictionary<string, object?> { ["message"] = message });

    public void Info(string eventName, IReadOnlyDictionary<string, object?>? fields = null) =>
        Write(HostLogLevel.Info, eventName, fields);

    private void Write(
        HostLogLevel level,
        string eventName,
        IReadOnlyDictionary<string, object?>? fields)
    {
        if (level < _minimumLevel)
        {
            return;
        }
        var entry = new Dictionary<string, object?>
        {
            ["timestamp"] = DateTimeOffset.UtcNow.ToString("O"),
            ["level"] = level.ToString().ToLowerInvariant(),
            ["event"] = eventName,
            ["app"] = _app,
            ["pid"] = Environment.ProcessId,
        };
        if (fields is not null)
        {
            foreach ((string name, object? value) in fields)
            {
                entry[name] = value;
            }
        }
        string line = JsonSerializer.Serialize(entry);
        lock (_gate)
        {
            _standardError.WriteLine(line);
            _standardError.Flush();
            try
            {
                _file?.WriteLine(line);
            }
            catch (Exception error) when (IsFileSinkError(error))
            {
                CloseFileSink();
                _standardError.WriteLine(JsonSerializer.Serialize(new Dictionary<string, object?>
                {
                    ["timestamp"] = DateTimeOffset.UtcNow.ToString("O"),
                    ["level"] = "error",
                    ["event"] = "log.file_unavailable",
                    ["app"] = _app,
                    ["pid"] = Environment.ProcessId,
                    ["message"] = error.Message,
                }));
                _standardError.Flush();
            }
        }
    }

    private static bool IsFileSinkError(Exception error) => error is
        IOException or
        UnauthorizedAccessException or
        ArgumentException or
        NotSupportedException;

    private void CloseFileSink()
    {
        StreamWriter? file = _file;
        _file = null;
        try
        {
            file?.Dispose();
        }
        catch (Exception error) when (IsFileSinkError(error))
        {
            // The mandatory stderr sink remains live; the optional file sink is disabled.
        }
    }

    public void Dispose() => CloseFileSink();
}
