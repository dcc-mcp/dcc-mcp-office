using System.Text.Json;

namespace Office.Automation.Host;

/// <summary>Layered operational settings: defaults, file, environment, CLI.</summary>
public sealed record HostSettings
{
    private static readonly HashSet<string> FileProperties = new(StringComparer.Ordinal)
    {
        "attach_timeout_seconds",
        "request_timeout_seconds",
        "recovery_timeout_streak",
        "busy_retry_count",
        "pipe_buffer_bytes",
        "log_level",
        "log_path",
        "template_directories",
    };

    public TimeSpan AttachTimeout { get; init; } = TimeSpan.FromSeconds(60);

    public TimeSpan RequestTimeout { get; init; } = TimeSpan.FromSeconds(120);

    public int RecoveryTimeoutStreak { get; init; } = 2;

    public int BusyRetryCount { get; init; } = 30;

    public int PipeBufferBytes { get; init; } = 64 * 1024;

    public HostLogLevel LogLevel { get; init; } = HostLogLevel.Info;

    public string? LogPath { get; init; }

    public IReadOnlyList<string> TemplateDirectories { get; init; } = Array.Empty<string>();

    public static HostSettings Load(
        IReadOnlyList<string> args,
        string baseDirectory,
        IReadOnlyDictionary<string, string?>? environment = null)
    {
        string? explicitConfig = Option(args, "--config=");
        if (explicitConfig is not null && string.IsNullOrWhiteSpace(explicitConfig))
        {
            throw new InvalidDataException("--config path must not be empty");
        }
        string configPath = explicitConfig is null
            ? Path.Combine(baseDirectory, "dcc-office-host.json")
            : Path.GetFullPath(explicitConfig);
        if (explicitConfig is not null && !File.Exists(configPath))
        {
            throw new InvalidDataException($"Host config file not found: {configPath}");
        }

        var mutable = new MutableSettings();
        if (File.Exists(configPath))
        {
            ApplyFile(mutable, configPath);
        }
        ApplyEnvironment(mutable, environment ?? ProcessEnvironment());
        ApplyCli(mutable, args);
        return mutable.Build();
    }

    private static void ApplyFile(MutableSettings settings, string path)
    {
        using JsonDocument document = JsonDocument.Parse(File.ReadAllText(path));
        JsonElement root = document.RootElement;
        if (root.ValueKind != JsonValueKind.Object)
        {
            throw new InvalidDataException("Host settings must be a JSON object");
        }
        string[] unknown = root.EnumerateObject()
            .Select(property => property.Name)
            .Where(name => !FileProperties.Contains(name))
            .ToArray();
        if (unknown.Length > 0)
        {
            throw new InvalidDataException(
                $"unknown Host settings: {string.Join(", ", unknown)}");
        }
        string configDirectory = Path.GetDirectoryName(path)!;
        settings.AttachTimeoutSeconds = Number(root, "attach_timeout_seconds", settings.AttachTimeoutSeconds);
        settings.RequestTimeoutSeconds = Number(root, "request_timeout_seconds", settings.RequestTimeoutSeconds);
        settings.RecoveryTimeoutStreak = Integer(root, "recovery_timeout_streak", settings.RecoveryTimeoutStreak);
        settings.BusyRetryCount = Integer(root, "busy_retry_count", settings.BusyRetryCount);
        settings.PipeBufferBytes = Integer(root, "pipe_buffer_bytes", settings.PipeBufferBytes);
        settings.LogLevel = Text(root, "log_level", settings.LogLevel)
            ?? settings.LogLevel;
        string? logPath = Text(root, "log_path", settings.LogPath);
        if (logPath is not null && string.IsNullOrWhiteSpace(logPath))
        {
            throw new InvalidDataException("log_path must not be empty");
        }
        settings.LogPath = logPath is null
            ? null
            : Path.GetFullPath(logPath, configDirectory);
        if (root.TryGetProperty("template_directories", out JsonElement directories))
        {
            if (directories.ValueKind != JsonValueKind.Array)
            {
                throw new InvalidDataException("template_directories must be an array");
            }
            foreach (JsonElement item in directories.EnumerateArray())
            {
                string directory = item.GetString()
                    ?? throw new InvalidDataException(
                        "template_directories entries must be strings");
                if (string.IsNullOrWhiteSpace(directory))
                {
                    throw new InvalidDataException(
                        "template_directories entries must not be empty");
                }
                settings.TemplateDirectories.Add(
                    Path.GetFullPath(directory, configDirectory));
            }
        }
    }

    private static void ApplyEnvironment(
        MutableSettings settings,
        IReadOnlyDictionary<string, string?> environment)
    {
        settings.AttachTimeoutSeconds = EnvironmentNumber(
            environment,
            "DCC_OFFICE_ATTACH_TIMEOUT_SECONDS",
            settings.AttachTimeoutSeconds);
        settings.RequestTimeoutSeconds = EnvironmentNumber(
            environment,
            "DCC_OFFICE_REQUEST_TIMEOUT_SECONDS",
            settings.RequestTimeoutSeconds);
        settings.RecoveryTimeoutStreak = EnvironmentInteger(
            environment,
            "DCC_OFFICE_RECOVERY_TIMEOUT_STREAK",
            settings.RecoveryTimeoutStreak);
        settings.BusyRetryCount = EnvironmentInteger(
            environment,
            "DCC_OFFICE_BUSY_RETRY_COUNT",
            settings.BusyRetryCount);
        settings.PipeBufferBytes = EnvironmentInteger(
            environment,
            "DCC_OFFICE_PIPE_BUFFER_BYTES",
            settings.PipeBufferBytes);
        settings.LogLevel = EnvironmentText(
            environment,
            "DCC_OFFICE_LOG_LEVEL",
            settings.LogLevel) ?? settings.LogLevel;
        settings.LogPath = EnvironmentText(
            environment,
            "DCC_OFFICE_LOG_PATH",
            settings.LogPath);
        string? templateDirectories = EnvironmentText(
            environment,
            "DCC_OFFICE_TEMPLATE_DIRS",
            fallback: null);
        if (!string.IsNullOrWhiteSpace(templateDirectories))
        {
            settings.TemplateDirectories.AddRange(templateDirectories.Split(
                Path.PathSeparator,
                StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));
        }
    }

    private static void ApplyCli(MutableSettings settings, IReadOnlyList<string> args)
    {
        settings.AttachTimeoutSeconds = OptionNumber(
            args,
            "--attach-timeout-seconds=",
            settings.AttachTimeoutSeconds);
        settings.RequestTimeoutSeconds = OptionNumber(
            args,
            "--request-timeout-seconds=",
            settings.RequestTimeoutSeconds);
        settings.RecoveryTimeoutStreak = OptionInteger(
            args,
            "--recovery-timeout-streak=",
            settings.RecoveryTimeoutStreak);
        settings.BusyRetryCount = OptionInteger(
            args,
            "--busy-retry-count=",
            settings.BusyRetryCount);
        settings.PipeBufferBytes = OptionInteger(
            args,
            "--pipe-buffer-bytes=",
            settings.PipeBufferBytes);
        settings.LogLevel = Option(args, "--log-level=") ?? settings.LogLevel;
        settings.LogPath = Option(args, "--log-path=") ?? settings.LogPath;
        settings.TemplateDirectories.AddRange(args
            .Where(arg => arg.StartsWith("--template-dir=", StringComparison.Ordinal))
            .Select(arg => arg["--template-dir=".Length..]));
    }

    private static IReadOnlyDictionary<string, string?> ProcessEnvironment() =>
        Environment.GetEnvironmentVariables()
            .Cast<System.Collections.DictionaryEntry>()
            .ToDictionary(
                entry => (string)entry.Key,
                entry => entry.Value?.ToString(),
                StringComparer.OrdinalIgnoreCase);

    private static string? Option(IReadOnlyList<string> args, string prefix) =>
        args.LastOrDefault(arg => arg.StartsWith(prefix, StringComparison.Ordinal))?[prefix.Length..];

    private static double OptionNumber(
        IReadOnlyList<string> args,
        string prefix,
        double fallback) => ParseNumber(Option(args, prefix), prefix, fallback);

    private static int OptionInteger(
        IReadOnlyList<string> args,
        string prefix,
        int fallback) => ParseInteger(Option(args, prefix), prefix, fallback);

    private static double EnvironmentNumber(
        IReadOnlyDictionary<string, string?> values,
        string name,
        double fallback) => ParseNumber(Value(values, name), name, fallback);

    private static int EnvironmentInteger(
        IReadOnlyDictionary<string, string?> values,
        string name,
        int fallback) => ParseInteger(Value(values, name), name, fallback);

    private static string? EnvironmentText(
        IReadOnlyDictionary<string, string?> values,
        string name,
        string? fallback) => Value(values, name) ?? fallback;

    private static string? Value(IReadOnlyDictionary<string, string?> values, string name) =>
        values.TryGetValue(name, out string? value) && !string.IsNullOrWhiteSpace(value)
            ? value
            : null;

    private static double Number(JsonElement root, string name, double fallback)
    {
        if (!root.TryGetProperty(name, out JsonElement value))
        {
            return fallback;
        }
        return value.ValueKind == JsonValueKind.Number && value.TryGetDouble(out double result)
            ? result
            : throw new InvalidDataException($"{name} must be a number");
    }

    private static int Integer(JsonElement root, string name, int fallback)
    {
        if (!root.TryGetProperty(name, out JsonElement value))
        {
            return fallback;
        }
        return value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out int result)
            ? result
            : throw new InvalidDataException($"{name} must be an integer");
    }

    private static string? Text(JsonElement root, string name, string? fallback)
    {
        if (!root.TryGetProperty(name, out JsonElement value))
        {
            return fallback;
        }
        return value.ValueKind == JsonValueKind.String
            ? value.GetString()
            : throw new InvalidDataException($"{name} must be a string");
    }

    private static double ParseNumber(string? value, string name, double fallback) =>
        value is null
            ? fallback
            : double.TryParse(
                value,
                System.Globalization.NumberStyles.AllowDecimalPoint,
                System.Globalization.CultureInfo.InvariantCulture,
                out double parsed)
                ? parsed
                : throw new InvalidDataException($"{name} must be a number");

    private static int ParseInteger(string? value, string name, int fallback) =>
        value is null
            ? fallback
            : int.TryParse(value, out int parsed)
                ? parsed
                : throw new InvalidDataException($"{name} must be an integer");

    private sealed class MutableSettings
    {
        public double AttachTimeoutSeconds { get; set; } = 60;
        public double RequestTimeoutSeconds { get; set; } = 120;
        public int RecoveryTimeoutStreak { get; set; } = 2;
        public int BusyRetryCount { get; set; } = 30;
        public int PipeBufferBytes { get; set; } = 64 * 1024;
        public string LogLevel { get; set; } = "info";
        public string? LogPath { get; set; }
        public List<string> TemplateDirectories { get; } = new();

        public HostSettings Build()
        {
            if (!double.IsFinite(AttachTimeoutSeconds)
                || !double.IsFinite(RequestTimeoutSeconds)
                || AttachTimeoutSeconds <= 0
                || RequestTimeoutSeconds <= 0
                || AttachTimeoutSeconds > TimeSpan.MaxValue.TotalSeconds
                || RequestTimeoutSeconds > TimeSpan.MaxValue.TotalSeconds)
            {
                throw new InvalidDataException("timeouts must be greater than zero");
            }
            if (RecoveryTimeoutStreak < 1 || BusyRetryCount < 1)
            {
                throw new InvalidDataException("retry and recovery counts must be positive");
            }
            if (PipeBufferBytes is < 4096 or > 4 * 1024 * 1024)
            {
                throw new InvalidDataException(
                    "pipe_buffer_bytes must be between 4096 and 4194304");
            }
            if (!Enum.TryParse(LogLevel, ignoreCase: true, out HostLogLevel logLevel)
                || !Enum.IsDefined(logLevel))
            {
                throw new InvalidDataException(
                    "log_level must be debug, info, warning, or error");
            }
            if (LogPath is not null && string.IsNullOrWhiteSpace(LogPath))
            {
                throw new InvalidDataException("log_path must not be empty");
            }
            if (TemplateDirectories.Any(string.IsNullOrWhiteSpace))
            {
                throw new InvalidDataException("template directories must not be empty");
            }
            return new HostSettings
            {
                AttachTimeout = TimeSpan.FromSeconds(AttachTimeoutSeconds),
                RequestTimeout = TimeSpan.FromSeconds(RequestTimeoutSeconds),
                RecoveryTimeoutStreak = RecoveryTimeoutStreak,
                BusyRetryCount = BusyRetryCount,
                PipeBufferBytes = PipeBufferBytes,
                LogLevel = logLevel,
                LogPath = LogPath,
                TemplateDirectories = TemplateDirectories.ToArray(),
            };
        }
    }
}
