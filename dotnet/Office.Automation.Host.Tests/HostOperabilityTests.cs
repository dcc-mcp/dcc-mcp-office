using System.IO.Pipes;
using System.Text.Json;
using Office.Automation.Host;
using Xunit;

namespace Office.Automation.Host.Tests;

public sealed class HostOperabilityTests
{
    [Fact]
    public void RuntimeSettingsLayerDefaultsFileEnvironmentAndCli()
    {
        string temporary = Path.Combine(
            Path.GetTempPath(),
            $"dcc-office-settings-{Guid.NewGuid():N}");
        Directory.CreateDirectory(temporary);
        try
        {
            File.WriteAllText(
                Path.Combine(temporary, "dcc-office-host.json"),
                """
                {
                  "attach_timeout_seconds": 70,
                  "request_timeout_seconds": 150,
                  "recovery_timeout_streak": 4,
                  "busy_retry_count": 20,
                  "pipe_buffer_bytes": 32768,
                  "log_level": "warning",
                  "template_directories": ["from-file"]
                }
                """);
            var environment = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase)
            {
                ["DCC_OFFICE_REQUEST_TIMEOUT_SECONDS"] = "180",
                ["DCC_OFFICE_LOG_LEVEL"] = "debug",
                ["DCC_OFFICE_TEMPLATE_DIRS"] = "from-env-a;from-env-b",
            };

            HostSettings settings = HostSettings.Load(
                [
                    "--attach-timeout-seconds=90",
                    "--busy-retry-count=12",
                    "--log-path=host.jsonl",
                    "--template-dir=from-cli",
                ],
                temporary,
                environment);

            Assert.Equal(TimeSpan.FromSeconds(90), settings.AttachTimeout);
            Assert.Equal(TimeSpan.FromSeconds(180), settings.RequestTimeout);
            Assert.Equal(4, settings.RecoveryTimeoutStreak);
            Assert.Equal(12, settings.BusyRetryCount);
            Assert.Equal(32768, settings.PipeBufferBytes);
            Assert.Equal(HostLogLevel.Debug, settings.LogLevel);
            Assert.Equal("host.jsonl", settings.LogPath);
            Assert.Equal(
                [
                    Path.Combine(temporary, "from-file"),
                    "from-env-a",
                    "from-env-b",
                    "from-cli",
                ],
                settings.TemplateDirectories);
        }
        finally
        {
            Directory.Delete(temporary, recursive: true);
        }
    }

    [Fact]
    public void JsonLoggerEmitsCorrelatedRequestOutcomeWithoutInputPayloads()
    {
        var output = new StringWriter();
        using var logger = new HostLogger(
            "powerpoint",
            output,
            HostLogLevel.Debug,
            filePath: null);

        logger.RequestCompleted(
            correlationId: "rpc:7",
            capability: "deck.compile",
            operationId: "op-123",
            durationMs: 42,
            outcomeCode: "ok");

        using JsonDocument document = JsonDocument.Parse(output.ToString());
        JsonElement entry = document.RootElement;
        Assert.Equal("request.completed", entry.GetProperty("event").GetString());
        Assert.Equal("powerpoint", entry.GetProperty("app").GetString());
        Assert.Equal("rpc:7", entry.GetProperty("correlation_id").GetString());
        Assert.Equal("deck.compile", entry.GetProperty("capability").GetString());
        Assert.Equal("op-123", entry.GetProperty("operation_id").GetString());
        Assert.Equal(42, entry.GetProperty("duration_ms").GetInt64());
        Assert.Equal("ok", entry.GetProperty("outcome_code").GetString());
        Assert.False(entry.TryGetProperty("input", out _));
    }

    [Fact]
    public void UnavailableOptionalLogFileFallsBackToStructuredStderr()
    {
        string directory = Path.Combine(
            Path.GetTempPath(),
            $"dcc-office-log-sink-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        try
        {
            var output = new StringWriter();

            using var logger = new HostLogger(
                "powerpoint",
                output,
                HostLogLevel.Debug,
                directory);
            logger.Info("host.ready");

            JsonElement[] entries = output.ToString().Split(
                    Environment.NewLine,
                    StringSplitOptions.RemoveEmptyEntries)
                .Select(line => JsonDocument.Parse(line).RootElement.Clone())
                .ToArray();
            Assert.Contains(entries, entry =>
                entry.GetProperty("event").GetString() == "log.file_unavailable");
            Assert.Contains(entries, entry =>
                entry.GetProperty("event").GetString() == "host.ready");
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void NonRecursiveGlobDoesNotWalkSubdirectoriesAndReportsUnmatchedSpecs()
    {
        string temporary = Path.Combine(
            Path.GetTempPath(),
            $"dcc-office-glob-{Guid.NewGuid():N}");
        string nested = Path.Combine(temporary, "nested");
        Directory.CreateDirectory(nested);
        try
        {
            string direct = Path.Combine(temporary, "direct.pptx");
            string descendant = Path.Combine(nested, "descendant.pptx");
            File.WriteAllText(direct, "fixture");
            File.WriteAllText(descendant, "fixture");

            InputResolution shallow = InputResolver.Resolve(
                [Path.Combine(temporary, "*.pptx"), Path.Combine(temporary, "missing.pptx")]);
            InputResolution recursive = InputResolver.Resolve(
                [Path.Combine(temporary, "**", "*.pptx")]);

            Assert.Equal([direct], shallow.Paths);
            Assert.Single(shallow.Warnings);
            Assert.Contains("missing.pptx", shallow.Warnings[0], StringComparison.Ordinal);
            Assert.Equal(
                [direct, descendant],
                recursive.Paths.Order(StringComparer.OrdinalIgnoreCase));
        }
        finally
        {
            Directory.Delete(temporary, recursive: true);
        }
    }

    [Fact]
    public async Task CancellingAnIdlePipeServerStopsItsAcceptLoopPromptly()
    {
        string pipe = $"dcc-office-cancel-{Guid.NewGuid():N}";
        using var cancellation = new CancellationTokenSource();
        var server = new OfficePipeServer(
            "powerpoint",
            request => request,
            $@"\\.\pipe\{pipe}");
        Task running = Task.Run(() => server.Run(cancellation.Token));
        await Task.Delay(50);

        cancellation.Cancel();
        bool stopped;
        try
        {
            await running.WaitAsync(TimeSpan.FromSeconds(2));
            stopped = true;
        }
        catch (TimeoutException)
        {
            stopped = false;
        }
        if (!stopped)
        {
            using var client = new NamedPipeClientStream(".", pipe, PipeDirection.InOut);
            client.Connect(1000);
            await running.WaitAsync(TimeSpan.FromSeconds(2));
        }

        Assert.True(stopped, "idle WaitForConnection did not observe cancellation");
    }

    [Fact]
    public async Task CancellingAConnectedIdlePipeServerStopsItsReadLoopPromptly()
    {
        string pipe = $"dcc-office-read-cancel-{Guid.NewGuid():N}";
        using var cancellation = new CancellationTokenSource();
        var server = new OfficePipeServer(
            "powerpoint",
            request => request,
            $@"\\.\pipe\{pipe}");
        Task running = Task.Run(() => server.Run(cancellation.Token));
        using var client = new NamedPipeClientStream(".", pipe, PipeDirection.InOut);
        await client.ConnectAsync(1000);

        cancellation.Cancel();

        await running.WaitAsync(TimeSpan.FromSeconds(2));
    }

    [Fact]
    public void MissingParentProcessCancelsTheHostLifetime()
    {
        using CancellationTokenSource lifetime = ParentProcessMonitor.Watch(
            int.MaxValue,
            TimeSpan.FromMilliseconds(10));

        Assert.True(lifetime.Token.WaitHandle.WaitOne(TimeSpan.FromSeconds(1)));
    }

    [Fact]
    public void HandshakeAdvertisesInstalledDesktopComWithoutStartingPowerPoint()
    {
        bool attachAttempted = false;
        using var router = new CommandRouter(
            "powerpoint",
            enableDesktopCom: true,
            desktopComAvailable: true,
            attachDesktop: _ =>
            {
                attachAttempted = true;
                return true;
            });

        using JsonDocument response = JsonDocument.Parse(router.Dispatch(
            """
            {"jsonrpc":"2.0","id":1,"method":"office.host.handshake","params":{"requested_app":"powerpoint"}}
            """));

        Assert.False(attachAttempted);
        Assert.False(router.ComAttachAttempted);
        Assert.DoesNotContain(router.DrainNotifications(), notification =>
            notification.Contains("office.application.started", StringComparison.Ordinal));
        Assert.Contains(
            "desktop_com",
            response.RootElement
                .GetProperty("result")
                .GetProperty("capability_manifest")
                .GetProperty("execution_modes")
                .EnumerateArray()
                .Select(item => item.GetString()));
    }
}
