using System.Text.Json;
using System.Text.Json.Nodes;
using Office.Automation.Com;
using Office.Automation.Host;
using Xunit;

namespace Office.Automation.Host.Tests;

public sealed class JobRuntimeTests
{
    [Fact]
    public void PingIsSideEffectFreeAndReportsHostStatus()
    {
        using var router = new CommandRouter("powerpoint", enableDesktopCom: true);

        JsonNode response = Parse(router.Dispatch(
            """
            {"jsonrpc":"2.0","id":1,"method":"office.host.ping","params":{}}
            """));

        Assert.False(router.ComAttached);
        Assert.Equal("unknown", response["result"]!["com_attach_state"]!.GetValue<string>());
        Assert.Equal("ready", response["result"]!["state"]!.GetValue<string>());
        Assert.False(response["result"]!["busy"]!.GetValue<bool>());
    }

    [Fact]
    public void BatchSubmissionReturnsAJobAndCanBePolledWithoutOffice()
    {
        string workspace = TempDirectory();
        string source = Path.Combine(workspace, "deck.pptx");
        File.WriteAllText(source, "fixture");
        try
        {
            using var router = new CommandRouter(
                "powerpoint",
                enableDesktopCom: false,
                workspaceRoot: workspace);
            string request = JsonSerializer.Serialize(new
            {
                jsonrpc = "2.0",
                id = 2,
                method = "office.command.execute",
                @params = new
                {
                    capability = "batch.convert",
                    input = new
                    {
                        inputs = new[] { source },
                        output_directory = Path.Combine(workspace, "pdf"),
                    },
                    policy = new { workspace_root = workspace },
                },
            });

            JsonNode submitted = Parse(router.Dispatch(request));
            string jobId = submitted["result"]!["job_id"]!.GetValue<string>();

            Assert.StartsWith("job:", jobId, StringComparison.Ordinal);
            Assert.Contains(
                submitted["result"]!["phase"]!.GetValue<string>(),
                new[] { "queued", "running", "failed" });

            JsonNode status = WaitForTerminal(router, jobId);
            Assert.Equal("failed", status["result"]!["phase"]!.GetValue<string>());
            Assert.NotNull(status["result"]!["error"]);

            string cancelRequest = JsonSerializer.Serialize(new
            {
                jsonrpc = "2.0",
                id = 4,
                method = "office.job.cancel",
                @params = new { job_id = jobId },
            });
            JsonNode cancellation = Parse(router.Dispatch(cancelRequest));
            Assert.False(cancellation["result"]!["accepted"]!.GetValue<bool>());
            Assert.Equal("failed", cancellation["result"]!["phase"]!.GetValue<string>());
        }
        finally
        {
            Directory.Delete(workspace, recursive: true);
        }
    }

    [Fact]
    public void CancellationIsObservedBetweenItemsAndEmitsProgress()
    {
        var notifications = new HostNotificationQueue(
            "powerpoint",
            () => "office-host:test");
        using var jobs = new InMemoryJobTracker(notifications, maxRetained: 4);
        using var firstItemStarted = new ManualResetEventSlim();
        using var releaseFirstItem = new ManualResetEventSlim();
        JobSnapshot submitted = jobs.Submit("batch.convert", 3, context =>
        {
            context.SetTotal(3);
            int completed = 0;
            for (int index = 0; index < 3; index++)
            {
                if (context.StopBeforeNextItem())
                {
                    break;
                }
                if (index == 0)
                {
                    firstItemStarted.Set();
                    Assert.True(releaseFirstItem.Wait(TimeSpan.FromSeconds(5)));
                }
                completed++;
                context.Report("converting", completed);
            }
            return new
            {
                changed = new { files = 3, succeeded = completed, failed = 3 - completed },
            };
        });

        Assert.True(firstItemStarted.Wait(TimeSpan.FromSeconds(5)));
        JobCancelResult cancellation = jobs.Cancel(submitted.JobId);
        releaseFirstItem.Set();
        JobSnapshot terminal = WaitForTerminal(jobs, submitted.JobId);

        Assert.True(cancellation.Accepted);
        Assert.Equal("cancelled", terminal.Phase);
        Assert.Equal(1, terminal.Completed);
        Assert.Equal(3, terminal.Total);

        JsonNode[] messages = WaitForCompletionNotifications(notifications, submitted.JobId);
        Assert.Contains(messages, message =>
            message["method"]?.GetValue<string>() == "office.job.progress"
            && message["params"]?["completed"]?.GetValue<int>() == 1);
        Assert.Contains(messages, message =>
            message["method"]?.GetValue<string>() == "office.job.completed"
            && message["params"]?["correlation_id"]?.GetValue<string>() == submitted.JobId);
        Assert.Contains(messages, message =>
            message["method"]?.GetValue<string>() == "office.application.busy"
            && message["params"]?["context"]?["busy"]?.GetValue<bool>() == false);
    }

    [Fact]
    public void EventEnvelopeCarriesRequiredCorrelationFields()
    {
        var notifications = new HostNotificationQueue(
            "powerpoint",
            () => "office-host:test");

        notifications.PublishEvent(
            "office.document.saved",
            "correlation-7",
            new { path = "deck.pptx" });

        JsonNode message = Parse(Assert.Single(notifications.Drain()));
        JsonNode parameters = message["params"]!;
        Assert.Equal("dcc-mcp-office", parameters["provider"]!.GetValue<string>());
        Assert.Equal("office-host:test", parameters["application_instance"]!.GetValue<string>());
        Assert.Equal("correlation-7", parameters["correlation_id"]!.GetValue<string>());
        Assert.NotEmpty(parameters["timestamp"]!.GetValue<string>());
    }

    [Fact]
    public void JobQueueRejectsWorkBeyondItsBound()
    {
        var notifications = new HostNotificationQueue(
            "powerpoint",
            () => "office-host:test");
        using var jobs = new InMemoryJobTracker(
            notifications,
            maxRetained: 1,
            maxPending: 1);
        using var started = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        jobs.Submit("batch.convert", 1, context =>
        {
            started.Set();
            release.Wait(TimeSpan.FromSeconds(5));
            context.Report("converting", 1);
            return new { changed = new { succeeded = 1, failed = 0 } };
        });
        Assert.True(started.Wait(TimeSpan.FromSeconds(5)));

        OfficeComException error = Assert.Throws<OfficeComException>(() =>
            jobs.Submit(
                "batch.convert",
                1,
                _ => new { changed = new { succeeded = 1, failed = 0 } }));
        release.Set();

        Assert.Equal(OfficeErrorCode.OfficeAppBusy, error.Code);
        Assert.Contains("queue is full", error.Message);
    }

    private static JsonNode WaitForTerminal(CommandRouter router, string jobId)
    {
        DateTime deadline = DateTime.UtcNow.AddSeconds(5);
        while (DateTime.UtcNow < deadline)
        {
            string request = JsonSerializer.Serialize(new
            {
                jsonrpc = "2.0",
                id = 3,
                method = "office.job.get",
                @params = new { job_id = jobId },
            });
            JsonNode response = Parse(router.Dispatch(request));
            string phase = response["result"]!["phase"]!.GetValue<string>();
            if (phase is "succeeded" or "partially_succeeded" or "failed" or "cancelled")
            {
                return response;
            }
            Thread.Sleep(10);
        }
        throw new TimeoutException($"job {jobId} did not become terminal");
    }

    private static JobSnapshot WaitForTerminal(InMemoryJobTracker jobs, string jobId)
    {
        DateTime deadline = DateTime.UtcNow.AddSeconds(5);
        while (DateTime.UtcNow < deadline)
        {
            JobSnapshot snapshot = jobs.Get(jobId);
            if (snapshot.Phase is "succeeded" or "partially_succeeded" or "failed" or "cancelled")
            {
                return snapshot;
            }
            Thread.Sleep(10);
        }
        throw new TimeoutException($"job {jobId} did not become terminal");
    }

    private static JsonNode[] WaitForCompletionNotifications(
        HostNotificationQueue notifications,
        string jobId)
    {
        var messages = new List<JsonNode>();
        DateTime deadline = DateTime.UtcNow.AddSeconds(5);
        while (DateTime.UtcNow < deadline)
        {
            messages.AddRange(notifications.Drain().Select(Parse));
            if (messages.Any(message =>
                message["method"]?.GetValue<string>() == "office.job.completed"
                && message["params"]?["correlation_id"]?.GetValue<string>() == jobId))
            {
                return messages.ToArray();
            }
            Thread.Sleep(10);
        }
        throw new TimeoutException($"job {jobId} emitted no completion notification");
    }

    private static string TempDirectory()
    {
        string path = Path.Combine(Path.GetTempPath(), $"dcc-jobs-{Guid.NewGuid():N}");
        Directory.CreateDirectory(path);
        return path;
    }

    private static JsonNode Parse(string json) =>
        JsonNode.Parse(json) ?? throw new InvalidOperationException("response was null");
}
