using System.Diagnostics;
using Xunit;

namespace Office.Automation.Runtime.Tests;

public sealed class StaDispatcherTests
{
    [Fact]
    public void TimedOutWorkCanFinishWithoutKillingDispatcher()
    {
        using var dispatcher = new StaDispatcher();
        using var releaseWork = new ManualResetEventSlim(false);

        Assert.Throws<StaSoftTimeoutException>(() =>
            dispatcher.Post(
                () => releaseWork.Wait(),
                TimeSpan.FromMilliseconds(50)));

        releaseWork.Set();

        var deadline = Stopwatch.StartNew();
        while (true)
        {
            try
            {
                Assert.Equal(42, dispatcher.Post(() => 42, TimeSpan.FromSeconds(2)));
                break;
            }
            catch (StaDispatcherBusyException) when (deadline.Elapsed < TimeSpan.FromSeconds(2))
            {
                Thread.Yield();
            }
        }
    }

    [Fact]
    public void TimedOutWorkRejectsNewRequestsUntilItFinishes()
    {
        using var dispatcher = new StaDispatcher();
        using var releaseWork = new ManualResetEventSlim(false);

        Assert.Throws<StaSoftTimeoutException>(() =>
            dispatcher.Post(
                () => releaseWork.Wait(),
                TimeSpan.FromMilliseconds(50)));

        try
        {
            var elapsed = Stopwatch.StartNew();
            var error = Assert.Throws<StaDispatcherBusyException>(() =>
                dispatcher.Post(() => 42, TimeSpan.FromSeconds(2)));
            elapsed.Stop();

            Assert.Contains("still running", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.True(
                elapsed.Elapsed < TimeSpan.FromMilliseconds(500),
                $"A wedged dispatcher must fail fast, but rejection took {elapsed.Elapsed}.");
        }
        finally
        {
            releaseWork.Set();
        }
    }

    [Fact]
    public void WorkExceptionsPreserveTheirOriginalType()
    {
        using var dispatcher = new StaDispatcher();

        var error = Assert.Throws<InvalidOperationException>(() =>
            dispatcher.Post<int>(
                () => throw new InvalidOperationException("expected failure"),
                TimeSpan.FromSeconds(2)));

        Assert.Equal("expected failure", error.Message);
    }
}
