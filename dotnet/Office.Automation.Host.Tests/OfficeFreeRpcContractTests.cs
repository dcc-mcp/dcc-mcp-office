using System.Diagnostics;
using System.Reflection;
using System.Text.Json;
using System.Text.Json.Nodes;
using Office.Automation.Host;
using Xunit;

namespace Office.Automation.Host.Tests;

public sealed class OfficeFreeRpcContractTests
{
    private const string MinimalDeckIr = """
        {
          "schema_version": "office-ir/1.0",
          "kind": "presentation",
          "document_id": "draft:office-free-contract",
          "metadata": {"title": "Office-free Contract", "language": "en-US"},
          "document": {
            "slides": [
              {
                "semantic_layout": "title_cover",
                "title": "Office-free Contract",
                "content_blocks": [{"type": "text", "paragraphs": ["fixture"]}],
                "speaker_notes": "fixture"
              }
            ]
          },
          "outputs": ["pptx"]
        }
        """;

    [Fact]
    public void RouterCanDisableDesktopComForOfficeFreeFixtures()
    {
        using var router = new CommandRouter("powerpoint", enableDesktopCom: false);

        JsonNode response = Dispatch(router, "handshake.request.json");

        JsonArray modes = response["result"]!["capability_manifest"]!["execution_modes"]!
            .AsArray();
        Assert.Equal(["openxml"], modes.Select(value => value!.GetValue<string>()));
    }

    [Fact]
    public void StdioModeReplaysGoldenWireFixturesWithoutOffice()
    {
        TextReader originalIn = Console.In;
        TextWriter originalOut = Console.Out;
        TextWriter originalError = Console.Error;
        var output = new StringWriter();
        var error = new StringWriter();
        try
        {
            string requests = string.Join(Environment.NewLine,
                Fixture("handshake.request.json"),
                Fixture("unknown-method.request.json"),
                Fixture("policy-denied.request.json"));
            Console.SetIn(new StringReader(requests));
            Console.SetOut(output);
            Console.SetError(error);

            int exitCode = Program.Main(["--app=powerpoint", "--stdio", "--openxml-only"]);

            Assert.Equal(0, exitCode);
        }
        finally
        {
            Console.SetIn(originalIn);
            Console.SetOut(originalOut);
            Console.SetError(originalError);
        }

        string[] messages = output.ToString().Split(
            Environment.NewLine,
            StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        JsonNode[] parsed = messages.Select(Parse).ToArray();
        string[] responses = parsed
            .Where(message => message["id"] is not null)
            .Select(message => message.ToJsonString())
            .ToArray();
        Assert.Equal(3, responses.Length);
        AssertJsonEqual(MaterializeExpected("handshake.expected.json"), responses[0]);
        AssertJsonEqual(Fixture("unknown-method.expected.json"), responses[1]);
        AssertJsonEqual(Fixture("policy-denied.expected.json"), responses[2]);
        JsonNode applicationStarted = Assert.Single(parsed, message =>
            message["method"]?.GetValue<string>() == "office.application.started");
        Assert.Equal(
            "host:handshake",
            applicationStarted["params"]!["correlation_id"]!.GetValue<string>());
        Assert.Equal(string.Empty, error.ToString());
    }

    [Fact]
    public void DeckCompileFixtureProducesInspectablePptxWithoutOffice()
    {
        string temp = Path.Combine(
            Path.GetTempPath(),
            $"dcc-office-free-contract-{Guid.NewGuid():N}");
        string output = Path.Combine(temp, "fixture.pptx");
        try
        {
            Directory.CreateDirectory(temp);
            string request = Fixture("deck-compile.request.json")
                .Replace(
                    JsonSerializer.Serialize("{{IR_JSON}}"),
                    JsonSerializer.Serialize(MinimalDeckIr),
                    StringComparison.Ordinal)
                .Replace(
                    JsonSerializer.Serialize("{{OUTPUT_PATH}}"),
                    JsonSerializer.Serialize(output),
                    StringComparison.Ordinal)
                .Replace(
                    JsonSerializer.Serialize("{{WORKSPACE_ROOT}}"),
                    JsonSerializer.Serialize(temp),
                    StringComparison.Ordinal);
            using var router = new CommandRouter(
                "powerpoint",
                enableDesktopCom: false,
                workspaceRoot: temp);

            JsonNode response = Parse(router.Dispatch(request));

            Assert.True(response["result"] is not null, response.ToJsonString());
            Assert.Equal("openxml", response["result"]!["backend"]!.GetValue<string>());
            Assert.Equal(1, response["result"]!["changed"]!["slides"]!.GetValue<int>());
            Assert.True(File.Exists(output));

            string inspect = JsonSerializer.Serialize(new
            {
                jsonrpc = "2.0",
                id = 5,
                method = "office.command.execute",
                @params = new
                {
                    capability = "document.inspect",
                    input = new { path = output, backend = "openxml" },
                    policy = new { workspace_root = temp },
                },
            });
            JsonNode inspected = Parse(router.Dispatch(inspect));
            Assert.Equal(
                1,
                inspected["result"]!["changed"]!["summary"]!["slide_count"]!
                    .GetValue<int>());
        }
        finally
        {
            if (Directory.Exists(temp))
            {
                Directory.Delete(temp, recursive: true);
            }
        }
    }

    private static JsonNode Dispatch(CommandRouter router, string fixture) =>
        Parse(router.Dispatch(Fixture(fixture)));

    private static string Fixture(string name) =>
        File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", name));

    private static string MaterializeExpected(string name)
    {
        string informationalVersion = typeof(CommandRouter).Assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()!
            .InformationalVersion;
        string providerVersion = informationalVersion.Split('+')[0];
        string hostId =
            $"office-host:powerpoint:session-{Process.GetCurrentProcess().SessionId}";
        return Fixture(name)
            .Replace("{{PROVIDER_VERSION}}", providerVersion, StringComparison.Ordinal)
            .Replace("{{HOST_ID}}", hostId, StringComparison.Ordinal);
    }

    private static JsonNode Parse(string json) =>
        JsonNode.Parse(json) ?? throw new InvalidOperationException("fixture produced null JSON");

    private static void AssertJsonEqual(string expected, string actual) =>
        Assert.True(
            JsonNode.DeepEquals(Parse(expected), Parse(actual)),
            $"Expected:{Environment.NewLine}{expected}{Environment.NewLine}Actual:{Environment.NewLine}{actual}");
}
