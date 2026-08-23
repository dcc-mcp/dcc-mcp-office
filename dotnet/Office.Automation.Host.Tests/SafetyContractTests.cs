using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Nodes;
using Office.Automation.Com;
using Office.Automation.Host;
using Xunit;

namespace Office.Automation.Host.Tests;

public sealed class SafetyContractTests
{
    private const string MinimalDeckIr = """
        {
          "schema_version": "office-ir/1.0",
          "kind": "presentation",
          "document_id": "draft:safety-contract",
          "metadata": {"title": "Safety Contract", "language": "en-US"},
          "document": {
            "slides": [{
              "semantic_layout": "title_cover",
              "title": "Safety Contract",
              "content_blocks": [{"type": "text", "paragraphs": ["fixture"]}],
              "speaker_notes": "fixture"
            }]
          },
          "outputs": ["pptx"]
        }
        """;

    [Fact]
    public void ExpectedRevisionIsRejectedInsteadOfIgnored()
    {
        string temp = TempDirectory();
        try
        {
            string request = ExecuteRequest(
                "deck.compile",
                new { ir = MinimalDeckIr, output = Path.Combine(temp, "deck.pptx") },
                temp,
                document: new { document_id = "deck-1", expected_revision = 7UL });
            using var router = new CommandRouter(
                "powerpoint",
                enableDesktopCom: false,
                workspaceRoot: temp);

            JsonNode response = Parse(router.Dispatch(request));

            Assert.Equal(
                OfficeErrorCode.OfficeCapabilityUnsupported.ToWireName(),
                response["error"]!["code"]!.GetValue<string>());
            Assert.Contains("expected_revision", response["error"]!["message"]!.GetValue<string>());
            Assert.False(File.Exists(Path.Combine(temp, "deck.pptx")));
        }
        finally
        {
            Directory.Delete(temp, recursive: true);
        }
    }

    [Fact]
    public void CommitRequiresStructuredConfirmationBeforeComAttach()
    {
        string temp = TempDirectory();
        string source = Path.Combine(temp, "deck.pptx");
        File.WriteAllText(source, "fixture");
        try
        {
            string request = ExecuteRequest(
                "batch.replace_text",
                new
                {
                    inputs = new[] { source },
                    rules = new[] { new { find = "fixture", replace = "updated" } },
                    dry_run = false,
                },
                temp);
            using var router = new CommandRouter(
                "powerpoint",
                enableDesktopCom: false,
                workspaceRoot: temp);

            JsonNode response = Parse(router.Dispatch(request));

            Assert.Equal(
                OfficeErrorCode.OfficeUserConfirmationRequired.ToWireName(),
                response["error"]!["code"]!.GetValue<string>());
            Assert.Contains("overwrite_original", response["error"]!["message"]!.GetValue<string>());
        }
        finally
        {
            Directory.Delete(temp, recursive: true);
        }
    }

    [Fact]
    public void PdfOverwriteRequiresStructuredConfirmationBeforeComAttach()
    {
        string temp = TempDirectory();
        string source = Path.Combine(temp, "deck.pptx");
        File.WriteAllText(source, "fixture");
        try
        {
            string request = ExecuteRequest(
                "batch.convert",
                new
                {
                    inputs = new[] { source },
                    output_directory = temp,
                    overwrite = "overwrite",
                },
                temp);
            using var router = new CommandRouter(
                "powerpoint",
                enableDesktopCom: false,
                workspaceRoot: temp);

            JsonNode response = Parse(router.Dispatch(request));

            Assert.Equal(
                OfficeErrorCode.OfficeUserConfirmationRequired.ToWireName(),
                response["error"]!["code"]!.GetValue<string>());
            Assert.Contains("overwrite_original", response["error"]!["message"]!.GetValue<string>());
        }
        finally
        {
            Directory.Delete(temp, recursive: true);
        }
    }

    [Fact]
    public void CheckpointIsAByteExactPreimageArtifact()
    {
        string temp = TempDirectory();
        string source = Path.Combine(temp, "deck.pptx");
        byte[] original = [0, 1, 2, 3, 254, 255];
        File.WriteAllBytes(source, original);
        try
        {
            JsonObject artifact = CommandRouter.CreateCheckpoint(source, "operation-42");
            string checkpoint = artifact["path"]!.GetValue<string>();

            Assert.NotEqual(source, checkpoint);
            Assert.Equal("checkpoint", artifact["kind"]!.GetValue<string>());
            Assert.Equal(original, File.ReadAllBytes(checkpoint));
            Assert.Equal(
                Convert.ToHexString(SHA256.HashData(original)).ToLowerInvariant(),
                artifact["sha256"]!.GetValue<string>());
        }
        finally
        {
            Directory.Delete(temp, recursive: true);
        }
    }

    [Fact]
    public void CompileRefusesToOverwriteAnExistingOutput()
    {
        string temp = TempDirectory();
        string output = Path.Combine(temp, "deck.pptx");
        byte[] original = [7, 8, 9];
        File.WriteAllBytes(output, original);
        try
        {
            string request = ExecuteRequest(
                "deck.compile",
                new { ir = MinimalDeckIr, output },
                temp);
            using var router = new CommandRouter(
                "powerpoint",
                enableDesktopCom: false,
                workspaceRoot: temp);

            JsonNode response = Parse(router.Dispatch(request));

            Assert.Equal(
                OfficeErrorCode.OfficeAccessDenied.ToWireName(),
                response["error"]!["code"]!.GetValue<string>());
            Assert.Equal(original, File.ReadAllBytes(output));
        }
        finally
        {
            Directory.Delete(temp, recursive: true);
        }
    }

    [Fact]
    public void WorkspaceTraversalIsDeniedBeforeAFileIsWritten()
    {
        string workspace = TempDirectory();
        string outside = TempDirectory();
        string output = Path.Combine(outside, "escaped.pptx");
        try
        {
            string request = ExecuteRequest(
                "deck.compile",
                new { ir = MinimalDeckIr, output },
                workspace);
            using var router = new CommandRouter(
                "powerpoint",
                enableDesktopCom: false,
                workspaceRoot: workspace);

            JsonNode response = Parse(router.Dispatch(request));

            Assert.Equal(
                OfficeErrorCode.OfficeAccessDenied.ToWireName(),
                response["error"]!["code"]!.GetValue<string>());
            Assert.False(File.Exists(output));
        }
        finally
        {
            Directory.Delete(workspace, recursive: true);
            Directory.Delete(outside, recursive: true);
        }
    }

    [Fact]
    public void ExtendedLengthWorkspaceRootMatchesTheNormalChildPath()
    {
        if (!OperatingSystem.IsWindows())
        {
            return;
        }
        string workspace = TempDirectory();
        try
        {
            string extendedRoot = $@"\\?\{workspace}";
            string output = Path.Combine(workspace, "deck.pptx");

            WorkspaceGuard.ValidatePaths([output], extendedRoot);
        }
        finally
        {
            Directory.Delete(workspace, recursive: true);
        }
    }

    [Fact]
    public void RequestCannotReplaceTheHostBoundWorkspaceRoot()
    {
        string workspace = TempDirectory();
        string outside = TempDirectory();
        string output = Path.Combine(outside, "escaped.pptx");
        try
        {
            string request = ExecuteRequest(
                "deck.compile",
                new { ir = MinimalDeckIr, output },
                outside);
            using var router = new CommandRouter(
                "powerpoint",
                enableDesktopCom: false,
                workspaceRoot: workspace);

            JsonNode response = Parse(router.Dispatch(request));

            Assert.Equal(
                OfficeErrorCode.OfficeAccessDenied.ToWireName(),
                response["error"]!["code"]!.GetValue<string>());
            Assert.Contains("workspace_root", response["error"]!["message"]!.GetValue<string>());
            Assert.False(File.Exists(output));
        }
        finally
        {
            Directory.Delete(workspace, recursive: true);
            Directory.Delete(outside, recursive: true);
        }
    }

    [Fact]
    public void WorkspaceSymlinkCannotRedirectAWriteOutsideTheBoundRoot()
    {
        string workspace = TempDirectory();
        string outside = TempDirectory();
        string link = Path.Combine(workspace, "redirect");
        string output = Path.Combine(link, "escaped.pptx");
        try
        {
            Directory.CreateSymbolicLink(link, outside);
            string request = ExecuteRequest(
                "deck.compile",
                new { ir = MinimalDeckIr, output },
                workspace);
            using var router = new CommandRouter(
                "powerpoint",
                enableDesktopCom: false,
                workspaceRoot: workspace);

            JsonNode response = Parse(router.Dispatch(request));

            Assert.Equal(
                OfficeErrorCode.OfficeAccessDenied.ToWireName(),
                response["error"]!["code"]!.GetValue<string>());
            Assert.False(File.Exists(Path.Combine(outside, "escaped.pptx")));
        }
        finally
        {
            if (Directory.Exists(link))
            {
                Directory.Delete(link);
            }
            Directory.Delete(workspace, recursive: true);
            Directory.Delete(outside, recursive: true);
        }
    }

    [Fact]
    public void OpenXmlAuditReportsObservedPostureWithoutClaimingComEnforcement()
    {
        string temp = TempDirectory();
        string output = Path.Combine(temp, "deck.pptx");
        try
        {
            string request = ExecuteRequest(
                "deck.compile",
                new { ir = MinimalDeckIr, output },
                temp);
            using var router = new CommandRouter(
                "powerpoint",
                enableDesktopCom: false,
                workspaceRoot: temp);

            JsonNode response = Parse(router.Dispatch(request));
            Assert.True(response["result"] is not null, response.ToJsonString());
            JsonNode security = response["result"]!["audit"]!["security"]!;

            Assert.False(
                security["automation_security"]!["applicable"]!.GetValue<bool>());
            Assert.False(
                security["automation_security"]!["enforced"]!.GetValue<bool>());
            Assert.True(security["workspace_only"]!.GetValue<bool>());
            Assert.True(File.Exists(output));
        }
        finally
        {
            Directory.Delete(temp, recursive: true);
        }
    }

    [Fact]
    public void SessionZeroCannotStartDesktopAutomation()
    {
        Assert.False(Program.IsDesktopAutomationAllowed(sessionId: 0));
        Assert.True(Program.IsDesktopAutomationAllowed(sessionId: 1));
    }

    [Fact]
    public void IndeterminateTimeoutIsMachineReadableOnTheWire()
    {
        JsonElement id = JsonSerializer.SerializeToElement(9);

        JsonNode response = Parse(CommandRouter.Error(
            id,
            OfficeErrorCode.OfficeRpcTimeout,
            "write may have completed",
            indeterminate: true));

        Assert.True(response["error"]!["data"]!["indeterminate"]!.GetValue<bool>());
    }

    private static string ExecuteRequest(
        string capability,
        object input,
        string workspaceRoot,
        object? document = null) =>
        JsonSerializer.Serialize(new
        {
            jsonrpc = "2.0",
            id = 1,
            method = "office.command.execute",
            @params = new
            {
                capability,
                document,
                input,
                policy = new { workspace_root = workspaceRoot },
            },
        });

    private static string TempDirectory()
    {
        string path = Path.Combine(Path.GetTempPath(), $"dcc-safety-{Guid.NewGuid():N}");
        Directory.CreateDirectory(path);
        return path;
    }

    private static JsonNode Parse(string json) =>
        JsonNode.Parse(json) ?? throw new InvalidOperationException("response was null");
}
