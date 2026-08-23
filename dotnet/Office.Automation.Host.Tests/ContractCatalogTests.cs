using System.Text.Json;
using Office.Automation.Com;
using Office.Automation.Host;
using Xunit;

namespace Office.Automation.Host.Tests;

public sealed class ContractCatalogTests
{
    [Fact]
    public void CatalogOwnsManifestDispatchSchemasAndErrorCodes()
    {
        CapabilityCatalog catalog = CapabilityCatalog.Current;
        Assert.Equal("office-capability-catalog/1.2", catalog.SchemaVersion);
        Assert.Equal("office-rpc/1", catalog.ProtocolVersion);
        Assert.Equal(5, catalog.Capabilities.Count);
        Assert.Equal(
            [
                "office.command.execute",
                "office.host.handshake",
                "office.host.ping",
                "office.host.shutdown",
                "office.job.cancel",
                "office.job.get",
            ],
            catalog.RpcMethods.Select(method => method.Name).Order());
        Assert.Equal(
            catalog.Errors.Select(error => error.Code).Order(),
            Enum.GetValues<OfficeErrorCode>().Select(error => error.ToWireName()).Order());

        IReadOnlyDictionary<string, string> expected =
            catalog.ManifestCapabilities("powerpoint", desktopComAvailable: false);
        using var router = new CommandRouter("powerpoint", enableDesktopCom: false);
        using JsonDocument response = JsonDocument.Parse(router.Dispatch(
            """
            {"jsonrpc":"2.0","id":1,"method":"office.host.handshake","params":{"requested_app":"powerpoint"}}
            """));
        var actual = response.RootElement
            .GetProperty("result")
            .GetProperty("capability_manifest")
            .GetProperty("capabilities")
            .EnumerateObject()
            .ToDictionary(property => property.Name, property => property.Value.GetString()!);
        Assert.Equal(expected, actual);
    }

    [Theory]
    [InlineData("deck_compile", "{\"ir\":\"{}\",\"output\":\"deck.pptx\"}")]
    [InlineData("document_inspect", "{\"path\":\"deck.pptx\",\"backend\":\"openxml\"}")]
    [InlineData("batch_convert", "{\"inputs\":[\"deck.pptx\"],\"output_directory\":\"out\"}")]
    [InlineData("batch_replace_text", "{\"inputs\":[\"deck.pptx\"],\"rules\":[{\"find\":\"a\",\"replace\":\"b\"}]}")]
    [InlineData("slide_render", "{\"path\":\"deck.pptx\",\"output_directory\":\"out\"}")]
    public void EveryCatalogSchemaAcceptsItsMinimalInput(string handler, string json)
    {
        CapabilityDefinition definition = Assert.Single(
            CapabilityCatalog.Current.Capabilities,
            capability => capability.Handler == handler);
        using JsonDocument input = JsonDocument.Parse(json);

        CapabilityCatalog.Current.ValidateInput(definition, input.RootElement);
    }

    [Theory]
    [InlineData("document_inspect", "{\"path\":\"deck.pptx\",\"backend\":\"guess\"}", "input.backend")]
    [InlineData("batch_replace_text", "{\"inputs\":[\"deck.pptx\"],\"rules\":[{\"find\":\"\",\"replace\":\"b\"}]}", "input.rules[0].find")]
    [InlineData("slide_render", "{\"path\":\"deck.pptx\",\"output_directory\":\"out\",\"width\":32}", "input.width")]
    public void CatalogSchemasRejectInvalidTypesEnumsAndBounds(
        string handler,
        string json,
        string expectedPath)
    {
        CapabilityDefinition definition = Assert.Single(
            CapabilityCatalog.Current.Capabilities,
            capability => capability.Handler == handler);
        using JsonDocument input = JsonDocument.Parse(json);

        OfficeArgumentException error = Assert.Throws<OfficeArgumentException>(() =>
            CapabilityCatalog.Current.ValidateInput(definition, input.RootElement));

        Assert.Contains(expectedPath, error.Message);
    }

    [Fact]
    public void RouterRejectsInputOutsideTheCatalogSchemaBeforeDispatch()
    {
        using var router = new CommandRouter("powerpoint", enableDesktopCom: false);

        using JsonDocument response = JsonDocument.Parse(router.Dispatch(
            """
            {"jsonrpc":"2.0","id":7,"method":"office.command.execute","params":{"capability":"deck.compile","input":{"ir":"{}","output":"deck.pptx","surprise":true},"policy":{}}}
            """));

        Assert.Equal(
            "OFFICE_INVALID_REQUEST",
            response.RootElement.GetProperty("error").GetProperty("code").GetString());
        Assert.Contains(
            "input.surprise",
            response.RootElement.GetProperty("error").GetProperty("message").GetString());
    }

    [Theory]
    [InlineData("{")]
    [InlineData("[]")]
    [InlineData("{\"jsonrpc\":\"2.0\",\"id\":1,\"params\":{}}")]
    public void MalformedJsonRpcUsesTheTypedInvalidRequestCode(string request)
    {
        using var router = new CommandRouter("powerpoint", enableDesktopCom: false);

        using JsonDocument response = JsonDocument.Parse(router.Dispatch(request));

        Assert.Equal(
            OfficeErrorCode.OfficeInvalidRequest.ToWireName(),
            response.RootElement.GetProperty("error").GetProperty("code").GetString());
    }

    [Fact]
    public void CatalogAcceptedButUnavailableBackendIsExplicitlyUnsupported()
    {
        using var router = new CommandRouter("powerpoint", enableDesktopCom: false);

        using JsonDocument response = JsonDocument.Parse(router.Dispatch(
            """
            {"jsonrpc":"2.0","id":8,"method":"office.command.execute","params":{"capability":"batch.convert","input":{"inputs":["deck.pptx"],"output_directory":"out","backend":"graph"},"policy":{}}}
            """));

        Assert.Equal(
            OfficeErrorCode.OfficeCapabilityUnsupported.ToWireName(),
            response.RootElement.GetProperty("error").GetProperty("code").GetString());
        Assert.Contains(
            "not implemented by the desktop sidecar",
            response.RootElement.GetProperty("error").GetProperty("message").GetString());
    }

    [Fact]
    public void OutputModesHonorTheBatchConvertSchema()
    {
        string temp = Path.Combine(Path.GetTempPath(), $"dcc-output-mode-{Guid.NewGuid():N}");
        Directory.CreateDirectory(temp);
        string existing = Path.Combine(temp, "deck.pdf");
        File.WriteAllText(existing, "existing");
        try
        {
            Assert.Equal(
                Path.Combine(temp, "deck.v2.pdf"),
                CommandRouter.OutputPathForMode(temp, "deck.pdf", "versioned"));
            Assert.Equal(
                existing,
                CommandRouter.OutputPathForMode(temp, "deck.pdf", "overwrite"));
            OfficeArgumentException error = Assert.Throws<OfficeArgumentException>(() =>
                CommandRouter.OutputPathForMode(temp, "deck.pdf", "fail"));
            Assert.Contains("overwrite is 'fail'", error.Message);
        }
        finally
        {
            Directory.Delete(temp, recursive: true);
        }
    }

    [Fact]
    public void JobStatusSchemaAcceptsNullsButRejectsWrongNullableTypes()
    {
        RpcMethodDefinition method = CapabilityCatalog.Current.FindRpcMethod("office.job.get")!;
        using JsonDocument invalid = JsonDocument.Parse(
            """
            {
              "job_id":"job:0123456789abcdef0123456789abcdef",
              "capability":"batch.convert",
              "phase":"failed",
              "stage":"failed",
              "completed":0,
              "total":1,
              "cancel_requested":false,
              "created_at":"2026-08-24T00:00:00Z",
              "updated_at":"2026-08-24T00:00:01Z",
              "result":null,
              "error":"not-an-error-object"
            }
            """);

        OfficeArgumentException error = Assert.Throws<OfficeArgumentException>(() =>
            CapabilityCatalog.Current.ValidateRpcResult(method, invalid.RootElement));

        Assert.Contains("result.error must be object or null", error.Message);
    }
}
