using System.Reflection;
using System.Text.Json;

namespace Office.Automation.Host;

internal sealed class CapabilityCatalog
{
    private const string CatalogResource = "Manifests.office-rpc.catalog.json";
    private static readonly Lazy<CapabilityCatalog> Instance = new(Load);

    private readonly IReadOnlyDictionary<string, CapabilityDefinition> _byName;
    private readonly IReadOnlyDictionary<string, RpcMethodDefinition> _rpcByName;
    private readonly IReadOnlyDictionary<string, JsonElement> _schemas;

    private CapabilityCatalog(CatalogDocument document, Dictionary<string, JsonElement> schemas)
    {
        SchemaVersion = document.SchemaVersion;
        ProtocolVersion = document.ProtocolVersion;
        Provider = document.Provider;
        CommandParamsSchema = document.CommandParamsSchema;
        SecurityPolicy = document.SecurityPolicy;
        RpcMethods = document.RpcMethods;
        Errors = document.Errors;
        Capabilities = document.Capabilities;
        _schemas = schemas;
        _byName = Capabilities.ToDictionary(capability => capability.Name, StringComparer.Ordinal);
        _rpcByName = RpcMethods.ToDictionary(method => method.Name, StringComparer.Ordinal);
    }

    internal static CapabilityCatalog Current => Instance.Value;

    internal string SchemaVersion { get; }

    internal string ProtocolVersion { get; }

    internal string Provider { get; }

    internal string CommandParamsSchema { get; }

    internal CatalogSecurityPolicy SecurityPolicy { get; }

    internal IReadOnlyList<RpcMethodDefinition> RpcMethods { get; }

    internal IReadOnlyList<CatalogErrorDefinition> Errors { get; }

    internal IReadOnlyList<CapabilityDefinition> Capabilities { get; }

    internal CapabilityDefinition? FindCapability(string name) =>
        _byName.GetValueOrDefault(name);

    internal RpcMethodDefinition? FindRpcMethod(string name) =>
        _rpcByName.GetValueOrDefault(name);

    internal IReadOnlyDictionary<string, string> ManifestCapabilities(
        string app,
        bool desktopComAvailable) =>
        Capabilities
            .Where(capability => capability.IsAvailable(app, desktopComAvailable))
            .OrderBy(capability => capability.Name, StringComparer.Ordinal)
            .ToDictionary(
                capability => capability.Name,
                capability => capability.Version,
                StringComparer.Ordinal);

    internal IReadOnlyList<string> ManifestExecutionModes(string app, bool desktopComAvailable) =>
        Capabilities
            .SelectMany(capability => capability.Availability)
            .Where(availability =>
                availability.Apps.Contains(app, StringComparer.OrdinalIgnoreCase)
                && (availability.ExecutionMode != "desktop_com" || desktopComAvailable))
            .Select(availability => availability.ExecutionMode)
            .Distinct(StringComparer.Ordinal)
            .OrderBy(mode => mode == "openxml" ? 0 : 1)
            .ThenBy(mode => mode, StringComparer.Ordinal)
            .ToArray();

    internal void ValidateInput(CapabilityDefinition capability, JsonElement input)
    {
        ValidateSchema(capability.InputSchema, input, "input");
    }

    internal void ValidateCommandParams(JsonElement parameters)
    {
        ValidateSchema(CommandParamsSchema, parameters, "params");
    }

    internal void ValidateRpcParams(RpcMethodDefinition method, JsonElement parameters)
    {
        ValidateSchema(method.ParamsSchema, parameters, "params");
    }

    internal void ValidateRpcResult(RpcMethodDefinition method, JsonElement result)
    {
        ValidateSchema(method.ResultSchema, result, "result");
    }

    internal void ValidateOutput(CapabilityDefinition capability, JsonElement output)
    {
        ValidateSchema(capability.OutputSchema, output, "result");
    }

    private void ValidateSchema(string schemaReference, JsonElement value, string path)
    {
        if (!_schemas.TryGetValue(schemaReference, out JsonElement schema))
        {
            throw new InvalidOperationException(
                $"catalog schema '{schemaReference}' was not embedded");
        }
        JsonSchemaValidator.Validate(value, schema, path);
    }

    private static CapabilityCatalog Load()
    {
        Assembly assembly = typeof(CapabilityCatalog).Assembly;
        using Stream catalogStream = assembly.GetManifestResourceStream(CatalogResource)
            ?? throw new InvalidOperationException($"missing embedded resource {CatalogResource}");
        var options = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower,
        };
        CatalogDocument document = JsonSerializer.Deserialize<CatalogDocument>(catalogStream, options)
            ?? throw new InvalidDataException("office-rpc catalog deserialized to null");
        ValidateDocument(document);

        var schemas = new Dictionary<string, JsonElement>(StringComparer.Ordinal);
        foreach (string schemaReference in document.Capabilities
            .SelectMany(capability => new[]
            {
                capability.InputSchema,
                capability.OutputSchema,
            })
            .Append(document.CommandParamsSchema)
            .Concat(document.RpcMethods.SelectMany(method => new[]
            {
                method.ParamsSchema,
                method.ResultSchema,
            }))
            .Distinct(StringComparer.Ordinal))
        {
            string resource = "Manifests." + schemaReference.Replace('/', '.');
            using Stream schemaStream = assembly.GetManifestResourceStream(resource)
                ?? throw new InvalidDataException(
                    $"catalog schema '{schemaReference}' is not embedded as '{resource}'");
            using JsonDocument schema = JsonDocument.Parse(schemaStream);
            schemas.Add(schemaReference, schema.RootElement.Clone());
        }
        return new CapabilityCatalog(document, schemas);
    }

    private static void ValidateDocument(CatalogDocument document)
    {
        if (document.SchemaVersion != "office-capability-catalog/1.2")
        {
            throw new InvalidDataException(
                $"unsupported capability catalog version '{document.SchemaVersion}'");
        }
        if (document.ProtocolVersion != HostBuildInfo.ProtocolVersion)
        {
            throw new InvalidDataException(
                $"catalog protocol '{document.ProtocolVersion}' does not match host '{HostBuildInfo.ProtocolVersion}'");
        }

        if (!document.SecurityPolicy.WorkspaceOnly)
        {
            throw new InvalidDataException("canonical security policy must confine access to the workspace");
        }
        if (document.SecurityPolicy.Actions.Count == 0
            || document.SecurityPolicy.Actions.Values.Any(action => action is not (
                "deny" or "confirm" or "checkpoint_and_confirm" or "deny_or_confirm")))
        {
            throw new InvalidDataException("canonical security policy contains an invalid action");
        }
        if (document.SecurityPolicy.Actions.GetValueOrDefault("overwrite_original")
            != "checkpoint_and_confirm")
        {
            throw new InvalidDataException(
                "canonical overwrite_original policy must require checkpoint and confirmation");
        }

        if (document.RpcMethods.Select(method => method.Name).Distinct(StringComparer.Ordinal).Count()
            != document.RpcMethods.Count)
        {
            throw new InvalidDataException("catalog RPC method names must be unique");
        }

        var errorCodes = document.Errors.Select(error => error.Code)
            .ToHashSet(StringComparer.Ordinal);
        if (errorCodes.Count != document.Errors.Count)
        {
            throw new InvalidDataException("catalog error codes must be unique");
        }
        if (document.Capabilities.Select(capability => capability.Name)
            .Distinct(StringComparer.Ordinal).Count() != document.Capabilities.Count)
        {
            throw new InvalidDataException("catalog capability names must be unique");
        }
        if (document.Capabilities.Select(capability => capability.Handler)
            .Distinct(StringComparer.Ordinal).Count() != document.Capabilities.Count)
        {
            throw new InvalidDataException("catalog handlers must be unique");
        }
        if (document.Capabilities.Select(capability => capability.McpTool)
            .Distinct(StringComparer.Ordinal).Count() != document.Capabilities.Count)
        {
            throw new InvalidDataException("catalog MCP tool mappings must be unique");
        }
        foreach (CapabilityDefinition capability in document.Capabilities)
        {
            if (!capability.HasKnownHandler)
            {
                throw new InvalidDataException(
                    $"catalog capability '{capability.Name}' has unknown handler '{capability.Handler}'");
            }
            if (capability.Errors.Any(error => !errorCodes.Contains(error)))
            {
                throw new InvalidDataException(
                    $"catalog capability '{capability.Name}' references an unknown error code");
            }
        }
    }

    private sealed class CatalogDocument
    {
        public string SchemaVersion { get; init; } = "";
        public string ProtocolVersion { get; init; } = "";
        public string Provider { get; init; } = "";
        public string CommandParamsSchema { get; init; } = "";
        public CatalogSecurityPolicy SecurityPolicy { get; init; } = new();
        public List<RpcMethodDefinition> RpcMethods { get; init; } = [];
        public List<CatalogErrorDefinition> Errors { get; init; } = [];
        public List<CapabilityDefinition> Capabilities { get; init; } = [];
    }
}

internal sealed class RpcMethodDefinition
{
    public string Name { get; init; } = "";

    public string ParamsSchema { get; init; } = "";

    public string ResultSchema { get; init; } = "";
}

internal sealed class CatalogSecurityPolicy
{
    public Dictionary<string, string> Actions { get; init; } = new(StringComparer.Ordinal);

    public bool WorkspaceOnly { get; init; }

    public Dictionary<string, List<string>> ExecuteMsoAllowlist { get; init; } =
        new(StringComparer.OrdinalIgnoreCase);

    public List<string> ExecuteMsoConfirm { get; init; } = [];
}

internal sealed class CatalogErrorDefinition
{
    public string Code { get; init; } = "";

    public bool Retryable { get; init; }
}

internal sealed class CapabilityDefinition
{
    public string Name { get; init; } = "";
    public string Version { get; init; } = "";
    public string Handler { get; init; } = "";
    public string McpTool { get; init; } = "";
    public string InputSchema { get; init; } = "";
    public string OutputSchema { get; init; } = "";
    public List<CapabilityAvailability> Availability { get; init; } = [];
    public List<string> Errors { get; init; } = [];

    internal bool HasKnownHandler => Handler is
        "deck_compile" or
        "document_inspect" or
        "batch_convert" or
        "batch_replace_text" or
        "slide_render";

    internal CapabilityHandler HandlerId => Handler switch
    {
        "deck_compile" => CapabilityHandler.DeckCompile,
        "document_inspect" => CapabilityHandler.DocumentInspect,
        "batch_convert" => CapabilityHandler.BatchConvert,
        "batch_replace_text" => CapabilityHandler.BatchReplaceText,
        "slide_render" => CapabilityHandler.SlideRender,
        _ => throw new InvalidOperationException($"unknown catalog handler '{Handler}'"),
    };

    internal bool IsAvailable(string app, bool desktopComAvailable) =>
        Availability.Any(availability =>
            availability.Apps.Contains(app, StringComparer.OrdinalIgnoreCase)
            && (availability.ExecutionMode != "desktop_com" || desktopComAvailable));
}

internal sealed class CapabilityAvailability
{
    public string ExecutionMode { get; init; } = "";
    public List<string> Apps { get; init; } = [];
}

internal enum CapabilityHandler
{
    DeckCompile,
    DocumentInspect,
    BatchConvert,
    BatchReplaceText,
    SlideRender,
}
