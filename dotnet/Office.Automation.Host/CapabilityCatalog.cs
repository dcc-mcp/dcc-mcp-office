using System.Reflection;
using System.Text.Json;

namespace Office.Automation.Host;

internal sealed class CapabilityCatalog
{
    private const string CatalogResource = "Manifests.office-rpc.catalog.json";
    private static readonly Lazy<CapabilityCatalog> Instance = new(Load);

    private readonly IReadOnlyDictionary<string, CapabilityDefinition> _byName;
    private readonly IReadOnlyDictionary<string, JsonElement> _schemas;

    private CapabilityCatalog(CatalogDocument document, Dictionary<string, JsonElement> schemas)
    {
        SchemaVersion = document.SchemaVersion;
        ProtocolVersion = document.ProtocolVersion;
        Provider = document.Provider;
        Errors = document.Errors;
        Capabilities = document.Capabilities;
        _schemas = schemas;
        _byName = Capabilities.ToDictionary(capability => capability.Name, StringComparer.Ordinal);
    }

    internal static CapabilityCatalog Current => Instance.Value;

    internal string SchemaVersion { get; }

    internal string ProtocolVersion { get; }

    internal string Provider { get; }

    internal IReadOnlyList<CatalogErrorDefinition> Errors { get; }

    internal IReadOnlyList<CapabilityDefinition> Capabilities { get; }

    internal CapabilityDefinition? FindCapability(string name) =>
        _byName.GetValueOrDefault(name);

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
        if (document.SchemaVersion != "office-capability-catalog/1.0")
        {
            throw new InvalidDataException(
                $"unsupported capability catalog version '{document.SchemaVersion}'");
        }
        if (document.ProtocolVersion != HostBuildInfo.ProtocolVersion)
        {
            throw new InvalidDataException(
                $"catalog protocol '{document.ProtocolVersion}' does not match host '{HostBuildInfo.ProtocolVersion}'");
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
        public List<CatalogErrorDefinition> Errors { get; init; } = [];
        public List<CapabilityDefinition> Capabilities { get; init; } = [];
    }
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
