using System.Text.Json.Serialization;
using Office.Automation.OpenXml;

namespace Office.Automation.Host;

/// <summary>
/// Resolves only materialized brand templates. A URI is advertised after its
/// package has passed path, schema, XML-part, layout, and style validation.
/// </summary>
public sealed class TemplateRegistry
{
    public const string DefaultUri = PresentationTemplatePackage.DefaultUri;

    private readonly Dictionary<string, TemplateEntry> _entries =
        new(StringComparer.OrdinalIgnoreCase);

    public TemplateRegistry(
        IEnumerable<string>? templateDirectories = null,
        bool includeDefaultDirectories = false)
    {
        Add(new TemplateEntry(
            PresentationTemplatePackage.EmbeddedDefault(),
            SourceKind: "embedded"));

        foreach (string directory in TemplateDirectories(
            templateDirectories,
            includeDefaultDirectories))
        {
            if (!Directory.Exists(directory))
            {
                continue;
            }
            foreach (string packagePath in Directory.EnumerateFiles(
                directory,
                "package.json",
                new EnumerationOptions
                {
                    AttributesToSkip = FileAttributes.ReparsePoint,
                    IgnoreInaccessible = true,
                    RecurseSubdirectories = true,
                }).Order(StringComparer.OrdinalIgnoreCase))
            {
                Add(new TemplateEntry(
                    PresentationTemplatePackage.LoadDirectory(
                        Path.GetDirectoryName(packagePath)!),
                    SourceKind: "file"));
            }
        }
    }

    public TemplateEntry? Resolve(string uri) =>
        _entries.TryGetValue(uri, out TemplateEntry? entry) ? entry : null;

    public IReadOnlyCollection<string> AllUris =>
        _entries.Keys.Order(StringComparer.OrdinalIgnoreCase).ToArray();

    public IReadOnlyDictionary<string, TemplatePackageCapability> Capabilities => _entries
        .OrderBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)
        .ToDictionary(
            pair => pair.Key,
            pair => new TemplatePackageCapability(
                pair.Value.Package.Version,
                "presentation",
                pair.Value.SourceKind,
                pair.Value.Package.Layouts),
            StringComparer.OrdinalIgnoreCase);

    private void Add(TemplateEntry entry)
    {
        if (!_entries.TryAdd(entry.Package.Uri, entry))
        {
            throw new InvalidDataException(
                $"duplicate materialized template URI '{entry.Package.Uri}'");
        }
    }

    private static IEnumerable<string> TemplateDirectories(
        IEnumerable<string>? configured,
        bool includeDefaults)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string directory in configured ?? Array.Empty<string>())
        {
            if (string.IsNullOrWhiteSpace(directory))
            {
                throw new InvalidDataException("template directory must not be empty");
            }
            string full = Path.GetFullPath(directory);
            if (seen.Add(full))
            {
                yield return full;
            }
        }

        if (!includeDefaults)
        {
            yield break;
        }

        string packaged = Path.Combine(AppContext.BaseDirectory, "templates");
        if (seen.Add(packaged))
        {
            yield return packaged;
        }
        string bundled = Path.GetFullPath(
            Path.Combine(AppContext.BaseDirectory, "..", "templates"));
        if (seen.Add(bundled))
        {
            yield return bundled;
        }

        string local = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "dcc-mcp",
            "office-templates");
        if (seen.Add(local))
        {
            yield return local;
        }
    }
}

public sealed record TemplateEntry(
    PresentationTemplatePackage Package,
    string SourceKind);

public sealed record TemplatePackageCapability(
    [property: JsonPropertyName("version")] string Version,
    [property: JsonPropertyName("kind")] string Kind,
    [property: JsonPropertyName("source_kind")] string SourceKind,
    [property: JsonPropertyName("layouts")] IReadOnlyCollection<string> Layouts);
