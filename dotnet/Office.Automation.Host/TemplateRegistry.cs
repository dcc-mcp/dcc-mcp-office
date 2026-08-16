using System.Text.Json;

namespace Office.Automation.Host;

/// <summary>
/// Brand template registry (proposal §15.4 / templates/README.md): resolves
/// brand:// URIs to template packages. The M1 host ships the built-in
/// dcc-mcp/default package (embedded Open XML skeletons); the registry file
/// is the source of truth for what else exists and which semantic layouts a
/// package materialises.
/// </summary>
public sealed class TemplateRegistry
{
    /// <summary>The package every deck.compile falls back to.</summary>
    public const string DefaultUri = "brand://dcc-mcp/default";

    private readonly Dictionary<string, TemplateEntry> _entries;

    public TemplateRegistry()
    {
        _entries = new Dictionary<string, TemplateEntry>(StringComparer.OrdinalIgnoreCase);
        using var document = JsonDocument.Parse(LoadRegistryJson());
        foreach (var template in document.RootElement.GetProperty("templates").EnumerateArray())
        {
            var uri = template.GetProperty("uri").GetString()!;
            _entries[uri] = new TemplateEntry(
                uri,
                template.GetProperty("version").GetString()!,
                template.TryGetProperty("kind", out var kind) ? kind.GetString() ?? "presentation" : "presentation",
                template.TryGetProperty("source", out var source) ? source.GetString() ?? "" : "",
                template.TryGetProperty("layouts", out var layouts)
                    ? layouts.EnumerateArray().Select(l => l.GetString() ?? "").Where(l => l.Length > 0).ToArray()
                    : Array.Empty<string>());
        }
    }

    /// <summary>Resolves a brand:// URI, or null when unknown.</summary>
    public TemplateEntry? Resolve(string uri) =>
        _entries.TryGetValue(uri, out var entry) ? entry : null;

    /// <summary>The built-in package (always present).</summary>
    public TemplateEntry Default =>
        _entries.TryGetValue(DefaultUri, out var entry)
            ? entry
            : throw new InvalidOperationException("brand registry lacks the built-in default package");

    public IReadOnlyCollection<string> AllUris => _entries.Keys;

    private static string LoadRegistryJson()
    {
        // 1) packaged next to the exe (deployment layout: exe-dir/templates/registry.json)
        var packaged = Path.Combine(AppContext.BaseDirectory, "templates", "registry.json");
        if (File.Exists(packaged))
        {
            return File.ReadAllText(packaged);
        }
        // 2) repo-relative (dev: cwd is the repository root under dotnet run)
        var repoRelative = Path.Combine(Directory.GetCurrentDirectory(), "templates", "registry.json");
        if (File.Exists(repoRelative))
        {
            return File.ReadAllText(repoRelative);
        }
        // 3) embedded copy (single-file publish keeps working)
        using var stream = typeof(TemplateRegistry).Assembly
            .GetManifestResourceStream("Templates.registry.json")
            ?? throw new InvalidOperationException("embedded brand registry missing");
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

}

/// <summary>One registry entry.</summary>
public sealed record TemplateEntry(
    string Uri,
    string Version,
    string Kind,
    string Source,
    string[] Layouts);
