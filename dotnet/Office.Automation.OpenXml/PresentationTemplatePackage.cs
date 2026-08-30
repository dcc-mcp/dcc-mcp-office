using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace Office.Automation.OpenXml;

/// <summary>
/// Materialized presentation template package. External packages inherit the
/// embedded default skeleton and may override XML parts, semantic layout
/// mappings, brand style, and logo without rebuilding the Host.
/// </summary>
public sealed class PresentationTemplatePackage
{
    public const string SchemaVersion = "office-template-package/1.0";
    public const string DefaultUri = "brand://dcc-mcp/default";

    private static readonly string[] PartNames =
    [
        "slide_master",
        "slide_layout",
        "theme",
        "slide",
        "notes_master",
        "notes_slide",
    ];

    private static readonly IReadOnlyDictionary<string, string> ExpectedRoots =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["slide_master"] = "sldMaster",
            ["slide_layout"] = "sldLayout",
            ["theme"] = "theme",
            ["slide"] = "sld",
            ["notes_master"] = "notesMaster",
            ["notes_slide"] = "notes",
        };

    private static readonly string[] CanonicalRenderers =
    [
        "title_cover",
        "section_cover",
        "two_columns",
        "comparison",
        "timeline",
        "kpi_dashboard",
        "technical_architecture",
        "image_left_text_right",
        "image_grid",
        "closing",
        "bullets",
    ];

    private static readonly HashSet<string> Renderers = new(
        CanonicalRenderers,
        StringComparer.Ordinal);

    private readonly IReadOnlyDictionary<string, string> _parts;
    private readonly IReadOnlyDictionary<string, string> _layouts;

    private PresentationTemplatePackage(
        string uri,
        string version,
        string brandName,
        IReadOnlyDictionary<string, string> layouts,
        IReadOnlyDictionary<string, string> parts,
        PresentationTemplateStyle style,
        byte[]? logo)
    {
        Uri = uri;
        Version = version;
        BrandName = brandName;
        _layouts = layouts;
        _parts = parts;
        Style = style;
        Logo = logo;
    }

    public string Uri { get; }

    public string Version { get; }

    public string BrandName { get; }

    public IReadOnlyCollection<string> Layouts => _layouts.Keys.Order(StringComparer.Ordinal).ToArray();

    public PresentationTemplateStyle Style { get; }

    public byte[]? Logo { get; }

    public string ResolveRenderer(string semanticLayout) =>
        _layouts.TryGetValue(semanticLayout, out string? renderer)
            ? renderer
            : "bullets";

    public string Part(string name) =>
        _parts.TryGetValue(name, out string? xml)
            ? xml
            : throw new InvalidDataException($"template package lacks required part '{name}'");

    public static PresentationTemplatePackage EmbeddedDefault(
        string version = "1.0.0",
        IEnumerable<string>? layouts = null)
    {
        string[] semanticLayouts = layouts?.ToArray() ?? CanonicalRenderers;
        var layoutMap = semanticLayouts.ToDictionary(
            layout => layout,
            layout => Renderers.Contains(layout) ? layout : "bullets",
            StringComparer.Ordinal);
        var parts = PartNames.ToDictionary(
            name => name,
            name => ReadEmbeddedPart(name),
            StringComparer.Ordinal);
        return new PresentationTemplatePackage(
            DefaultUri,
            version,
            "dcc-mcp",
            layoutMap,
            parts,
            PresentationTemplateStyle.Default,
            ReadEmbeddedBytes("logo-light.png"));
    }

    public static PresentationTemplatePackage LoadDirectory(string directory)
    {
        string root = CanonicalDirectory(directory);
        string packagePath = Path.Combine(root, "package.json");
        if (!File.Exists(packagePath))
        {
            throw new InvalidDataException($"template package is missing {packagePath}");
        }

        using JsonDocument document = JsonDocument.Parse(File.ReadAllText(packagePath));
        JsonElement package = document.RootElement;
        string schema = RequiredString(package, "schema_version");
        if (!string.Equals(schema, SchemaVersion, StringComparison.Ordinal))
        {
            throw new InvalidDataException(
                $"template package schema '{schema}' is unsupported; expected '{SchemaVersion}'");
        }
        string uri = RequiredString(package, "uri");
        if (!Regex.IsMatch(
            uri,
            "^brand://[^/]+/.+$",
            RegexOptions.CultureInvariant))
        {
            throw new InvalidDataException("template package uri must start with brand://");
        }
        string version = RequiredString(package, "version");
        if (!System.Version.TryParse(version, out System.Version? parsedVersion)
            || parsedVersion.Build < 0
            || parsedVersion.Revision >= 0)
        {
            throw new InvalidDataException(
                $"template package version '{version}' must be major.minor.patch");
        }
        string kind = RequiredString(package, "kind");
        if (!string.Equals(kind, "presentation", StringComparison.Ordinal))
        {
            throw new InvalidDataException(
                $"template package kind '{kind}' is unsupported; expected 'presentation'");
        }
        string baseUri = RequiredString(package, "extends");
        if (!string.Equals(baseUri, DefaultUri, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidDataException(
                $"template package extends '{baseUri}'; only '{DefaultUri}' is supported");
        }

        var inherited = EmbeddedDefault();
        var layouts = ReadLayouts(package);
        var parts = PartNames.ToDictionary(
            name => name,
            inherited.Part,
            StringComparer.Ordinal);
        if (package.TryGetProperty("parts", out JsonElement partMap))
        {
            if (partMap.ValueKind != JsonValueKind.Object)
            {
                throw new InvalidDataException("template package parts must be an object");
            }
            foreach (JsonProperty part in partMap.EnumerateObject())
            {
                if (!ExpectedRoots.TryGetValue(part.Name, out string? expectedRoot))
                {
                    throw new InvalidDataException(
                        $"template package part '{part.Name}' is not supported");
                }
                string relativePath = part.Value.GetString()
                    ?? throw new InvalidDataException(
                        $"template package part '{part.Name}' must be a path string");
                string xml = ReadContainedText(root, relativePath);
                ValidateXmlPart(part.Name, expectedRoot, xml);
                parts[part.Name] = xml;
            }
        }

        PresentationTemplateStyle style = ReadStyle(package, inherited.Style);
        string brandName = OptionalString(package, "brand_name", inherited.BrandName);
        byte[]? logo = inherited.Logo;
        if (package.TryGetProperty("media", out JsonElement media)
            && media.ValueKind == JsonValueKind.Object
            && media.TryGetProperty("logo", out JsonElement logoPath))
        {
            logo = ReadContainedBytes(root, logoPath.GetString()
                ?? throw new InvalidDataException("template package media.logo must be a path string"));
        }

        return new PresentationTemplatePackage(
            uri,
            version,
            brandName,
            layouts,
            parts,
            style,
            logo);
    }

    private static IReadOnlyDictionary<string, string> ReadLayouts(JsonElement package)
    {
        if (!package.TryGetProperty("layouts", out JsonElement layouts)
            || layouts.ValueKind != JsonValueKind.Object)
        {
            throw new InvalidDataException("template package layouts must be a non-empty object");
        }
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (JsonProperty layout in layouts.EnumerateObject())
        {
            string renderer = layout.Value.GetString()
                ?? throw new InvalidDataException(
                    $"template layout '{layout.Name}' renderer must be a string");
            if (layout.Name.Length == 0 || !Renderers.Contains(renderer))
            {
                throw new InvalidDataException(
                    $"template layout '{layout.Name}' maps to unsupported renderer '{renderer}'");
            }
            result.Add(layout.Name, renderer);
        }
        if (result.Count == 0)
        {
            throw new InvalidDataException("template package layouts must not be empty");
        }
        return result;
    }

    private static PresentationTemplateStyle ReadStyle(
        JsonElement package,
        PresentationTemplateStyle defaults)
    {
        if (!package.TryGetProperty("style", out JsonElement style))
        {
            return defaults;
        }
        if (style.ValueKind != JsonValueKind.Object)
        {
            throw new InvalidDataException("template package style must be an object");
        }
        JsonElement colors = style.TryGetProperty("colors", out JsonElement colorMap)
            ? colorMap
            : default;
        JsonElement fonts = style.TryGetProperty("fonts", out JsonElement fontMap)
            ? fontMap
            : default;
        return new PresentationTemplateStyle(
            Color(colors, "background", defaults.Background),
            Color(colors, "background_soft", defaults.BackgroundSoft),
            Color(colors, "panel", defaults.Panel),
            Color(colors, "accent", defaults.Accent),
            Color(colors, "accent_secondary", defaults.AccentSecondary),
            Color(colors, "text", defaults.Text),
            Color(colors, "muted", defaults.Muted),
            Color(colors, "ghost", defaults.Ghost),
            OptionalString(fonts, "latin", defaults.LatinFont),
            OptionalString(fonts, "display_latin", defaults.LatinDisplayFont),
            OptionalString(fonts, "east_asian", defaults.EastAsianFont));
    }

    private static string Color(JsonElement colors, string name, string fallback)
    {
        string value = OptionalString(colors, name, fallback).ToUpperInvariant();
        if (!Regex.IsMatch(value, "^[0-9A-F]{6}$", RegexOptions.CultureInvariant))
        {
            throw new InvalidDataException(
                $"template style color '{name}' must be a six-digit RGB hex value");
        }
        return value;
    }

    private static string OptionalString(JsonElement parent, string name, string fallback)
    {
        if (parent.ValueKind == JsonValueKind.Undefined
            || !parent.TryGetProperty(name, out JsonElement value))
        {
            return fallback;
        }
        string result = value.GetString()
            ?? throw new InvalidDataException($"template property '{name}' must be a string");
        return result.Length > 0
            ? result
            : throw new InvalidDataException($"template property '{name}' must not be empty");
    }

    private static string RequiredString(JsonElement parent, string name)
    {
        if (!parent.TryGetProperty(name, out JsonElement value)
            || value.ValueKind != JsonValueKind.String
            || string.IsNullOrWhiteSpace(value.GetString()))
        {
            throw new InvalidDataException(
                $"template package property '{name}' is required and must be a string");
        }
        return value.GetString()!;
    }

    private static void ValidateXmlPart(string part, string expectedRoot, string xml)
    {
        try
        {
            XDocument document = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
            if (document.Root?.Name.LocalName != expectedRoot)
            {
                throw new InvalidDataException(
                    $"template part '{part}' root must be '{expectedRoot}'");
            }
        }
        catch (InvalidDataException)
        {
            throw;
        }
        catch (Exception error) when (error is System.Xml.XmlException)
        {
            throw new InvalidDataException($"template part '{part}' is not valid XML", error);
        }
    }

    private static string ReadContainedText(string root, string relativePath) =>
        File.ReadAllText(ContainedPath(root, relativePath));

    private static byte[] ReadContainedBytes(string root, string relativePath) =>
        File.ReadAllBytes(ContainedPath(root, relativePath));

    private static string ContainedPath(string root, string relativePath)
    {
        if (Path.IsPathRooted(relativePath))
        {
            throw new InvalidDataException("template package paths must be relative");
        }
        string candidate = Path.GetFullPath(relativePath, root);
        if (!IsWithin(root, candidate) || !File.Exists(candidate))
        {
            throw new InvalidDataException(
                $"template package path '{relativePath}' is missing or outside its package");
        }
        string resolved = new FileInfo(candidate).ResolveLinkTarget(returnFinalTarget: true)?.FullName
            ?? candidate;
        if (!IsWithin(root, Path.GetFullPath(resolved)))
        {
            throw new InvalidDataException(
                $"template package path '{relativePath}' resolves outside its package");
        }
        return candidate;
    }

    private static string CanonicalDirectory(string directory)
    {
        string full = Path.GetFullPath(directory);
        if (!Directory.Exists(full))
        {
            throw new DirectoryNotFoundException($"template package directory not found: {full}");
        }
        return new DirectoryInfo(full).ResolveLinkTarget(returnFinalTarget: true)?.FullName
            ?? full;
    }

    private static bool IsWithin(string root, string candidate)
    {
        string prefix = root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            + Path.DirectorySeparatorChar;
        return candidate.StartsWith(prefix, StringComparison.OrdinalIgnoreCase);
    }

    private static string ReadEmbeddedPart(string name)
    {
        string resourceName = name switch
        {
            "slide_master" => "slideMaster.xml",
            "slide_layout" => "slideLayout.xml",
            "theme" => "theme.xml",
            "slide" => "slide.xml",
            "notes_master" => "notesMaster.xml",
            "notes_slide" => "notesSlide.xml",
            _ => throw new InvalidDataException($"unknown embedded template part '{name}'"),
        };
        using Stream stream = typeof(PresentationTemplatePackage).Assembly
            .GetManifestResourceStream(
                $"Office.Automation.OpenXml.Templates.{resourceName}")
            ?? throw new InvalidOperationException(
                $"missing embedded template part '{resourceName}'");
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    private static byte[] ReadEmbeddedBytes(string name)
    {
        using Stream stream = typeof(PresentationTemplatePackage).Assembly
            .GetManifestResourceStream($"Office.Automation.OpenXml.Templates.{name}")
            ?? throw new InvalidOperationException($"missing embedded template asset '{name}'");
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }
}

public sealed record PresentationTemplateStyle(
    string Background,
    string BackgroundSoft,
    string Panel,
    string Accent,
    string AccentSecondary,
    string Text,
    string Muted,
    string Ghost,
    string LatinFont,
    string LatinDisplayFont,
    string EastAsianFont)
{
    public static PresentationTemplateStyle Default { get; } = new(
        "0F141E",
        "171F2E",
        "1E2A3D",
        "4D9DE0",
        "7FC8A9",
        "E8ECF2",
        "9AA7BC",
        "223047",
        "Aptos",
        "Aptos Display",
        "Microsoft YaHei");
}
