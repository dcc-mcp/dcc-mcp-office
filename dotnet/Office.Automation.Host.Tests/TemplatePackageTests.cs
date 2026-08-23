using System.IO.Compression;
using System.Text.Json;
using Office.Automation.Host;
using Xunit;

namespace Office.Automation.Host.Tests;

public sealed class TemplatePackageTests
{
    private const string TemplateUri = "brand://test/studio";

    [Fact]
    public void ExternalPackageIsAdvertisedAndMaterializedWithoutRebuildingTheHost()
    {
        string temporary = Path.Combine(
            Path.GetTempPath(),
            $"dcc-office-template-package-{Guid.NewGuid():N}");
        string templateRoot = Path.Combine(temporary, "templates");
        string packageRoot = Path.Combine(templateRoot, "studio");
        string output = Path.Combine(temporary, "studio.pptx");
        Directory.CreateDirectory(packageRoot);
        try
        {
            WritePackage(packageRoot);
            using var router = new CommandRouter(
                "powerpoint",
                enableDesktopCom: false,
                workspaceRoot: temporary,
                templateDirectories: [templateRoot]);

            using JsonDocument handshake = JsonDocument.Parse(router.Dispatch(
                """
                {"jsonrpc":"2.0","id":1,"method":"office.host.handshake","params":{"requested_app":"powerpoint"}}
                """));
            JsonElement package = handshake.RootElement
                .GetProperty("result")
                .GetProperty("capability_manifest")
                .GetProperty("template_packages")
                .GetProperty(TemplateUri);
            Assert.Equal("1.2.3", package.GetProperty("version").GetString());
            Assert.Equal("file", package.GetProperty("source_kind").GetString());
            Assert.Contains(
                "studio_cover",
                package.GetProperty("layouts").EnumerateArray().Select(item => item.GetString()));

            string ir = """
                {
                  "schema_version":"office-ir/1.0",
                  "kind":"presentation",
                  "document_id":"draft:studio-template",
                  "metadata":{"title":"Studio Template","language":"en-US"},
                  "document":{"slides":[{
                    "semantic_layout":"studio_cover",
                    "title":"External template",
                    "content_blocks":[{"type":"text","paragraphs":["fixture"]}],
                    "speaker_notes":"fixture"
                  }]},
                  "outputs":["pptx"]
                }
                """;
            string request = JsonSerializer.Serialize(new
            {
                jsonrpc = "2.0",
                id = 2,
                method = "office.command.execute",
                @params = new
                {
                    capability = "deck.compile",
                    input = new { ir, output, template = TemplateUri },
                    policy = new { workspace_root = temporary },
                },
            });

            using JsonDocument response = JsonDocument.Parse(router.Dispatch(request));

            Assert.True(
                response.RootElement.TryGetProperty("result", out JsonElement result),
                response.RootElement.ToString());
            Assert.Empty(result.GetProperty("warnings").EnumerateArray());
            using ZipArchive archive = ZipFile.OpenRead(output);
            string theme = ReadEntry(archive, "ppt/theme/theme1.xml");
            string slide = ReadEntry(archive, "ppt/slides/slide1.xml");
            Assert.Contains("Studio Test Theme", theme, StringComparison.Ordinal);
            Assert.Contains("AA1122", slide, StringComparison.Ordinal);
            Assert.Contains("Aptos Display", slide, StringComparison.Ordinal);
            Assert.Contains("Studio Test", slide, StringComparison.Ordinal);
        }
        finally
        {
            if (Directory.Exists(temporary))
            {
                Directory.Delete(temporary, recursive: true);
            }
        }
    }

    [Fact]
    public void PackagePartTraversalIsRejectedBeforeTheTemplateIsAdvertised()
    {
        string temporary = Path.Combine(
            Path.GetTempPath(),
            $"dcc-office-template-traversal-{Guid.NewGuid():N}");
        string templateRoot = Path.Combine(temporary, "templates");
        string packageRoot = Path.Combine(templateRoot, "unsafe");
        Directory.CreateDirectory(packageRoot);
        try
        {
            File.WriteAllText(Path.Combine(temporary, "outside.xml"), "<a:theme xmlns:a=\"urn:test\"/>");
            File.WriteAllText(
                Path.Combine(packageRoot, "package.json"),
                """
                {
                  "schema_version":"office-template-package/1.0",
                  "uri":"brand://test/unsafe",
                  "version":"1.0.0",
                  "kind":"presentation",
                  "extends":"brand://dcc-mcp/default",
                  "layouts":{"bullets":"bullets"},
                  "parts":{"theme":"../../outside.xml"}
                }
                """);

            InvalidDataException error = Assert.Throws<InvalidDataException>(() =>
                new TemplateRegistry([templateRoot]));

            Assert.Contains("outside its package", error.Message, StringComparison.Ordinal);
        }
        finally
        {
            if (Directory.Exists(temporary))
            {
                Directory.Delete(temporary, recursive: true);
            }
        }
    }

    [Fact]
    public void ShippedStudioLightPackageIsMaterializedFromTheRepository()
    {
        string templates = FindTemplatesDirectory();

        var registry = new TemplateRegistry([templates]);
        TemplateEntry? package = registry.Resolve("brand://dcc-mcp/studio-light");

        Assert.NotNull(package);
        Assert.Equal("1.0.0", package.Package.Version);
        Assert.Equal("F7F9FC", package.Package.Style.Background);
        Assert.Contains("technical_architecture", package.Package.Layouts);
    }

    private static void WritePackage(string packageRoot)
    {
        File.WriteAllText(
            Path.Combine(packageRoot, "package.json"),
            """
            {
              "schema_version":"office-template-package/1.0",
              "uri":"brand://test/studio",
              "version":"1.2.3",
              "kind":"presentation",
              "extends":"brand://dcc-mcp/default",
              "brand_name":"Studio Test",
              "layouts":{"studio_cover":"title_cover","bullets":"bullets"},
              "parts":{"theme":"theme.xml"},
              "style":{
                "colors":{
                  "background":"AA1122",
                  "background_soft":"BB2233",
                  "panel":"CC3344",
                  "accent":"DD4455",
                  "accent_secondary":"EE5566",
                  "text":"F8F8F8",
                  "muted":"DDDDDD",
                  "ghost":"993344"
                },
                "fonts":{"latin":"Aptos Display","east_asian":"Microsoft YaHei"}
              }
            }
            """);
        File.WriteAllText(
            Path.Combine(packageRoot, "theme.xml"),
            """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Studio Test Theme">
              <a:themeElements>
                <a:clrScheme name="Studio Test"><a:dk1><a:srgbClr val="111111"/></a:dk1></a:clrScheme>
                <a:fontScheme name="Studio Test"><a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/></a:minorFont></a:fontScheme>
                <a:fmtScheme name="Studio Test"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
              </a:themeElements>
            </a:theme>
            """);
    }

    private static string ReadEntry(ZipArchive archive, string path)
    {
        ZipArchiveEntry entry = archive.GetEntry(path)
            ?? throw new InvalidOperationException($"missing package part {path}");
        using var reader = new StreamReader(entry.Open());
        return reader.ReadToEnd();
    }

    private static string FindTemplatesDirectory()
    {
        DirectoryInfo? directory = new(AppContext.BaseDirectory);
        while (directory is not null)
        {
            string candidate = Path.Combine(directory.FullName, "templates");
            if (File.Exists(Path.Combine(candidate, "registry.json")))
            {
                return candidate;
            }
            directory = directory.Parent;
        }
        throw new DirectoryNotFoundException("repository templates directory not found");
    }
}
