using System.Text.Json;
using Office.Automation.OpenXml;
using Office.Automation.Runtime;

namespace Office.Automation.Host;

/// <summary>
/// dcc-office-host — per-application Office sidecar entry point.
///
/// M1: self-implemented Open XML surface (zero NuGet dependencies, see the
/// dependency policy). The host speaks office-rpc JSON-RPC over stdin/stdout
/// (named pipe per proposal §12 comes later); COM attachment stays on the
/// roadmap behind the same envelope.
///
/// Commands (office.command.execute):
///   capability=deck.compile     params.input {ir, output}  → Deck IR JSON → PPTX
///   capability=document.inspect params.input {path}        → deck info + per-slide shapes
/// </summary>
public static class Program
{
    private static readonly string[] SupportedApps =
    {
        "powerpoint", "word", "excel", "outlook-classic", "visio", "project", "access",
    };

    public static int Main(string[] args)
    {
        string? app = null;
        bool selfTest = false;
        foreach (var arg in args)
        {
            if (arg.StartsWith("--app=", StringComparison.Ordinal))
            {
                app = arg["--app=".Length..];
            }
            else if (arg == "--self-test")
            {
                selfTest = true;
            }
        }

        if (app is null || !SupportedApps.Contains(app))
        {
            Console.Error.WriteLine(
                "usage: dcc-office-host --app=<powerpoint|word|excel|outlook-classic|visio|project|access> [--self-test]");
            return 2;
        }

        if (selfTest)
        {
            return SelfTest(app);
        }

        // M0 STA machinery still proves the runtime queue starts cleanly.
        using (var sta = new StaDispatcher())
        {
            var threadId = sta.Post(() => Environment.CurrentManagedThreadId);
            Console.Error.WriteLine($"sta dispatch ok: submitted on {Environment.CurrentManagedThreadId}, ran on {threadId}");
        }

        return JsonRpcLoop(app);
    }

    private static int JsonRpcLoop(string app)
    {
        using var stdin = Console.OpenStandardInput();
        using var reader = new StreamReader(stdin, Console.InputEncoding);
        string? line;
        while ((line = reader.ReadLine()) is not null)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }
            string response;
            try
            {
                response = Dispatch(line, app);
            }
            catch (Exception exc)
            {
                response = JsonSerializer.Serialize(new
                {
                    jsonrpc = "2.0",
                    id = (string?)null,
                    error = new { code = -32603, message = exc.Message },
                });
            }
            Console.Out.WriteLine(response);
            Console.Out.Flush();
        }
        return 0;
    }

    private static string Dispatch(string line, string app)
    {
        using var request = JsonDocument.Parse(line);
        var root = request.RootElement;
        string? id = null;
        if (root.TryGetProperty("id", out var idElement) && idElement.ValueKind == JsonValueKind.String)
        {
            id = idElement.GetString();
        }
        var method = root.GetProperty("method").GetString() ?? "";
        var result = Execute(method, root.TryGetProperty("params", out var p) ? p : default);
        return JsonSerializer.Serialize(new { jsonrpc = "2.0", id, result });
    }

    private static object Execute(string method, JsonElement parameters)
    {
        return method switch
        {
            "office.command.execute" => ExecuteCommand(parameters),
            "office.host.ping" => new { app = "powerpoint", protocol_version = "office-rpc/1" },
            _ => throw new InvalidOperationException($"unknown method: {method}"),
        };
    }

    private static object ExecuteCommand(JsonElement parameters)
    {
        var capability = parameters.GetProperty("capability").GetString() ?? "";
        var input = parameters.GetProperty("input");
        return capability switch
        {
            "deck.compile" => Compile(input),
            "document.inspect" => Inspect(input),
            _ => throw new InvalidOperationException($"OFFICE_CAPABILITY_UNSUPPORTED: {capability}"),
        };
    }

    private static object Compile(JsonElement input)
    {
        string ir = input.GetProperty("ir").GetString() ?? throw new InvalidOperationException("input.ir required");
        string output = input.GetProperty("output").GetString() ?? throw new InvalidOperationException("input.output required");
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(output))!);
        PptxWriter.CompileDeck(ir, output);
        var info = PptxInspector.Inspect(output);
        return new
        {
            operation_id = Guid.NewGuid().ToString("N"),
            backend = "openxml",
            output,
            slides = info.SlideCount,
        };
    }

    private static object Inspect(JsonElement input)
    {
        string path = input.GetProperty("path").GetString() ?? throw new InvalidOperationException("input.path required");
        var info = PptxInspector.Inspect(path);
        return new
        {
            backend = "openxml",
            slide_count = info.SlideCount,
            title = info.Title,
            slides = info.Slides.Select(s => new
            {
                index = s.Index,
                shapes = s.ShapeCount,
                pictures = s.Pictures,
                pictures_without_alt = s.PicturesWithoutAlt,
                has_notes = s.HasNotes,
            }).ToArray(),
        };
    }

    private static int SelfTest(string app)
    {
        string tempDir = Path.Combine(Path.GetTempPath(), "dcc-office-host-self-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        try
        {
            string irPath = Path.Combine(tempDir, "sample-deck.json");
            File.WriteAllText(irPath, SampleDeckIr);
            string pptxPath = Path.Combine(tempDir, "sample-deck.pptx");
            PptxWriter.CompileDeck(irPath, pptxPath);
            var info = PptxInspector.Inspect(pptxPath);
            bool ok = info.SlideCount == 3
                      && info.Slides.All(s => s.ShapeCount > 0)
                      && info.Slides.All(s => s.HasNotes);
            Console.Out.WriteLine(JsonSerializer.Serialize(new { app, ok, slides = info.SlideCount }));
            return ok ? 0 : 1;
        }
        finally
        {
            Directory.Delete(tempDir, recursive: true);
        }
    }

    private const string SampleDeckIr = """
        {
          "schema_version": "office-ir/1.0",
          "kind": "presentation",
          "document_id": "draft:self-test",
          "metadata": {"title": "Self Test", "language": "zh-CN"},
          "document": {
            "slides": [
              {"semantic_layout": "title_cover", "title": "Self Test Deck",
               "content_blocks": [{"type": "text", "paragraphs": ["host self test"]}],
               "speaker_notes": "cover"},
              {"semantic_layout": "bullets", "title": "Checks",
               "content_blocks": [{"type": "bullets", "items": ["compile ok", "inspect ok", "notes ok"]}],
               "speaker_notes": "bullets"},
              {"semantic_layout": "bullets", "title": "Close",
               "content_blocks": [{"type": "bullets", "items": ["done"]}],
               "speaker_notes": "close"}
            ]
          },
          "outputs": ["pptx"]
        }
        """;
}
