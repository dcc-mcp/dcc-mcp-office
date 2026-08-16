using System.Text.Json;
using Office.Automation.Com;
using Office.Automation.Host;
using Office.Automation.OpenXml;
using Office.Automation.Runtime;

/// <summary>
/// dcc-office-host — per-application Office sidecar entry point (proposal §8.2).
///
/// Modes:
///   --app=<app> --pipe       named-pipe JSON-RPC server (default) — the
///                            gateway-facing transport from proposal §12
///   --app=<app> --stdio      stdin/stdout JSON-RPC loop (local debugging)
///   --app=<app> --self-test        Open XML round-trip, no Office required
///   --app=<app> --self-test-com    Open XML + real COM probe (PDF convert,
///                                  replace-text dry-run/commit, slide previews)
///
/// Commands (office.command.execute capabilities):
///   deck.compile          {input:{ir, output}}                → Deck IR → PPTX
///   document.inspect      {input:{path, backend?}}            → structure summary
///   batch.convert         {input:{inputs, output_directory}}  → PDF per file (COM)
///   batch.replace_text    {input:{inputs, rules, scope?, dry_run?}}
///   slide.render          {input:{path, output_directory, width?, height?}}
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
        bool selfTestCom = false;
        bool stdio = false;
        string? pipeName = null;
        foreach (var arg in args)
        {
            switch (arg)
            {
                case "--self-test":
                    selfTest = true;
                    break;
                case "--self-test-com":
                    selfTestCom = true;
                    break;
                case "--stdio":
                    stdio = true;
                    break;
                default:
                    if (arg.StartsWith("--app=", StringComparison.Ordinal))
                    {
                        app = arg["--app=".Length..];
                    }
                    else if (arg.StartsWith("--pipe-name=", StringComparison.Ordinal))
                    {
                        pipeName = arg["--pipe-name=".Length..];
                    }
                    break;
            }
        }

        if (app is null || !SupportedApps.Contains(app))
        {
            Console.Error.WriteLine(
                "usage: dcc-office-host --app=<powerpoint|word|excel|outlook-classic|visio|project|access> [--pipe|--stdio] [--self-test|--self-test-com]");
            return 2;
        }

        if (selfTest)
        {
            return SelfTest(app, probeCom: false);
        }
        if (selfTestCom)
        {
            return SelfTest(app, probeCom: true);
        }

        using var router = new CommandRouter(app);
        if (stdio)
        {
            return JsonRpcLoop(router);
        }

        var server = new OfficePipeServer(
            app,
            requestLine => router.Dispatch(requestLine),
            pipeName,
            () => router.ShutdownRequested);
        Console.Error.WriteLine($"office-host[{app}] listening on {server.PipeName}");
        using var cts = new CancellationTokenSource();
        Console.CancelKeyPress += (_, eventArgs) =>
        {
            eventArgs.Cancel = true;
            cts.Cancel();
        };
        server.Run(cts.Token, () => router.ShutdownRequested);
        return 0;
    }

    private static int JsonRpcLoop(CommandRouter router)
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
            string response = router.Dispatch(line);
            Console.Out.WriteLine(response);
            Console.Out.Flush();
        }
        return 0;
    }

    // ------------------------------------------------------------- self-test

    private static int SelfTest(string app, bool probeCom)
    {
        string tempDir = Path.Combine(Path.GetTempPath(), "dcc-office-host-self-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        try
        {
            // Open XML round-trip: Deck IR → PPTX → inspect (no Office needed).
            string irPath = Path.Combine(tempDir, "sample-deck.json");
            File.WriteAllText(irPath, SampleDeckIr);
            string pptxPath = Path.Combine(tempDir, "sample-deck.pptx");
            PptxWriter.CompileDeck(irPath, pptxPath);
            var info = PptxInspector.Inspect(pptxPath);
            bool openXmlOk = info.SlideCount == 3
                             && info.Slides.All(s => s.ShapeCount > 0)
                             && info.Slides.All(s => s.HasNotes);

            bool comOk = false;
            bool comSkipped = false;
            string comDetail = "";
            if (probeCom)
            {
                try
                {
                    comOk = ComProbe(pptxPath, tempDir);
                    comDetail = "ok";
                }
                catch (OfficeComException ex) when (ex.Code == OfficeErrorCode.OfficeAppNotInstalled)
                {
                    comSkipped = true;
                    comDetail = $"skipped: {ex.Message}";
                }
                catch (OfficeComException ex)
                {
                    comDetail = $"failed: {ex.Code.ToWireName()}: {ex.Message}";
                }
            }

            Console.Out.WriteLine(JsonSerializer.Serialize(new
            {
                app,
                openxml_ok = openXmlOk,
                com = probeCom ? comDetail : "not probed",
                com_ok = comOk,
                com_skipped = comSkipped,
                slides = info.SlideCount,
            }));

            bool ok = openXmlOk && (!probeCom || comOk || comSkipped);
            return ok ? 0 : 1;
        }
        finally
        {
            try { Directory.Delete(tempDir, recursive: true); } catch { }
        }
    }

    /// <summary>
    /// Real COM probe (needs PowerPoint on this machine, user session only —
    /// proposal §8.1 never automates from Session 0): PDF export, replace-text
    /// dry-run + commit, slide preview render.
    /// </summary>
    private static bool ComProbe(string pptxPath, string tempDir)
    {
        using var sta = new StaDispatcher();
        var backend = ComBackendFactory.Create("powerpoint", sta);
        backend.Attach(TimeSpan.FromSeconds(90));

        string pdfPath = Path.Combine(tempDir, "sample-deck.pdf");
        var convert = backend.ConvertToPdf(pptxPath, pdfPath);
        if (!convert.Ok || convert.PageCount < 1)
        {
            return false;
        }

        // dry-run first: must find matches without touching the file
        var dry = backend.ReplaceText(pptxPath,
            new[] { new ReplaceRuleInput { Find = "Checks", Replace = "Verification" } },
            new[] { "body" }, dryRun: true);
        if (dry.TotalMatched < 1)
        {
            return false;
        }

        // commit: the match must be replaced and saved
        var commit = backend.ReplaceText(pptxPath,
            new[] { new ReplaceRuleInput { Find = "Checks", Replace = "Verification" } },
            new[] { "body" }, dryRun: false);
        if (commit.TotalReplaced < 1)
        {
            return false;
        }

        // slide previews: one PNG per slide, all present on disk
        string previewDir = Path.Combine(tempDir, "previews");
        var previews = backend.ExportSlidePreviews(pptxPath, previewDir, 640, 360);
        return previews is not null
               && previews.Count == 3
               && previews.All(p => p.Ok && p.Path is not null && File.Exists(p.Path));
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
