using System.Collections;
using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Office.Automation.Runtime;

namespace Office.Automation.Com;

/// <summary>
/// PowerPoint COM backend (proposal §6.2 / Phase 1): batch PDF export,
/// replace-text with dry-run, COM inspection, per-slide preview render plus
/// overflow detection (§27 criterion 6).
///
/// Office interop notes (late-bound IDispatch, zero NuGet dependencies):
///   - MsoTriState: -1 = msoTrue, 0 = msoFalse,
///   - shape.Type: 6 = msoGroup, 13 = msoPicture, 14 = msoPlaceholder,
///   - ppSaveAsPDF = 32,
///   - Presentations.Open(FileName, ReadOnly, Untitled, WithWindow).
/// </summary>
public sealed class PowerPointBackend : OfficeComBackend
{
    private const int MsoTrue = -1;
    private const int MsoFalse = 0;
    private const int PpSaveAsPdf = 32;

    public PowerPointBackend(StaDispatcher sta)
        : base(OfficeAppKind.PowerPoint, sta)
    {
    }

    protected override string DocumentKind => "presentation";

    protected override void ApplySecurityDefaults(dynamic app)
    {
        // §19: macros disabled while opening untrusted files; alerts off.
        try { app.AutomationSecurity = 3; } catch { } // msoAutomationSecurityForceDisable
        try { app.DisplayAlerts = 1; } catch { }      // ppAlertsNone
    }

    protected override ComLease OpenReadOnly(string path)
    {
        if (!File.Exists(path))
        {
            throw DocumentNotFound(path);
        }
        dynamic pres = Application.Presentations.Open(path, MsoTrue, MsoFalse, MsoFalse);
        return new ComLease(pres);
    }

    protected override ComLease OpenEditable(string path)
    {
        if (!File.Exists(path))
        {
            throw DocumentNotFound(path);
        }
        dynamic pres = Application.Presentations.Open(path, MsoFalse, MsoFalse, MsoFalse);
        return new ComLease(pres);
    }

    /// <summary>Previews render off a windowed-but-hidden presentation.</summary>
    protected override ComLease OpenForRender(string path)
    {
        if (!File.Exists(path))
        {
            throw DocumentNotFound(path);
        }
        dynamic pres = Application.Presentations.Open(path, MsoTrue, MsoFalse, MsoTrue);
        return new ComLease(pres);
    }

    protected override void SaveEditable(dynamic document) => document.Save();

    protected override void ExportPdf(dynamic document, string outputPath) =>
        document.SaveAs(outputPath, PpSaveAsPdf);

    protected override void CloseQuietly(dynamic document)
    {
        try { document.Close(); } catch { }
    }

    public override FileConvertOutcome ConvertToPdf(string path, string outputPath)
    {
        var attachError = AttachOrError();
        if (attachError is not null)
        {
            return ErrorOutcome(path, outputPath, attachError);
        }
        try
        {
            RunRequest($"convert {path}", () =>
            {
                using var doc = OpenReadOnly(path);
                ExportPdf(doc.Target!, outputPath);
                CloseQuietly(doc.Target!);
            });
            ValidatePdfOutput(outputPath, path);
            return new FileConvertOutcome
            {
                InputPath = path,
                OutputPath = outputPath,
                Ok = true,
                Bytes = new FileInfo(outputPath).Length,
                PageCount = CountPdfPages(outputPath),
            };
        }
        catch (OfficeComException ex)
        {
            return ErrorOutcome(path, outputPath, ex);
        }
    }

    public override InspectOutcome Inspect(string path)
    {
        AttachOrThrow();
        var summary = RunRequest($"inspect {path}", () =>
        {
            using var doc = OpenReadOnly(path);
            dynamic pres = doc.Target!;
            double slideWidth = SafeNumber(pres.PageSetup.SlideWidth);
            double slideHeight = SafeNumber(pres.PageSetup.SlideHeight);
            var slides = new JsonArray();
            int index = 0;
            foreach (dynamic slide in (IEnumerable)pres.Slides)
            {
                index++;
                int pictures = 0;
                long textLength = 0;
                string title = "";
                int shapes = slide.Shapes.Count;
                foreach (dynamic shape in (IEnumerable)slide.Shapes)
                {
                    int type = SafeNumber(shape.Type) is var t ? (int)t : 0;
                    if (type == 13)
                    {
                        pictures++;
                    }
                    try
                    {
                        if (shape.HasTextFrame == MsoTrue && shape.TextFrame.HasText == MsoTrue)
                        {
                            string text = (string)shape.TextFrame.TextRange.Text;
                            textLength += text.Length;
                            if (title.Length == 0)
                            {
                                title = FirstLine(text);
                            }
                        }
                    }
                    catch
                    {
                        // Placeholder-less shapes can throw on TextFrame probing; skip.
                    }
                }
                bool hasNotes = false;
                try { hasNotes = slide.HasNotesPage == MsoTrue; } catch { }
                slides.Add(new JsonObject
                {
                    ["index"] = index,
                    ["shapes"] = shapes,
                    ["pictures"] = pictures,
                    ["has_notes"] = hasNotes,
                    ["text_length"] = textLength,
                    ["title"] = title,
                });
            }
            return new JsonObject
            {
                ["slide_count"] = index,
                ["slide_width_pt"] = slideWidth,
                ["slide_height_pt"] = slideHeight,
                ["slides"] = slides,
            };
        });
        return new InspectOutcome { Path = path, Kind = DocumentKind, Summary = summary };
    }

    public override ReplaceOutcome ReplaceText(
        string path, IReadOnlyList<ReplaceRuleInput> rules, IReadOnlyList<string> scope, bool dryRun)
    {
        AttachOrThrow();
        var outcome = new ReplaceOutcome { Path = path, DryRun = dryRun };
        outcome.Rules.AddRange(rules.Select(r => new RuleOutcome { Find = r.Find, Replace = r.Replace }));

        bool wantBody = scope.Contains("body", StringComparer.OrdinalIgnoreCase) || scope.Count == 0;
        bool wantNotes = scope.Contains("notes", StringComparer.OrdinalIgnoreCase);
        if (scope.Contains("comments", StringComparer.OrdinalIgnoreCase)
            || scope.Contains("charts", StringComparer.OrdinalIgnoreCase))
        {
            outcome.Warnings.Add("scope 'comments'/'charts' is not yet supported by the PowerPoint COM backend");
        }
        outcome.ScopeCovered.AddRange(wantBody ? new[] { "body" } : Array.Empty<string>());
        if (wantNotes)
        {
            outcome.ScopeCovered.Add("notes");
        }

        var counters = rules.ToDictionary(
            r => r.Find, _ => new MatchCounters(), StringComparer.Ordinal);
        bool apply = !dryRun;
        try
        {
            RunRequest($"replace_text {path}", () =>
            {
                using var doc = apply ? OpenEditable(path) : OpenReadOnly(path);
                dynamic pres = doc.Target!;
                foreach (dynamic slide in (IEnumerable)pres.Slides)
                {
                    if (wantBody)
                    {
                        foreach (dynamic shape in (IEnumerable)slide.Shapes)
                        {
                            ProcessShapeText(shape, rules, counters, apply, outcome.Warnings);
                        }
                    }
                    if (wantNotes)
                    {
                        try
                        {
                            foreach (dynamic shape in (IEnumerable)slide.NotesPage.Shapes)
                            {
                                ProcessShapeText(shape, rules, counters, apply, outcome.Warnings);
                            }
                        }
                        catch
                        {
                            // Slides without notes pages simply contribute nothing.
                        }
                    }
                }
                if (apply)
                {
                    SaveEditable(pres);
                }
                CloseQuietly(pres);
            });
        }
        catch (OfficeComException ex)
        {
            // §15.2 dry-run contract: report per-file failure, never a silent drop.
            outcome.Warnings.Add($"replace failed: {ex.Code.ToWireName()}: {ex.Message}");
        }

        foreach (var rule in outcome.Rules)
        {
            var counter = counters.TryGetValue(rule.Find, out var c)
                ? c
                : new MatchCounters();
            rule.Matched = counter.Matched;
            rule.Replaced = counter.Replaced;
        }
        return outcome;
    }

    public override List<SlidePreviewOutcome> ExportSlidePreviews(
        string path, string outputDirectory, int width, int height)
    {
        AttachOrThrow();
        Directory.CreateDirectory(outputDirectory);
        return RunRequest($"render {path}", () =>
        {
            using var doc = OpenForRender(path);
            dynamic pres = doc.Target!;
            double slideWidth = SafeNumber(pres.PageSetup.SlideWidth);
            double slideHeight = SafeNumber(pres.PageSetup.SlideHeight);
            var results = new List<SlidePreviewOutcome>();
            int index = 0;
            foreach (dynamic slide in (IEnumerable)pres.Slides)
            {
                index++;
                string outFile = Path.Combine(outputDirectory, $"slide-{index:D3}.png");
                try
                {
                    slide.Export(outFile, "PNG", width, height);
                }
                catch (COMException ex)
                {
                    results.Add(new SlidePreviewOutcome
                    {
                        SlideIndex = index,
                        Width = width,
                        Height = height,
                        Ok = false,
                        Error = ex.Message,
                    });
                    continue;
                }
                var overflow = new List<OverflowShape>();
                foreach (dynamic shape in (IEnumerable)slide.Shapes)
                {
                    CollectOverflow(shape, slideWidth, slideHeight, overflow);
                }
                results.Add(new SlidePreviewOutcome
                {
                    SlideIndex = index,
                    Path = outFile,
                    Width = width,
                    Height = height,
                    Ok = true,
                    Overflow = overflow,
                });
            }
            return results;
        });
    }

    // ------------------------------------------------------------- internals

    private void ProcessShapeText(
        dynamic shape, IReadOnlyList<ReplaceRuleInput> rules,
        Dictionary<string, MatchCounters> counters, bool apply, List<string> warnings)
    {
        try
        {
            int type = (int)SafeNumber(shape.Type);
            if (type == 6) // msoGroup
            {
                foreach (dynamic child in (IEnumerable)shape.GroupItems)
                {
                    ProcessShapeText(child, rules, counters, apply, warnings);
                }
                return;
            }
            if (shape.HasTable == MsoTrue)
            {
                dynamic table = shape.Table;
                for (int r = 1; r <= table.Rows.Count; r++)
                {
                    for (int c = 1; c <= table.Columns.Count; c++)
                    {
                        ReplaceInTextRange(table.Cell(r, c).Shape.TextFrame.TextRange, rules, counters, apply);
                    }
                }
                return;
            }
            if (shape.HasChart == MsoTrue)
            {
                warnings.Add($"shape '{ShapeName(shape)}' is a chart; chart text replacement is not supported");
                return;
            }
            if (shape.HasTextFrame == MsoTrue && shape.TextFrame.HasText == MsoTrue)
            {
                ReplaceInTextRange(shape.TextFrame.TextRange, rules, counters, apply);
            }
        }
        catch
        {
            // Unsupported shape kinds degrade to a warning, never a crash.
            warnings.Add($"shape skipped: {ShapeName(shape)}");
        }
    }

    private void ReplaceInTextRange(
        dynamic range, IReadOnlyList<ReplaceRuleInput> rules,
        Dictionary<string, MatchCounters> counters, bool apply)
    {
        string text = (string)range.Text;
        if (string.IsNullOrEmpty(text))
        {
            return;
        }
        string updated = text;
        foreach (var rule in rules)
        {
            var regex = BuildRegex(rule);
            int count = regex.Matches(updated).Count;
            if (count == 0)
            {
                continue;
            }
            counters[rule.Find].Matched += count;
            if (apply)
            {
                updated = regex.Replace(updated, rule.Replace);
                counters[rule.Find].Replaced += count;
            }
        }
        if (apply && !ReferenceEquals(updated, text) && updated != text)
        {
            range.Text = updated;
        }
    }

    private void CollectOverflow(dynamic shape, double slideWidth, double slideHeight, List<OverflowShape> overflow)
    {
        try
        {
            int type = (int)SafeNumber(shape.Type);
            if (type == 6)
            {
                foreach (dynamic child in (IEnumerable)shape.GroupItems)
                {
                    CollectOverflow(child, slideWidth, slideHeight, overflow);
                }
                return;
            }
            double left = SafeNumber(shape.Left);
            double top = SafeNumber(shape.Top);
            double width = SafeNumber(shape.Width);
            double height = SafeNumber(shape.Height);
            const double tolerance = 1.0; // pt
            if (left < -tolerance || top < -tolerance
                || left + width > slideWidth + tolerance
                || top + height > slideHeight + tolerance)
            {
                overflow.Add(new OverflowShape
                {
                    Name = ShapeName(shape),
                    ShapeId = (int)SafeNumber(shape.Id),
                    Kind = KindName(type),
                    Left = left,
                    Top = top,
                    Right = left + width,
                    Bottom = top + height,
                });
            }
        }
        catch
        {
            // Off-slide connectors etc. may not expose bounds; skip quietly.
        }
    }

    private static Regex BuildRegex(ReplaceRuleInput rule)
    {
        var options = RegexOptions.None;
        if (rule.CaseInsensitive)
        {
            options |= RegexOptions.IgnoreCase;
        }
        return new Regex(Regex.Escape(rule.Find), options | RegexOptions.CultureInvariant);
    }

    private static string ShapeName(dynamic shape)
    {
        try { return (string)shape.Name; }
        catch { return "(unnamed)"; }
    }

    private static string KindName(int type) => type switch
    {
        6 => "group",
        13 => "picture",
        14 => "placeholder",
        19 => "table",
        _ => "shape",
    };

    private static string FirstLine(string text)
    {
        text = text.Trim();
        int cut = text.IndexOfAny(new[] { '\r', '\n' });
        string line = cut >= 0 ? text[..cut] : text;
        return line.Length > 120 ? line[..120] + "…" : line;
    }

    private sealed class MatchCounters
    {
        public int Matched;
        public int Replaced;
    }
}
