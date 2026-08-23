using System.Collections;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Office.Automation.Runtime;

namespace Office.Automation.Com;

/// <summary>
/// Excel COM backend (proposal §6.2 / Phase 2 preview): batch PDF export,
/// replace-text across cell values with dry-run, COM inspection.
///
/// Office interop notes (late-bound IDispatch, zero NuGet dependencies):
///   - xlTypePDF = 0, xlUpdateLinksNever = 0,
///   - Workbooks.Open(Filename, UpdateLinks, ReadOnly),
///   - ExportAsFixedFormat(Type, Filename),
///   - UsedRange.Value2 returns object[,] for ranges, a scalar for one cell.
/// </summary>
public sealed class ExcelBackend : OfficeComBackend
{
    /// <summary>Safety cap: one sheet scan is bounded (proposal §24 stress).</summary>
    private const long MaxScanCells = 500_000;

    /// <summary>Safety cap: bounded per-sheet write-back.</summary>
    private const int MaxCellWrites = 2_000;

    private const int XlTypePdf = 0;
    private const int XlUpdateLinksNever = 0;

    public ExcelBackend(
        StaDispatcher sta,
        TimeSpan? requestTimeout = null,
        int timeoutStreakForRecovery = 2)
        : base(OfficeAppKind.Excel, sta, requestTimeout, timeoutStreakForRecovery)
    {
    }

    protected override string DocumentKind => "workbook";

    protected override OfficeSecurityPosture ApplySecurityDefaults(dynamic app)
    {
        bool displayAlerts = VerifySecuritySetting(
            "Excel.DisplayAlerts",
            () => { app.DisplayAlerts = false; },
            () => Convert.ToBoolean(app.DisplayAlerts),
            false,
            OfficeErrorCode.OfficeBackendUnavailable);
        bool askToUpdateLinks = VerifySecuritySetting(
            "Excel.AskToUpdateLinks",
            () => { app.AskToUpdateLinks = false; },
            () => Convert.ToBoolean(app.AskToUpdateLinks),
            false,
            OfficeErrorCode.OfficeExternalLinkBlocked);
        int automationSecurity = VerifySecuritySetting(
            "Excel.AutomationSecurity",
            () => { app.AutomationSecurity = 3; },
            () => Convert.ToInt32(app.AutomationSecurity),
            3,
            OfficeErrorCode.OfficeMacroBlocked);
        return new OfficeSecurityPosture
        {
            AutomationSecurity = automationSecurity,
            DisplayAlertsDisabled = !displayAlerts,
            ExternalLinksAutoUpdateDisabled = !askToUpdateLinks,
        };
    }

    protected override ComLease OpenReadOnly(string path) => OpenWorkbook(path, readOnly: true);

    protected override ComLease OpenEditable(string path) => OpenWorkbook(path, readOnly: false);

    private ComLease OpenWorkbook(string path, bool readOnly)
    {
        if (!File.Exists(path))
        {
            throw DocumentNotFound(path);
        }
        dynamic wb = Application.Workbooks.Open(path, XlUpdateLinksNever, readOnly);
        return new ComLease(wb);
    }

    protected override void SaveEditable(dynamic document) => document.Save();

    protected override void ExportPdf(dynamic document, string outputPath) =>
        document.ExportAsFixedFormat(XlTypePdf, outputPath);

    protected override void CloseQuietly(dynamic document)
    {
        try { document.Close(false); } catch { }
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
            }, mayWrite: true);
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
            dynamic wb = doc.Target!;
            string title = "";
            try { title = (string)wb.BuiltinDocumentProperties["Title"].Value; } catch { }
            var sheets = new JsonArray();
            int totalValueCells = 0;
            int totalCharts = 0;
            foreach (dynamic ws in (IEnumerable)wb.Worksheets)
            {
                long rows = 0;
                long cols = 0;
                long valueCells = 0;
                long charts = 0;
                string name = (string)ws.Name;
                try
                {
                    dynamic used = ws.UsedRange;
                    if (used is not null)
                    {
                        rows = (long)SafeNumber(used.Rows.Count);
                        cols = (long)SafeNumber(used.Columns.Count);
                        valueCells = (long)SafeNumber(used.Count) - (long)SafeNumber(used.SpecialCells(4).Count);
                    }
                    charts = (long)SafeNumber(ws.ChartObjects().Count);
                }
                catch
                {
                    // Empty sheets have no UsedRange; charts collection can throw.
                }
                totalValueCells += (int)valueCells;
                totalCharts += (int)charts;
                sheets.Add(new JsonObject
                {
                    ["name"] = name,
                    ["rows"] = rows,
                    ["columns"] = cols,
                    ["value_cells"] = valueCells,
                    ["charts"] = charts,
                });
            }
            var summary = new JsonObject
            {
                ["title"] = title,
                ["sheet_count"] = sheets.Count,
                ["chart_sheet_count"] = SafeNumber(wb.Charts.Count),
                ["total_value_cells"] = totalValueCells,
                ["total_charts"] = totalCharts,
                ["sheets"] = sheets,
            };
            CloseQuietly(wb);
            return summary;
        });
        return new InspectOutcome { Path = path, Kind = DocumentKind, Summary = summary };
    }

    public override ReplaceOutcome ReplaceText(
        string path, IReadOnlyList<ReplaceRuleInput> rules, IReadOnlyList<string> scope, bool dryRun)
    {
        AttachOrThrow();
        var outcome = new ReplaceOutcome { Path = path, DryRun = dryRun };
        outcome.Rules.AddRange(rules.Select(r => new RuleOutcome { Find = r.Find, Replace = r.Replace }));
        outcome.ScopeCovered.Add("body");
        foreach (var unsupported in new[] { "headers", "footers", "notes", "comments", "charts" })
        {
            if (scope.Contains(unsupported, StringComparer.OrdinalIgnoreCase))
            {
                outcome.Warnings.Add($"scope '{unsupported}' is not yet supported by the Excel COM backend");
            }
        }

        var counters = rules.ToDictionary(r => r.Find, _ => new MatchCounters(), StringComparer.Ordinal);
        bool apply = !dryRun;
        try
        {
            RunRequest($"replace_text {path}", () =>
            {
                using var doc = apply ? OpenEditable(path) : OpenReadOnly(path);
                dynamic wb = doc.Target!;
                foreach (dynamic ws in (IEnumerable)wb.Worksheets)
                {
                    ScanSheet(ws, rules, counters, apply, outcome.Warnings);
                }
                if (apply)
                {
                    SaveEditable(wb);
                }
                CloseQuietly(wb);
            }, mayWrite: apply);
        }
        catch (OfficeComException ex)
        {
            outcome.Indeterminate = ex.Indeterminate;
            outcome.Warnings.Add($"replace failed: {ex.Code.ToWireName()}: {ex.Message}");
        }

        foreach (var rule in outcome.Rules)
        {
            var counter = counters.TryGetValue(rule.Find, out var c) ? c : new MatchCounters();
            rule.Matched = counter.Matched;
            rule.Replaced = counter.Replaced;
        }
        return outcome;
    }

    private void ScanSheet(
        dynamic ws, IReadOnlyList<ReplaceRuleInput> rules,
        Dictionary<string, MatchCounters> counters, bool apply, List<string> warnings)
    {
        string sheetName;
        try { sheetName = (string)ws.Name; } catch { sheetName = "(unknown)"; }
        dynamic used = ws.UsedRange;
        if (used is null)
        {
            return;
        }
        long rows;
        long cols;
        try
        {
            rows = (long)SafeNumber(used.Rows.Count);
            cols = (long)SafeNumber(used.Columns.Count);
        }
        catch
        {
            return;
        }
        if (rows * cols > MaxScanCells)
        {
            warnings.Add($"sheet '{sheetName}' exceeds the {MaxScanCells}-cell scan cap and was skipped");
            return;
        }

        object[,] values;
        try
        {
            dynamic raw = used.Value2;
            values = raw is object[,] matrix ? matrix : new object[1, 1] { { raw } };
        }
        catch
        {
            return;
        }

        // COM SAFEARRAYs from Value2 can be one-based — index via the
        // array's own bounds and translate back to 1-based sheet coordinates.
        int lowerRow = values.GetLowerBound(0);
        int lowerCol = values.GetLowerBound(1);
        var writes = new List<(int Row, int Col, string Value)>();
        for (int r = lowerRow; r <= values.GetUpperBound(0); r++)
        {
            for (int c = lowerCol; c <= values.GetUpperBound(1); c++)
            {
                if (values[r, c] is not string text || text.Length == 0)
                {
                    continue;
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
                if (apply && updated != text)
                {
                    if (writes.Count >= MaxCellWrites)
                    {
                        warnings.Add($"sheet '{sheetName}' exceeds the {MaxCellWrites}-write cap; remaining changes skipped");
                        break;
                    }
                    writes.Add((r - lowerRow + 1, c - lowerCol + 1, updated));
                }
            }
            if (apply && writes.Count >= MaxCellWrites)
            {
                break;
            }
        }

        if (apply)
        {
            foreach (var (row, col, value) in writes)
            {
                used.Cells[row, col] = value;
            }
        }
    }

    private static Regex BuildRegex(ReplaceRuleInput rule)
    {
        var options = RegexOptions.CultureInvariant;
        if (rule.CaseInsensitive)
        {
            options |= RegexOptions.IgnoreCase;
        }
        return new Regex(Regex.Escape(rule.Find), options);
    }

    private sealed class MatchCounters
    {
        public int Matched;
        public int Replaced;
    }
}
