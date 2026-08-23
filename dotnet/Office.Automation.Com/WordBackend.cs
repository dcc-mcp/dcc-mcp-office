using System.Collections;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Office.Automation.Runtime;

namespace Office.Automation.Com;

/// <summary>
/// Word COM backend (proposal §6.2 / Phase 2 preview): batch PDF export,
/// replace-text across body/headers/footers with dry-run, COM inspection.
///
/// Office interop notes (late-bound IDispatch, zero NuGet dependencies):
///   - wdFormatPDF = 17, wdAlertsNone = 0,
///   - wdStatisticWords = 0, wdStatisticPages = 2, wdStatisticParagraphs = 4,
///   - Find.Execute(FindText, MatchCase, MatchWholeWord, MatchWildcards,
///     MatchSoundsLike, MatchAllWordForms, Forward, Wrap, Format, ReplaceWith,
///     Replace) with wdFindContinue = 1, wdReplaceNone = 0, wdReplaceAll = 2,
///   - wdStoryType: main text 1, even header 6, primary header 7, even
///     footer 8, primary footer 9, first-page header 10, first-page footer 11.
/// </summary>
public sealed class WordBackend : OfficeComBackend
{
    private const int WdFormatPdf = 17;
    private const int WdAlertsNone = 0;
    private const int WdFindStop = 0;
    private const int WdFindContinue = 1;
    private const int WdReplaceAll = 2;

    // Story types per scope (first story of each type; §15.2 note: linked
    // stories of later sections come with the NextStoryRange ladder later).
    private static readonly int[] HeaderStories = { 6, 7, 10 };
    private static readonly int[] FooterStories = { 8, 9, 11 };

    public WordBackend(StaDispatcher sta)
        : base(OfficeAppKind.Word, sta)
    {
    }

    protected override string DocumentKind => "document";

    protected override OfficeSecurityPosture ApplySecurityDefaults(dynamic app)
    {
        int displayAlerts = VerifySecuritySetting(
            "Word.DisplayAlerts",
            () => { app.DisplayAlerts = WdAlertsNone; },
            () => Convert.ToInt32(app.DisplayAlerts),
            WdAlertsNone,
            OfficeErrorCode.OfficeBackendUnavailable);
        int automationSecurity = VerifySecuritySetting(
            "Word.AutomationSecurity",
            () => { app.AutomationSecurity = 3; },
            () => Convert.ToInt32(app.AutomationSecurity),
            3,
            OfficeErrorCode.OfficeMacroBlocked);
        bool updateLinks = VerifySecuritySetting(
            "Word.Options.UpdateLinksAtOpen",
            () => { app.Options.UpdateLinksAtOpen = false; },
            () => Convert.ToBoolean(app.Options.UpdateLinksAtOpen),
            false,
            OfficeErrorCode.OfficeExternalLinkBlocked);
        return new OfficeSecurityPosture
        {
            AutomationSecurity = automationSecurity,
            DisplayAlertsDisabled = displayAlerts == WdAlertsNone,
            ExternalLinksAutoUpdateDisabled = !updateLinks,
        };
    }

    protected override ComLease OpenReadOnly(string path) => OpenDocument(path, readOnly: true);

    protected override ComLease OpenEditable(string path) => OpenDocument(path, readOnly: false);

    private ComLease OpenDocument(string path, bool readOnly)
    {
        if (!File.Exists(path))
        {
            throw DocumentNotFound(path);
        }
        dynamic doc = Application.Documents.Open(path, false, readOnly, false);
        return new ComLease(doc);
    }

    protected override void SaveEditable(dynamic document) => document.Save();

    protected override void ExportPdf(dynamic document, string outputPath) =>
        document.SaveAs2(outputPath, WdFormatPdf);

    protected override void CloseQuietly(dynamic document)
    {
        try { document.Close(0); } catch { } // wdDoNotSaveChanges
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
            dynamic d = doc.Target!;
            string title = "";
            try { title = (string)d.BuiltInDocumentProperties["Title"].Value; } catch { }
            bool hasHeader = false;
            bool hasFooter = false;
            try
            {
                hasHeader = d.Sections[1].Headers[1].Exists == -1
                    && !string.IsNullOrWhiteSpace((string)d.Sections[1].Headers[1].Range.Text);
                hasFooter = d.Sections[1].Footers[1].Exists == -1
                    && !string.IsNullOrWhiteSpace((string)d.Sections[1].Footers[1].Range.Text);
            }
            catch { }
            var summary = new JsonObject
            {
                ["title"] = title,
                ["paragraphs"] = SafeNumber(d.Paragraphs.Count),
                ["words"] = SafeNumber(d.ComputeStatistics(0)),
                ["pages"] = SafeNumber(d.ComputeStatistics(2)),
                ["tables"] = SafeNumber(d.Tables.Count),
                ["sections"] = SafeNumber(d.Sections.Count),
                ["has_header"] = hasHeader,
                ["has_footer"] = hasFooter,
            };
            // Close every request: an open document would block Quit on a
            // save prompt (the app runs hidden).
            CloseQuietly(d);
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

        bool wantBody = scope.Contains("body", StringComparer.OrdinalIgnoreCase) || scope.Count == 0;
        bool wantHeaders = scope.Contains("headers", StringComparer.OrdinalIgnoreCase);
        bool wantFooters = scope.Contains("footers", StringComparer.OrdinalIgnoreCase);
        if (wantBody)
        {
            outcome.ScopeCovered.Add("body");
        }
        if (wantHeaders)
        {
            outcome.ScopeCovered.Add("headers");
        }
        if (wantFooters)
        {
            outcome.ScopeCovered.Add("footers");
        }
        foreach (var unsupported in new[] { "comments", "charts" })
        {
            if (scope.Contains(unsupported, StringComparer.OrdinalIgnoreCase))
            {
                outcome.Warnings.Add($"scope '{unsupported}' is not yet supported by the Word COM backend");
            }
        }

        var counters = rules.ToDictionary(r => r.Find, _ => new MatchCounters(), StringComparer.Ordinal);
        bool apply = !dryRun;
        try
        {
            RunRequest($"replace_text {path}", () =>
            {
                using var doc = apply ? OpenEditable(path) : OpenReadOnly(path);
                dynamic d = doc.Target!;
                // StoryRanges yields the FIRST story of each type; the
                // NextStoryRange ladder walks the linked stories of every
                // further section so multi-section headers/footers are fully
                // covered (proposal §15.2 scope "headers"/"footers").
                foreach (dynamic firstStory in (IEnumerable)d.StoryRanges)
                {
                    dynamic? story = firstStory;
                    while (story is not null)
                    {
                        int storyType = (int)SafeNumber(story.StoryType);
                        bool inScope = (storyType == 1 && wantBody)
                            || (HeaderStories.Contains(storyType) && wantHeaders)
                            || (FooterStories.Contains(storyType) && wantFooters);
                        if (inScope)
                        {
                            foreach (var rule in rules)
                            {
                                int matched = CountInStory(story, rule);
                                counters[rule.Find].Matched += matched;
                                if (apply && matched > 0)
                                {
                                    bool replaced = ReplaceInStory(story, rule);
                                    if (replaced)
                                    {
                                        counters[rule.Find].Replaced += matched;
                                    }
                                    else
                                    {
                                        outcome.Warnings.Add($"'{rule.Find}' matched {matched} but Find/Replace did not report success");
                                    }
                                }
                            }
                        }
                        try
                        {
                            story = story.NextStoryRange;
                        }
                        catch
                        {
                            story = null;
                        }
                    }
                }
                if (apply)
                {
                    SaveEditable(d);
                }
                CloseQuietly(d);
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

    private static int CountInStory(dynamic story, ReplaceRuleInput rule)
    {
        // Count on a DUPLICATE range: after each hit Word redefines the range
        // to the found text, so counting must not consume the story range the
        // commit step needs. Wrap = wdFindStop keeps the search from wrapping
        // around and re-matching the first hit forever (Word quirk: Execute()
        // with a FindText argument also restarts from the top each call, so
        // the loop relies on property-driven, no-argument Execute calls).
        dynamic range = story.Duplicate;
        dynamic find = range.Find;
        find.ClearFormatting();
        find.Replacement.ClearFormatting();
        find.Text = rule.Find;
        find.MatchCase = !rule.CaseInsensitive;
        find.MatchWholeWord = false;
        find.MatchWildcards = false;
        find.Forward = true;
        find.Wrap = WdFindStop;
        int count = 0;
        while ((bool)find.Execute())
        {
            count++;
        }
        return count;
    }

    private static bool ReplaceInStory(dynamic story, ReplaceRuleInput rule)
    {
        dynamic range = story.Duplicate;
        dynamic find = range.Find;
        find.ClearFormatting();
        find.Replacement.ClearFormatting();
        return (bool)find.Execute(
            rule.Find, !rule.CaseInsensitive, false, false, false, false,
            true, WdFindContinue, false, rule.Replace, WdReplaceAll);
    }

    private sealed class MatchCounters
    {
        public int Matched;
        public int Replaced;
    }
}
