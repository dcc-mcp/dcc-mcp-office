using System.Text.Json.Nodes;

namespace Office.Automation.Com;

// Result DTOs crossing the COM backend boundary. Everything here is plain
// data: no COM reference ever leaves the STA thread (proposal §9.3).

public sealed class OfficeAppInfo
{
    public string Name { get; init; } = "";
    public string Version { get; init; } = "";
    public string Bitness { get; init; } = "";
    public string Language { get; init; } = "";
}

/// <summary>Security settings read back from the live Office Application.</summary>
public sealed class OfficeSecurityPosture
{
    public int AutomationSecurity { get; init; }
    public bool DisplayAlertsDisabled { get; init; }
    public bool? ExternalLinksAutoUpdateDisabled { get; init; }
}

/// <summary>Per-file outcome of a batch convert (proposal §15.1).</summary>
public sealed class FileConvertOutcome
{
    public string InputPath { get; init; } = "";
    public string? OutputPath { get; init; }
    public bool Ok { get; init; }
    public string? ErrorCode { get; init; }
    public string? Error { get; init; }
    public long Bytes { get; init; }
    public int PageCount { get; init; }
    public string Backend { get; init; } = "desktop_com";
    public bool Indeterminate { get; init; }
    public List<string> Warnings { get; } = new();
}

/// <summary>Per-file inspection summary; per-app shape under Summary.</summary>
public sealed class InspectOutcome
{
    public string Path { get; init; } = "";
    public string Backend { get; init; } = "desktop_com";
    /// <summary>presentation | document | workbook.</summary>
    public string Kind { get; init; } = "";
    public JsonObject Summary { get; init; } = new();
}

/// <summary>One replacement rule resolved against one file (proposal §15.2).</summary>
public sealed class RuleOutcome
{
    public string Find { get; init; } = "";
    public string Replace { get; init; } = "";
    public int Matched { get; set; }
    public int Replaced { get; set; }
}

/// <summary>Per-file outcome of batch.replace_text (proposal §15.2).</summary>
public sealed class ReplaceOutcome
{
    public string Path { get; init; } = "";
    public string Backend { get; init; } = "desktop_com";
    public bool DryRun { get; init; }
    public bool Indeterminate { get; set; }
    public List<RuleOutcome> Rules { get; } = new();
    public List<string> ScopeCovered { get; } = new();
    public List<string> Warnings { get; } = new();
    public int TotalMatched => Rules.Sum(r => r.Matched);
    public int TotalReplaced => Rules.Sum(r => r.Replaced);
}

/// <summary>A shape whose bounds exceed the slide (proposal §27 criterion 6).</summary>
public sealed class OverflowShape
{
    public string Name { get; init; } = "";
    public int ShapeId { get; init; }
    public string Kind { get; init; } = "";
    public double Left { get; init; }
    public double Top { get; init; }
    public double Right { get; init; }
    public double Bottom { get; init; }
}

/// <summary>One rendered slide preview plus its overflow report.</summary>
public sealed class SlidePreviewOutcome
{
    public int SlideIndex { get; init; }
    public string? Path { get; init; }
    public int Width { get; init; }
    public int Height { get; init; }
    public bool Ok { get; init; }
    public string? Error { get; init; }
    public List<OverflowShape> Overflow { get; init; } = new();
}

/// <summary>Replace rule input (proposal §15.2).</summary>
public sealed class ReplaceRuleInput
{
    public string Find { get; set; } = "";
    public string Replace { get; set; } = "";
    /// <summary>literal | case_insensitive. Anything else falls back to literal with a warning.</summary>
    public string Match { get; set; } = "literal";

    public bool CaseInsensitive => string.Equals(Match, "case_insensitive", StringComparison.OrdinalIgnoreCase);
}
