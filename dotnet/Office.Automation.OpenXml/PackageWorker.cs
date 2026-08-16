namespace Office.Automation.OpenXml;

/// <summary>
/// Batch structural worker over Open XML packages (DOCX/XLSX/PPTX).
///
/// This backend is a <b>structural compiler and batch engine — never a
/// renderer</b> (proposal §6.1). It must NOT be asked to produce final
/// pagination, final slide visuals, native formula recalc or high-fidelity
/// PDF; those go through the COM sidecar or Microsoft Graph.
///
/// M1 scope: bulk text replace (semantic, never raw XML string replace),
/// property/style/table/placeholder edits, template-based generation,
/// parallel multi-file processing on machines without Office.
/// </summary>
public static class PackageWorker
{
    /// <summary>Preflight: package openable, extension/content match,
    /// relationship integrity (proposal §18.1).</summary>
    public static void ValidatePackage(string path)
    {
        // M1: implement via DocumentFormat.OpenXml.Packaging.OpenXmlPackage.
        throw new NotImplementedException("M1: Open XML validation not wired yet.");
    }
}
