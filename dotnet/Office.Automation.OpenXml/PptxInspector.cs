using System.IO.Packaging;
using System.Xml.Linq;

namespace Office.Automation.OpenXml;

/// <summary>
/// Self-implemented PPTX inspector — zero NuGet dependencies.
///
/// Read path for inventory/analysis with /slide[i]/shape[j] path
/// addressing (OfficeCLI research). Used by the host's inspect command and
/// as the future backing for the Python-side readers.
/// </summary>
public static class PptxInspector
{
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    public sealed record SlideInfo(
        int Index,
        int ShapeCount,
        int Pictures,
        int PicturesWithoutAlt,
        bool HasNotes,
        double WidthInches,
        double HeightInches);

    public sealed record DeckInfo(
        int SlideCount,
        IReadOnlyList<SlideInfo> Slides,
        string? Title)
    {
        public override string ToString() => $"slides={SlideCount}";
    }

    public static DeckInfo Inspect(string pptxPath)
    {
        using var package = Package.Open(pptxPath, FileMode.Open, FileAccess.Read);
        var presentationPart = package.GetPart(new Uri("/ppt/presentation.xml", UriKind.Relative));
        var presentation = XDocument.Load(presentationPart.GetStream()).Root!;
        var size = presentation.Element(P + "sldSz")!;
        double widthIn = EmuToInches(long.Parse(size.Attribute("cx")!.Value));
        double heightIn = EmuToInches(long.Parse(size.Attribute("cy")!.Value));

        var slides = new List<SlideInfo>();
        var slideUris = package.GetParts()
            .Where(part => part.Uri.ToString().StartsWith("/ppt/slides/slide", StringComparison.Ordinal)
                           && part.Uri.ToString().EndsWith(".xml", StringComparison.Ordinal))
            .OrderBy(part => part.Uri.ToString(), StringComparer.Ordinal)
            .ToList();
        foreach (var slidePart in slideUris)
        {
            var slide = XDocument.Load(slidePart.GetStream()).Root!;
            string fileName = slidePart.Uri.ToString().Split('/').Last();
            int index = int.Parse(fileName.Replace("slide", "", StringComparison.Ordinal).Replace(".xml", "", StringComparison.Ordinal));
            var shapes = slide.Descendants(P + "sp").ToList();
            int pictures = slide.Descendants(P + "pic").Count();
            int withoutAlt = slide.Descendants(P + "pic").Count(pic =>
            {
                var cNvPr = pic.Element(P + "nvPicPr")?.Element(P + "cNvPr");
                string? descr = cNvPr?.Attribute("descr")?.Value;
                string? name = cNvPr?.Attribute("name")?.Value;
                bool autoFilled = !string.IsNullOrEmpty(descr) && (
                    descr!.EndsWith(".png", StringComparison.OrdinalIgnoreCase) ||
                    descr.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) ||
                    descr.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase) ||
                    descr.EndsWith(".webp", StringComparison.OrdinalIgnoreCase) ||
                    descr.EndsWith(".svg", StringComparison.OrdinalIgnoreCase) ||
                    descr.Equals(name, StringComparison.Ordinal));
                return string.IsNullOrEmpty(descr) || autoFilled;
            });
            bool hasNotes = package.PartExists(new Uri($"/ppt/notesSlides/notesSlide{index}.xml", UriKind.Relative));
            slides.Add(new SlideInfo(index, shapes.Count, pictures, withoutAlt, hasNotes, widthIn, heightIn));
        }

        string? title = null;
        var corePart = package.GetParts().FirstOrDefault(part => part.Uri.ToString().Equals("/docProps/core.xml", StringComparison.Ordinal));
        if (corePart is not null)
        {
            var core = XDocument.Load(corePart.GetStream()).Root!;
            title = core.Descendants().FirstOrDefault(e => e.Name.LocalName == "title")?.Value;
        }
        return new DeckInfo(slides.Count, slides, title);
    }

    private static double EmuToInches(long emu) => emu / 914400.0;
}
