using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;

namespace Office.Automation.Com;

/// <summary>Counts page dictionaries in regular and Flate-compressed PDF object streams.</summary>
internal static class PdfPageCounter
{
    private static readonly Regex PageObjectPattern = new(
        @"/Type\s*/Page(?![A-Za-z])",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex ObjectStreamPattern = new(
        @"(?<dictionary><<(?:(?!>>).)*?/Type\s*/ObjStm(?:(?!>>).)*?>>)\s*stream(?:\r\n|\n|\r)",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.Singleline);

    private static readonly Regex LengthPattern = new(
        @"/Length\s+(?<length>\d+)",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    public static int Count(string path)
    {
        byte[] bytes = File.ReadAllBytes(path);
        string text = Encoding.Latin1.GetString(bytes);
        int pages = PageObjectPattern.Matches(text).Count;

        foreach (Match stream in ObjectStreamPattern.Matches(text))
        {
            string dictionary = stream.Groups["dictionary"].Value;
            if (!dictionary.Contains("/FlateDecode", StringComparison.Ordinal))
            {
                continue;
            }

            Match lengthMatch = LengthPattern.Match(dictionary);
            if (!lengthMatch.Success
                || !int.TryParse(lengthMatch.Groups["length"].Value, out int length)
                || length < 0
                || stream.Index + stream.Length + length > bytes.Length)
            {
                continue;
            }

            pages += CountCompressedPages(bytes, stream.Index + stream.Length, length);
        }

        return pages;
    }

    private static int CountCompressedPages(byte[] bytes, int offset, int length)
    {
        try
        {
            using var input = new MemoryStream(bytes, offset, length, writable: false);
            using var inflater = new ZLibStream(input, CompressionMode.Decompress);
            using var output = new MemoryStream();
            inflater.CopyTo(output);
            return PageObjectPattern.Matches(Encoding.Latin1.GetString(output.ToArray())).Count;
        }
        catch (InvalidDataException)
        {
            return 0;
        }
    }
}
