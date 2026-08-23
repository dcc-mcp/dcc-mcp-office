using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using Office.Automation.Runtime;
using Xunit;

namespace Office.Automation.Com.Tests;

public sealed class ComReliabilityTests
{
    [Theory]
    [InlineData(unchecked((int)0x80010001), OfficeErrorCode.OfficeAppBusy)]
    [InlineData(unchecked((int)0x8001010A), OfficeErrorCode.OfficeAppBusy)]
    [InlineData(unchecked((int)0x80040154), OfficeErrorCode.OfficeAppNotInstalled)]
    [InlineData(unchecked((int)0x80030020), OfficeErrorCode.OfficeDocumentLocked)]
    [InlineData(unchecked((int)0x80030021), OfficeErrorCode.OfficeDocumentLocked)]
    [InlineData(unchecked((int)0x80070020), OfficeErrorCode.OfficeDocumentLocked)]
    [InlineData(unchecked((int)0x80070021), OfficeErrorCode.OfficeDocumentLocked)]
    [InlineData(unchecked((int)0x80030002), OfficeErrorCode.OfficeDocumentNotFound)]
    [InlineData(unchecked((int)0x80030003), OfficeErrorCode.OfficeDocumentNotFound)]
    [InlineData(unchecked((int)0x80070002), OfficeErrorCode.OfficeDocumentNotFound)]
    [InlineData(unchecked((int)0x80070003), OfficeErrorCode.OfficeDocumentNotFound)]
    [InlineData(unchecked((int)0x80030005), OfficeErrorCode.OfficeAccessDenied)]
    [InlineData(unchecked((int)0x80070005), OfficeErrorCode.OfficeAccessDenied)]
    [InlineData(unchecked((int)0x800300FB), OfficeErrorCode.OfficeFileCorrupt)]
    [InlineData(unchecked((int)0x80030109), OfficeErrorCode.OfficeFileCorrupt)]
    public void KnownHresultsMapWithoutDependingOnLocalizedMessages(
        int hresult,
        OfficeErrorCode expected)
    {
        var result = Classify(hresult, "本地化的 Office 错误消息");

        Assert.Equal(expected, result.Code);
    }

    [Fact]
    public void UnknownHresultDoesNotBecomeRetryableFromEnglishMessageText()
    {
        const int unspecifiedFailure = unchecked((int)0x80004005);

        var english = Classify(unspecifiedFailure, "the application is busy");
        var localized = Classify(unspecifiedFailure, "应用程序正忙");

        Assert.Equal("OFFICE_UNCLASSIFIED", english.Code.ToWireName());
        Assert.Equal(english.Code, localized.Code);
    }

    [Fact]
    public void ModalOwnershipRequiresTheDisabledMainWindow()
    {
        var disabledMain = new IntPtr(100);

        Assert.True(IsOwnedByDisabledMain(disabledMain, disabledMain));
        Assert.False(IsOwnedByDisabledMain(new IntPtr(200), disabledMain));
        Assert.False(IsOwnedByDisabledMain(IntPtr.Zero, disabledMain));
    }

    [Fact]
    public void OfficeRegistryVersionsAreOrderedNewestFirst()
    {
        string[] keys = ["Common", "16.0", "9.0", "17.0", "ClickToRun"];

        Assert.Equal(["17.0", "16.0", "9.0"], OrderOfficeVersionKeys(keys));
    }

    [Fact]
    public void PdfPageCountIncludesFlateCompressedObjectStreams()
    {
        string path = Path.Combine(Path.GetTempPath(), $"dcc-office-{Guid.NewGuid():N}.pdf");
        try
        {
            byte[] objectStream = Encoding.Latin1.GetBytes(
                "1 0 << /Type /Page /Parent 3 0 R >> " +
                "2 0 << /Type /Page /Parent 3 0 R >>");
            byte[] compressed = Compress(objectStream);
            using (var file = File.Create(path))
            {
                byte[] prefix = Encoding.ASCII.GetBytes(
                    $"%PDF-1.7\n5 0 obj\n<< /Type /ObjStm /N 2 /First 8 /Filter /FlateDecode /Length {compressed.Length} >>\nstream\n");
                file.Write(prefix);
                file.Write(compressed);
                file.Write(Encoding.ASCII.GetBytes("\nendstream\nendobj\n%%EOF\n"));
            }

            Assert.Equal(2, PdfPageCounter.Count(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void PdfPageCountDoesNotCountThePagesTreeNode()
    {
        string path = Path.Combine(Path.GetTempPath(), $"dcc-office-{Guid.NewGuid():N}.pdf");
        try
        {
            File.WriteAllText(
                path,
                "%PDF-1.4\n1 0 obj << /Type /Pages /Count 1 >> endobj\n" +
                "2 0 obj << /Type /Page /Parent 1 0 R >> endobj\n%%EOF\n",
                Encoding.ASCII);

            Assert.Equal(1, PdfPageCounter.Count(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    private static OfficeComException Classify(int hresult, string message)
    {
        return OfficeComBackend.MapComException(
            new COMException(message, hresult),
            "test operation");
    }

    private static bool IsOwnedByDisabledMain(IntPtr owner, IntPtr disabledMain) =>
        ModalDialogDetector.IsOwnedByDisabledMain(owner, disabledMain);

    private static string[] OrderOfficeVersionKeys(IEnumerable<string> keys) =>
        OfficeComBackend.OrderOfficeVersionKeys(keys).ToArray();

    private static byte[] Compress(byte[] value)
    {
        using var output = new MemoryStream();
        using (var compressor = new ZLibStream(output, CompressionLevel.SmallestSize, leaveOpen: true))
        {
            compressor.Write(value);
        }
        return output.ToArray();
    }
}
