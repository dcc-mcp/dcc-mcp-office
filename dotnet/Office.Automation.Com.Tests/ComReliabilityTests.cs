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
    public void ComFailureDuringAWriteIsMarkedIndeterminate()
    {
        var error = OfficeComBackend.MapComException(
            new COMException("save failed", unchecked((int)0x80004005)),
            "save document",
            mayHaveWritten: true);

        Assert.True(error.Indeterminate);
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

    [Fact]
    public void SoftTimedOutWriteIsMarkedIndeterminate()
    {
        using var sta = new StaDispatcher();
        using var backend = new TimeoutProbeBackend(sta);

        OfficeComException error = Assert.Throws<OfficeComException>(() =>
            backend.RunTimedWork(mayWrite: true));

        Assert.Equal(OfficeErrorCode.OfficeRpcTimeout, error.Code);
        Assert.True(error.Indeterminate);
    }

    [Fact]
    public void SoftTimedOutReadIsNotMarkedIndeterminate()
    {
        using var sta = new StaDispatcher();
        using var backend = new TimeoutProbeBackend(sta);

        OfficeComException error = Assert.Throws<OfficeComException>(() =>
            backend.RunTimedWork(mayWrite: false));

        Assert.False(error.Indeterminate);
    }

    [Fact]
    public void SecuritySettingFailureAndReadbackMismatchFailClosed()
    {
        OfficeComException assignment = Assert.Throws<OfficeComException>(() =>
            OfficeComBackend.VerifySecuritySetting(
                "AutomationSecurity",
                apply: () => throw new COMException("denied"),
                observe: () => 3,
                expected: 3,
                OfficeErrorCode.OfficeMacroBlocked));
        OfficeComException mismatch = Assert.Throws<OfficeComException>(() =>
            OfficeComBackend.VerifySecuritySetting(
                "AutomationSecurity",
                apply: () => { },
                observe: () => 1,
                expected: 3,
                OfficeErrorCode.OfficeMacroBlocked));

        Assert.Equal(OfficeErrorCode.OfficeMacroBlocked, assignment.Code);
        Assert.Equal(OfficeErrorCode.OfficeMacroBlocked, mismatch.Code);
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

    private sealed class TimeoutProbeBackend : OfficeComBackend
    {
        internal TimeoutProbeBackend(StaDispatcher sta)
            : base(OfficeAppKind.PowerPoint, sta)
        {
        }

        protected override string DocumentKind => "test";

        internal void RunTimedWork(bool mayWrite) =>
            RunRequest(
                "test request",
                () => Thread.Sleep(100),
                TimeSpan.FromMilliseconds(10),
                mayWrite);

        public override FileConvertOutcome ConvertToPdf(string path, string outputPath) =>
            throw new NotSupportedException();

        public override InspectOutcome Inspect(string path) => throw new NotSupportedException();

        public override ReplaceOutcome ReplaceText(
            string path,
            IReadOnlyList<ReplaceRuleInput> rules,
            IReadOnlyList<string> scope,
            bool dryRun) => throw new NotSupportedException();

        protected override ComLease OpenReadOnly(string path) => throw new NotSupportedException();

        protected override ComLease OpenEditable(string path) => throw new NotSupportedException();

        protected override void SaveEditable(dynamic document) => throw new NotSupportedException();

        protected override void ExportPdf(dynamic document, string outputPath) =>
            throw new NotSupportedException();

        protected override void CloseQuietly(dynamic document) => throw new NotSupportedException();

        protected override OfficeSecurityPosture ApplySecurityDefaults(dynamic app) =>
            throw new NotSupportedException();
    }
}
