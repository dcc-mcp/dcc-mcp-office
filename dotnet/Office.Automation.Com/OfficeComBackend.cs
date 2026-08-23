using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32;
using Office.Automation.Runtime;

namespace Office.Automation.Com;

public enum OfficeAppKind
{
    PowerPoint,
    Word,
    Excel,
}

/// <summary>
/// Base class for the per-application COM backends (proposal §6.2 / §9).
///
/// Contract every backend upholds:
///   - one Application instance per sidecar, created lazily on first use,
///   - every COM call runs on the single STA queue (proposal §9.2),
///   - AutomationSecurity forced to disable macros, display alerts off,
///     external link auto-update off (proposal §19 — deny by default),
///   - documents open read-only and close quietly after each request,
///   - leaf COM objects are request-scoped (ComLease) and never escape the
///     STA thread (proposal §9.3),
///   - soft timeouts report OFFICE_MODAL_DIALOG when a modal window of the
///     Office process blocks the request, OFFICE_RPC_TIMEOUT otherwise,
///   - repeated soft timeouts trigger sidecar recovery (app quit/relaunch).
/// </summary>
public abstract class OfficeComBackend : IDisposable
{
    private static class ComHResult
    {
        public const uint RpcECallRejected = 0x80010001;
        public const uint RpcEServerCallRetryLater = 0x8001010A;
        public const uint RegDbEClassNotRegistered = 0x80040154;
        public const uint StgEFileNotFound = 0x80030002;
        public const uint StgEPathNotFound = 0x80030003;
        public const uint StgEAccessDenied = 0x80030005;
        public const uint StgEShareViolation = 0x80030020;
        public const uint StgELockViolation = 0x80030021;
        public const uint StgEInvalidHeader = 0x800300FB;
        public const uint StgEDocFileCorrupt = 0x80030109;
        public const uint Win32FileNotFound = 0x80070002;
        public const uint Win32PathNotFound = 0x80070003;
        public const uint Win32AccessDenied = 0x80070005;
        public const uint Win32SharingViolation = 0x80070020;
        public const uint Win32LockViolation = 0x80070021;
    }

    /// <summary>Default per-request soft timeout.</summary>
    private readonly OfficeAppKind _kind;
    private readonly StaDispatcher _sta;
    private readonly TimeSpan _requestTimeout;
    private readonly int _timeoutStreakForRecovery;
    private dynamic? _application;
    private bool _attached;
    private int? _processId;
    private OfficeSecurityPosture? _securityPosture;
    private int _timeoutStreak;
    private bool _disposed;

    protected OfficeComBackend(
        OfficeAppKind kind,
        StaDispatcher sta,
        TimeSpan? requestTimeout = null,
        int timeoutStreakForRecovery = 2)
    {
        if ((requestTimeout is not null && requestTimeout <= TimeSpan.Zero)
            || timeoutStreakForRecovery < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(requestTimeout));
        }
        _kind = kind;
        _sta = sta;
        _requestTimeout = requestTimeout ?? TimeSpan.FromSeconds(120);
        _timeoutStreakForRecovery = timeoutStreakForRecovery;
    }

    public string AppName => _kind switch
    {
        OfficeAppKind.PowerPoint => "powerpoint",
        OfficeAppKind.Word => "word",
        OfficeAppKind.Excel => "excel",
        _ => throw new ArgumentOutOfRangeException(),
    };

    public string ProgId => _kind switch
    {
        OfficeAppKind.PowerPoint => "PowerPoint.Application",
        OfficeAppKind.Word => "Word.Application",
        OfficeAppKind.Excel => "Excel.Application",
        _ => throw new ArgumentOutOfRangeException(),
    };

    /// <summary>presentation | document | workbook.</summary>
    protected abstract string DocumentKind { get; }

    public bool IsAttached => _attached;

    public int? OfficeProcessId => _processId;

    public OfficeSecurityPosture? SecurityPosture => _securityPosture;

    protected StaDispatcher Sta => _sta;

    /// <summary>The live Application RCW; requires <see cref="Attach"/> first.</summary>
    protected dynamic Application =>
        _application ?? throw new InvalidOperationException($"{AppName} backend is not attached");

    /// <summary>
    /// Creates the Application instance on the STA thread. Throws
    /// <see cref="OfficeComException"/> with OFFICE_APP_NOT_INSTALLED when the
    /// ProgID is missing (proposal §10.2 progressive discovery).
    /// </summary>
    public void Attach(TimeSpan timeout)
    {
        ThrowIfDisposed();
        if (_attached)
        {
            return;
        }

        AttachState state;
        try
        {
            state = _sta.Post(CreateApplication, timeout);
        }
        catch (StaSoftTimeoutException)
        {
            throw MapTimeout($"attach {AppName}", timeout);
        }
        catch (StaDispatcherBusyException ex)
        {
            throw MapBusy($"attach {AppName}", ex);
        }

        _application = state.Instance;
        _processId = state.ProcessId;
        _securityPosture = state.SecurityPosture;
        _attached = true;
        Interlocked.Exchange(ref _timeoutStreak, 0);
    }

    private AttachState CreateApplication()
    {
        Type? type;
        try
        {
            type = Type.GetTypeFromProgID(ProgId);
        }
        catch (Exception ex)
        {
            throw new OfficeComException(OfficeErrorCode.OfficeAppNotInstalled,
                $"{AppName} is not installed (ProgID {ProgId} unresolvable): {ex.Message}");
        }
        if (type is null)
        {
            throw new OfficeComException(OfficeErrorCode.OfficeAppNotInstalled,
                $"{AppName} is not installed (ProgID {ProgId} missing from the registry)");
        }

        dynamic instance;
        try
        {
            instance = Activator.CreateInstance(type)!;
        }
        catch (COMException ex)
        {
            throw MapComException(ex, $"attach {AppName}");
        }
        OfficeSecurityPosture securityPosture;
        try
        {
            securityPosture = ApplySecurityDefaults(instance);
        }
        catch
        {
            try
            {
                instance.Quit();
            }
            catch (Exception cleanupError)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"Office cleanup after security failure also failed: {cleanupError.Message}");
            }
            ComObjectLifecycle.ReleaseRequestScoped((object?)instance);
            throw;
        }
        try
        {
            instance.Visible = false;
        }
        catch
        {
            // PowerPoint has no Visible in headless automation; ignore.
        }
        return new AttachState
        {
            Instance = instance,
            ProcessId = TryGetProcessId(instance),
            SecurityPosture = securityPosture,
        };
    }

    /// <summary>Attach only if the app is installed; false when absent.</summary>
    public bool TryAttach(TimeSpan timeout)
    {
        try
        {
            Attach(timeout);
            return true;
        }
        catch (OfficeComException ex) when (ex.Code == OfficeErrorCode.OfficeAppNotInstalled)
        {
            return false;
        }
    }

    public OfficeAppInfo GetApplicationInfo()
    {
        if (!_attached || _application is null)
        {
            return new OfficeAppInfo { Name = AppName };
        }
        dynamic app = _application;
        string version = SafeGet(() => (string?)app.Version) ?? "";
        string language = SafeGet(ReadOfficeUiLanguage) ?? "";
        return new OfficeAppInfo
        {
            Name = AppName,
            Version = version,
            Bitness = Environment.Is64BitProcess ? "x64" : "x86",
            Language = language,
        };
    }

    // ---------------------------------------------------------------- core ops

    public abstract FileConvertOutcome ConvertToPdf(string path, string outputPath);

    public abstract InspectOutcome Inspect(string path);

    public abstract ReplaceOutcome ReplaceText(
        string path, IReadOnlyList<ReplaceRuleInput> rules, IReadOnlyList<string> scope, bool dryRun);

    /// <summary>Slide preview export; only PowerPoint implements this.</summary>
    public virtual List<SlidePreviewOutcome>? ExportSlidePreviews(
        string path, string outputDirectory, int width, int height) => null;

    // ------------------------------------------------------------- shared pieces

    /// <summary>Runs void work on the STA queue with the standard error ladder.</summary>
    protected void RunRequest(
        string context,
        Action work,
        TimeSpan? timeout = null,
        bool mayWrite = false)
    {
        ThrowIfDisposed();
        try
        {
            _sta.Post(work, timeout ?? _requestTimeout);
            Interlocked.Exchange(ref _timeoutStreak, 0);
        }
        catch (StaSoftTimeoutException)
        {
            throw MapTimeout(context, timeout ?? _requestTimeout, mayWrite);
        }
        catch (StaDispatcherBusyException ex)
        {
            throw MapBusy(context, ex);
        }
        catch (OfficeComException)
        {
            throw;
        }
        catch (COMException ex)
        {
            throw MapComException(ex, context, mayWrite);
        }
    }

    /// <summary>Runs one request on the STA queue with the standard error ladder.</summary>
    protected T RunRequest<T>(
        string context,
        Func<T> work,
        TimeSpan? timeout = null,
        bool mayWrite = false)
    {
        ThrowIfDisposed();
        try
        {
            T result = _sta.Post(work, timeout ?? _requestTimeout);
            Interlocked.Exchange(ref _timeoutStreak, 0);
            return result;
        }
        catch (StaSoftTimeoutException)
        {
            throw MapTimeout(context, timeout ?? _requestTimeout, mayWrite);
        }
        catch (StaDispatcherBusyException ex)
        {
            throw MapBusy(context, ex);
        }
        catch (OfficeComException)
        {
            throw;
        }
        catch (COMException ex)
        {
            throw MapComException(ex, context, mayWrite);
        }
    }

    /// <summary>Timeout ladder: modal dialog detection, then sidecar recovery.</summary>
    private OfficeComException MapTimeout(
        string context,
        TimeSpan? timeout = null,
        bool mayWrite = false)
    {
        int streak = Interlocked.Increment(ref _timeoutStreak);
        var modalTitle = _processId is int pid
            ? ModalDialogDetector.FindModalDialogTitle(pid)
            : null;
        if (streak >= _timeoutStreakForRecovery)
        {
            RecoverAfterTimeout();
            Interlocked.Exchange(ref _timeoutStreak, 0);
            return new OfficeComException(
                OfficeErrorCode.OfficeBackendUnavailable,
                $"{context}: {AppName} sidecar recovered after repeated request timeouts.",
                indeterminate: mayWrite);
        }
        if (modalTitle is not null)
        {
            return new OfficeComException(
                OfficeErrorCode.OfficeModalDialog,
                $"{context}: a modal dialog blocks {AppName}: {modalTitle}",
                indeterminate: mayWrite);
        }
        return new OfficeComException(
            OfficeErrorCode.OfficeRpcTimeout,
            $"{context}: request exceeded {(timeout ?? _requestTimeout).TotalSeconds:F0}s soft timeout.",
            indeterminate: mayWrite);
    }

    private static OfficeComException MapBusy(string context, StaDispatcherBusyException ex) =>
        new(OfficeErrorCode.OfficeAppBusy, $"{context}: {ex.Message}", ex);

    /// <summary>Quit and reset the Application instance (sidecar recovery, §9.2).</summary>
    private void RecoverAfterTimeout()
    {
        try
        {
            _sta.Post(() =>
            {
                if (_application is not null)
                {
                    try { _application.Quit(); } catch { }
                }
            }, TimeSpan.FromSeconds(30));
        }
        catch
        {
            // Quit hung too: force-kill the Office process and let the next
            // Attach create a fresh instance.
            if (_processId is int pid)
            {
                try { System.Diagnostics.Process.GetProcessById(pid).Kill(); } catch { }
            }
        }
        ComObjectLifecycle.ReleaseRequestScoped((object?)_application);
        _application = null;
        _attached = false;
        _processId = null;
        _securityPosture = null;
    }

    protected internal static OfficeComException MapComException(
        COMException ex,
        string context,
        bool mayHaveWritten = false)
    {
        uint hr = unchecked((uint)ex.HResult);
        OfficeErrorCode code = hr switch
        {
            ComHResult.RpcECallRejected or
            ComHResult.RpcEServerCallRetryLater => OfficeErrorCode.OfficeAppBusy,
            ComHResult.RegDbEClassNotRegistered => OfficeErrorCode.OfficeAppNotInstalled,
            ComHResult.StgEShareViolation or
            ComHResult.StgELockViolation or
            ComHResult.Win32SharingViolation or
            ComHResult.Win32LockViolation => OfficeErrorCode.OfficeDocumentLocked,
            ComHResult.StgEFileNotFound or
            ComHResult.StgEPathNotFound or
            ComHResult.Win32FileNotFound or
            ComHResult.Win32PathNotFound => OfficeErrorCode.OfficeDocumentNotFound,
            ComHResult.StgEAccessDenied or
            ComHResult.Win32AccessDenied => OfficeErrorCode.OfficeAccessDenied,
            ComHResult.StgEInvalidHeader or
            ComHResult.StgEDocFileCorrupt => OfficeErrorCode.OfficeFileCorrupt,
            _ => OfficeErrorCode.OfficeUnclassified,
        };
        return new OfficeComException(
            code,
            $"{context}: {ex.Message}",
            ex,
            indeterminate: mayHaveWritten);
    }

    protected static OfficeComException DocumentNotFound(string path) =>
        new(OfficeErrorCode.OfficeDocumentNotFound, $"document not found: {path}");

    /// <summary>Fails when the output is missing, empty, or not a PDF (§15.1 validation).</summary>
    protected static void ValidatePdfOutput(string outputPath, string inputPath)
    {
        var file = new FileInfo(outputPath);
        if (!file.Exists || file.Length == 0)
        {
            throw new OfficeComException(OfficeErrorCode.OfficeRenderTimeout,
                $"PDF export produced no output for {inputPath}");
        }
        using var stream = file.OpenRead();
        Span<byte> magic = stackalloc byte[5];
        if (stream.Read(magic) < 4 || Encoding.ASCII.GetString(magic[..4]) != "%PDF")
        {
            throw new OfficeComException(OfficeErrorCode.OfficeRenderTimeout,
                $"PDF export for {inputPath} is not a valid PDF");
        }
    }

    protected static int CountPdfPages(string outputPath)
    {
        return PdfPageCounter.Count(outputPath);
    }

    /// <summary>Opens a document read-only; callers wrap the result in a using.</summary>
    protected abstract ComLease OpenReadOnly(string path);

    /// <summary>Opens a document with write access (replace-text commit path).</summary>
    protected abstract ComLease OpenEditable(string path);

    /// <summary>Opens in a mode suitable for rendering (windowed but hidden).</summary>
    protected virtual ComLease OpenForRender(string path) => OpenReadOnly(path);

    /// <summary>Persists an editable document after mutations.</summary>
    protected abstract void SaveEditable(dynamic document);

    protected abstract void ExportPdf(dynamic document, string outputPath);

    protected abstract void CloseQuietly(dynamic document);

    protected abstract OfficeSecurityPosture ApplySecurityDefaults(dynamic app);

    /// <summary>Applies a security setting and proves the live value before use.</summary>
    internal static T VerifySecuritySetting<T>(
        string name,
        Action apply,
        Func<T> observe,
        T expected,
        OfficeErrorCode errorCode)
    {
        try
        {
            apply();
            T actual = observe();
            if (!EqualityComparer<T>.Default.Equals(actual, expected))
            {
                throw new OfficeComException(
                    errorCode,
                    $"security setting {name} read back '{actual}', expected '{expected}'");
            }
            return actual;
        }
        catch (OfficeComException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new OfficeComException(
                errorCode,
                $"security setting {name} could not be enforced: {ex.Message}",
                ex);
        }
    }

    // ------------------------------------------------------------- helpers

    /// <summary>Attaches the app, mapping failure to an error instead of throwing.</summary>
    protected OfficeComException? AttachOrError()
    {
        try
        {
            Attach(_requestTimeout);
            return null;
        }
        catch (OfficeComException ex)
        {
            return ex;
        }
        catch (Exception ex)
        {
            return new OfficeComException(OfficeErrorCode.OfficeBackendUnavailable,
                $"attach {AppName} failed: {ex.Message}", ex);
        }
    }

    /// <summary>Attaches the app or throws the mapped error.</summary>
    protected void AttachOrThrow()
    {
        var error = AttachOrError();
        if (error is not null)
        {
            throw error;
        }
    }

    protected static FileConvertOutcome ErrorOutcome(string path, string outputPath, OfficeComException ex) =>
        new()
        {
            InputPath = path,
            OutputPath = outputPath,
            Ok = false,
            ErrorCode = ex.Code.ToWireName(),
            Error = ex.Message,
            Indeterminate = ex.Indeterminate,
        };

    protected static double SafeNumber(dynamic value)
    {
        try { return Convert.ToDouble(value); }
        catch { return 0; }
    }

    private static int? TryGetProcessId(dynamic app)
    {
        try
        {
            long hwnd = (long)app.HWND;
            if (hwnd != 0)
            {
                uint pid;
                GetWindowThreadProcessId(new IntPtr(hwnd), out pid);
                return (int)pid;
            }
        }
        catch
        {
            // Headless instances have no HWND; modal detection simply stays off.
        }
        return null;
    }

    private static string? ReadOfficeUiLanguage()
    {
        try
        {
            foreach (RegistryHive hive in new[] { RegistryHive.LocalMachine, RegistryHive.CurrentUser })
            {
                foreach (RegistryView view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
                {
                    using RegistryKey baseKey = RegistryKey.OpenBaseKey(hive, view);
                    using RegistryKey? officeKey = baseKey.OpenSubKey(@"SOFTWARE\Microsoft\Office");
                    if (officeKey is null)
                    {
                        continue;
                    }
                    foreach (string versionKey in OrderOfficeVersionKeys(officeKey.GetSubKeyNames()))
                    {
                        using RegistryKey? languageKey = officeKey.OpenSubKey(
                            $@"{versionKey}\Common\LanguageResources");
                        if (languageKey?.GetValue("UILanguage") is int lcid && lcid != 0)
                        {
                            return new System.Globalization.CultureInfo(lcid).Name;
                        }
                    }
                }
            }
        }
        catch
        {
        }
        return null;
    }

    internal static IEnumerable<string> OrderOfficeVersionKeys(IEnumerable<string> keys) =>
        keys.Select(key => (Key: key, Parsed: Version.TryParse(key, out var version) ? version : null))
            .Where(item => item.Parsed is not null)
            .OrderByDescending(item => item.Parsed)
            .Select(item => item.Key);

    private static string? SafeGet(Func<string?> getter)
    {
        try { return getter(); }
        catch { return null; }
    }

    private void ThrowIfDisposed()
    {
        if (_disposed)
        {
            throw new ObjectDisposedException(GetType().Name);
        }
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }
        _disposed = true;
        if (_application is null)
        {
            return;
        }
        dynamic app = _application;
        try
        {
            _sta.Post(() =>
            {
                try { app.Quit(); } catch { }
            }, TimeSpan.FromSeconds(30));
        }
        catch
        {
            if (_processId is int pid)
            {
                try { System.Diagnostics.Process.GetProcessById(pid).Kill(); } catch { }
            }
        }
        ComObjectLifecycle.ReleaseRequestScoped((object?)app);
        _application = null;
        _securityPosture = null;
    }

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    /// <summary>STA-marshalled attach result (plain class, no COM identity).</summary>
    private sealed class AttachState
    {
        public dynamic? Instance;
        public int? ProcessId;
        public OfficeSecurityPosture? SecurityPosture;
    }
}
