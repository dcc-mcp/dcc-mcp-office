using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
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
    /// <summary>Default per-request soft timeout.</summary>
    protected static readonly TimeSpan RequestTimeout = TimeSpan.FromSeconds(120);

    /// <summary>Two consecutive soft timeouts trip sidecar recovery (§9.2).</summary>
    private const int TimeoutStreakForRecovery = 2;

    private readonly OfficeAppKind _kind;
    private readonly StaDispatcher _sta;
    private dynamic? _application;
    private bool _attached;
    private int? _processId;
    private int _timeoutStreak;
    private bool _disposed;

    protected OfficeComBackend(OfficeAppKind kind, StaDispatcher sta)
    {
        _kind = kind;
        _sta = sta;
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

        AttachState state = _sta.Post(() =>
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
            ApplySecurityDefaults(instance);
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
            };
        }, timeout);

        _application = state.Instance;
        _processId = state.ProcessId;
        _attached = true;
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
    protected void RunRequest(string context, Action work, TimeSpan? timeout = null)
    {
        ThrowIfDisposed();
        try
        {
            _sta.Post(work, timeout ?? RequestTimeout);
            Interlocked.Exchange(ref _timeoutStreak, 0);
        }
        catch (StaSoftTimeoutException)
        {
            throw MapTimeout(context);
        }
        catch (OfficeComException)
        {
            throw;
        }
        catch (COMException ex)
        {
            throw MapComException(ex, context);
        }
    }

    /// <summary>Runs one request on the STA queue with the standard error ladder.</summary>
    protected T RunRequest<T>(string context, Func<T> work, TimeSpan? timeout = null)
    {
        ThrowIfDisposed();
        try
        {
            T result = _sta.Post(work, timeout ?? RequestTimeout);
            Interlocked.Exchange(ref _timeoutStreak, 0);
            return result;
        }
        catch (StaSoftTimeoutException)
        {
            throw MapTimeout(context);
        }
        catch (OfficeComException)
        {
            throw;
        }
        catch (COMException ex)
        {
            throw MapComException(ex, context);
        }
    }

    /// <summary>Timeout ladder: modal dialog detection, then sidecar recovery.</summary>
    private OfficeComException MapTimeout(string context)
    {
        int streak = Interlocked.Increment(ref _timeoutStreak);
        var modalTitle = _processId is int pid
            ? ModalDialogDetector.FindModalDialogTitle(pid)
            : null;
        if (streak >= TimeoutStreakForRecovery)
        {
            RecoverAfterTimeout();
            return new OfficeComException(OfficeErrorCode.OfficeBackendUnavailable,
                $"{context}: {AppName} sidecar recovered after repeated request timeouts.");
        }
        if (modalTitle is not null)
        {
            return new OfficeComException(OfficeErrorCode.OfficeModalDialog,
                $"{context}: a modal dialog blocks {AppName}: {modalTitle}");
        }
        return new OfficeComException(OfficeErrorCode.OfficeRpcTimeout,
            $"{context}: request exceeded {RequestTimeout.TotalSeconds:F0}s soft timeout.");
    }

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
    }

    protected static OfficeComException MapComException(COMException ex, string context)
    {
        uint hr = unchecked((uint)ex.HResult);
        if (hr == 0x80010001 /* RPC_E_CALL_REJECTED */)
        {
            return new OfficeComException(OfficeErrorCode.OfficeAppBusy,
                $"{context}: {ex.Message}", ex);
        }
        string message = ex.Message;
        if (message.Contains("busy", StringComparison.OrdinalIgnoreCase))
        {
            return new OfficeComException(OfficeErrorCode.OfficeAppBusy, $"{context}: {message}", ex);
        }
        if (message.Contains("not installed", StringComparison.OrdinalIgnoreCase)
            || message.Contains("is not available", StringComparison.OrdinalIgnoreCase))
        {
            return new OfficeComException(OfficeErrorCode.OfficeAppNotInstalled, $"{context}: {message}", ex);
        }
        if (message.Contains("protected view", StringComparison.OrdinalIgnoreCase))
        {
            return new OfficeComException(OfficeErrorCode.OfficeProtectedView, $"{context}: {message}", ex);
        }
        if (message.Contains("locked", StringComparison.OrdinalIgnoreCase))
        {
            return new OfficeComException(OfficeErrorCode.OfficeDocumentLocked, $"{context}: {message}", ex);
        }
        if (message.Contains("not found", StringComparison.OrdinalIgnoreCase)
            || message.Contains("does not exist", StringComparison.OrdinalIgnoreCase)
            || message.Contains("could not be found", StringComparison.OrdinalIgnoreCase))
        {
            return new OfficeComException(OfficeErrorCode.OfficeDocumentNotFound, $"{context}: {message}", ex);
        }
        if (message.Contains("corrupt", StringComparison.OrdinalIgnoreCase))
        {
            return new OfficeComException(OfficeErrorCode.OfficeFileCorrupt, $"{context}: {message}", ex);
        }
        if (message.Contains("macro", StringComparison.OrdinalIgnoreCase))
        {
            return new OfficeComException(OfficeErrorCode.OfficeMacroBlocked, $"{context}: {message}", ex);
        }
        return new OfficeComException(OfficeErrorCode.OfficeBackendUnavailable, $"{context}: {message}", ex);
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

    private static readonly Regex PageRegex =
        new(@"/Type\s*/Page[^s]", RegexOptions.Compiled);

    protected static int CountPdfPages(string outputPath)
    {
        string text = File.ReadAllText(outputPath, Encoding.Latin1);
        return PageRegex.Matches(text).Count;
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

    protected abstract void ApplySecurityDefaults(dynamic app);

    // ------------------------------------------------------------- helpers

    /// <summary>Attaches the app, mapping failure to an error instead of throwing.</summary>
    protected OfficeComException? AttachOrError()
    {
        try
        {
            Attach(RequestTimeout);
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
            foreach (var hive in new[] { Microsoft.Win32.Registry.LocalMachine, Microsoft.Win32.Registry.CurrentUser })
            {
                using var key = hive.OpenSubKey(@"SOFTWARE\Microsoft\Office\16.0\Common\LanguageResources");
                if (key?.GetValue("UILanguage") is int lcid && lcid != 0)
                {
                    return new System.Globalization.CultureInfo(lcid).Name;
                }
            }
        }
        catch
        {
        }
        return null;
    }

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
    }

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    /// <summary>STA-marshalled attach result (plain class, no COM identity).</summary>
    private sealed class AttachState
    {
        public dynamic? Instance;
        public int? ProcessId;
    }
}
