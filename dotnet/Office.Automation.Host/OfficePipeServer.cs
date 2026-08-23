using System.ComponentModel;
using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using Microsoft.Win32.SafeHandles;

namespace Office.Automation.Host;

/// <summary>
/// Named-pipe JSON-RPC 2.0 server for office-host (proposal §12).
///
/// Pipe: \\.\pipe\dcc-mcp-office-{app}-{user_sid}-{session_id} — same shape
/// the Rust protocol crate builds. Framing: one JSON object per line (UTF-8).
///
/// The pipe is created through CreateNamedPipe with an explicit DACL granting
/// only the current user, SYSTEM and Administrators (§12.1 "按当前用户 SID
/// 配置 ACL"): post-creation SetAccessControl is denied on default-descriptor
/// pipes because the default DACL grants the owner no WRITE_DAC, so the ACL
/// must ride SECURITY_ATTRIBUTES at creation time.
/// </summary>
public sealed class OfficePipeServer
{
    private readonly string _app;
    private readonly string _pipeName;
    private readonly Func<string, string> _dispatch;
    private readonly Func<IReadOnlyList<string>>? _drainNotifications;
    private readonly Func<bool>? _shouldStop;

    public OfficePipeServer(
        string app,
        Func<string, string> dispatch,
        string? explicitPipeName = null,
        Func<bool>? shouldStop = null,
        Func<IReadOnlyList<string>>? drainNotifications = null)
    {
        _app = app;
        _pipeName = explicitPipeName ?? BuildPipeName(app);
        _dispatch = dispatch;
        _shouldStop = shouldStop;
        _drainNotifications = drainNotifications;
    }

    public string PipeName => _pipeName;

    public static string BuildPipeName(string app)
    {
        string sid = WindowsIdentity.GetCurrent().User?.Value ?? "unknown-sid";
        int session = System.Diagnostics.Process.GetCurrentProcess().SessionId;
        return $@"\\.\pipe\dcc-mcp-office-{app}-{sid}-{session}";
    }

    /// <summary>
    /// Accepts clients one at a time and serves them until they disconnect.
    /// One sidecar = one app = one serialized command stream (§8.2).
    /// </summary>
    public void Run(CancellationToken cancellation, Func<bool>? shouldStop = null)
    {
        while (!cancellation.IsCancellationRequested && (shouldStop is null || !shouldStop()))
        {
            try
            {
                var server = CreateServerPipe();
                try
                {
                    server.WaitForConnection();
                    Serve(server);
                }
                finally
                {
                    server.Dispose();
                }
            }
            catch (OperationCanceledException)
            {
                return;
            }
            catch (Win32Exception ex)
            {
                // e.g. ERROR_INVALID_NAME for a malformed --pipe-name: report
                // and stop instead of crashing the process silently.
                Console.Error.WriteLine($"[office-host:{_app}] pipe failure ({ex.NativeErrorCode}): {ex.Message}");
                return;
            }
            catch (IOException)
            {
                // A client that vanished mid-request just ends this connection.
                if (cancellation.IsCancellationRequested)
                {
                    return;
                }
            }
        }
    }

    private NamedPipeServerStream CreateServerPipe()
    {
        string userSid = WindowsIdentity.GetCurrent().User?.Value
            ?? throw new InvalidOperationException("no user SID for pipe ACL");
        // D:P — DACL; GA — full access for SYSTEM, Administrators, current user.
        string sddl = $"D:P(A;;GA;;;SY)(A;;GA;;;BA)(A;;GA;;;{userSid})";
        if (!NativeMethods.ConvertStringSecurityDescriptorToSecurityDescriptor(
                sddl, 1, out IntPtr descriptor, IntPtr.Zero))
        {
            throw new Win32Exception(Marshal.GetLastWin32Error(), "ConvertStringSecurityDescriptorToSecurityDescriptor");
        }
        try
        {
            var attributes = new NativeMethods.SECURITY_ATTRIBUTES
            {
                nLength = Marshal.SizeOf<NativeMethods.SECURITY_ATTRIBUTES>(),
                lpSecurityDescriptor = descriptor,
                bInheritHandle = false,
            };
            SafePipeHandle handle = NativeMethods.CreateNamedPipe(
                _pipeName,
                NativeMethods.PIPE_ACCESS_DUPLEX,
                NativeMethods.PIPE_TYPE_BYTE | NativeMethods.PIPE_READMODE_BYTE | NativeMethods.PIPE_WAIT,
                1,
                64 * 1024,
                64 * 1024,
                0,
                ref attributes);
            if (handle.IsInvalid)
            {
                throw new Win32Exception(Marshal.GetLastWin32Error(), "CreateNamedPipe");
            }
            return new NamedPipeServerStream(PipeDirection.InOut, isAsync: false, isConnected: false, handle);
        }
        finally
        {
            NativeMethods.LocalFree(descriptor);
        }
    }

    private void Serve(Stream stream)
    {
        using var reader = new StreamReader(stream, new UTF8Encoding(false), detectEncodingFromByteOrderMarks: false, bufferSize: 64 * 1024, leaveOpen: true);
        using var writer = new StreamWriter(stream, new UTF8Encoding(false), bufferSize: 64 * 1024, leaveOpen: true)
        {
            NewLine = "\n",
        };
        var writerGate = new object();
        using var notificationCancellation = new CancellationTokenSource();
        Task? notificationPump = _drainNotifications is null
            ? null
            : Task.Run(() => PumpNotifications(
                writer,
                writerGate,
                notificationCancellation.Token));
        try
        {
            string? line;
            while ((line = reader.ReadLine()) is not null)
            {
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }
                string response = _dispatch(line);
                lock (writerGate)
                {
                    writer.WriteLine(response);
                    if (_drainNotifications is not null)
                    {
                        foreach (string notification in _drainNotifications())
                        {
                            writer.WriteLine(notification);
                        }
                    }
                    writer.Flush();
                }
                // office.host.shutdown sets the stop flag: leave this connection
                // right after replying instead of blocking on the next line.
                if (_shouldStop is not null && _shouldStop())
                {
                    return;
                }
            }
        }
        finally
        {
            notificationCancellation.Cancel();
            if (notificationPump is not null)
            {
                try { notificationPump.Wait(TimeSpan.FromSeconds(1)); }
                catch (AggregateException) { }
            }
        }
    }

    private void PumpNotifications(
        StreamWriter writer,
        object writerGate,
        CancellationToken cancellation)
    {
        while (!cancellation.IsCancellationRequested)
        {
            IReadOnlyList<string> messages = _drainNotifications!();
            if (messages.Count == 0)
            {
                cancellation.WaitHandle.WaitOne(TimeSpan.FromMilliseconds(20));
                continue;
            }
            try
            {
                lock (writerGate)
                {
                    foreach (string message in messages)
                    {
                        writer.WriteLine(message);
                    }
                    writer.Flush();
                }
            }
            catch (IOException)
            {
                return;
            }
        }
    }

    private static class NativeMethods
    {
        public const uint PIPE_ACCESS_DUPLEX = 0x3;
        public const uint PIPE_TYPE_BYTE = 0x0;
        public const uint PIPE_READMODE_BYTE = 0x0;
        public const uint PIPE_WAIT = 0x0;

        [StructLayout(LayoutKind.Sequential)]
        public struct SECURITY_ATTRIBUTES
        {
            public int nLength;
            public IntPtr lpSecurityDescriptor;
            [MarshalAs(UnmanagedType.Bool)]
            public bool bInheritHandle;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern SafePipeHandle CreateNamedPipe(
            string lpName,
            uint dwOpenMode,
            uint dwPipeMode,
            uint nMaxInstances,
            uint nOutBufferSize,
            uint nInBufferSize,
            uint nDefaultTimeOut,
            ref SECURITY_ATTRIBUTES lpSecurityAttributes);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ConvertStringSecurityDescriptorToSecurityDescriptor(
            string stringSecurityDescriptor,
            uint stringSdRevision,
            out IntPtr securityDescriptor,
            IntPtr securityDescriptorSize);

        [DllImport("kernel32.dll")]
        public static extern IntPtr LocalFree(IntPtr hMem);
    }
}
