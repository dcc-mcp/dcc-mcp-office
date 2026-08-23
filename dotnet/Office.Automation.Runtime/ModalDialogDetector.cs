using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace Office.Automation.Runtime;

/// <summary>
/// Detects modal dialogs owned by an Office process (proposal §20 MVP
/// criterion 8 "模态框检测").
///
/// Heuristic: the Office main window is disabled while a visible, enabled
/// popup window of the same process is on screen — the classic modal-dialog
/// state. The detector returns the popup's title so the error surfaced to
/// the gateway names the dialog instead of a generic timeout.
/// </summary>
public static class ModalDialogDetector
{
    /// <summary>Finds a modal dialog of <paramref name="processId"/>, or null.</summary>
    public static string? FindModalDialogTitle(int processId)
    {
        // First pass: collect all top-level windows of the process.
        var windows = new List<WindowInfo>();
        NativeMethods.EnumWindows((hWnd, _) =>
        {
            NativeMethods.GetWindowThreadProcessId(hWnd, out uint pid);
            if (pid == (uint)processId && NativeMethods.IsWindowVisible(hWnd))
            {
                windows.Add(new WindowInfo(hWnd, NativeMethods.IsWindowEnabled(hWnd), GetTitle(hWnd)));
            }
            return true;
        }, IntPtr.Zero);

        var disabledMain = windows.FirstOrDefault(w => !w.Enabled);
        if (disabledMain.hWnd == IntPtr.Zero)
        {
            return null;
        }

        // A modal dialog is a visible + enabled window owned by a disabled one.
        foreach (var popup in windows)
        {
            if (!popup.Enabled)
            {
                continue;
            }
            IntPtr owner = NativeMethods.GetWindow(popup.hWnd, 4 /* GW_OWNER */);
            if (IsOwnedByDisabledMain(owner, disabledMain.hWnd))
            {
                return popup.Title.Length > 0 ? popup.Title : "(untitled dialog)";
            }
        }
        return null;
    }

    internal static bool IsOwnedByDisabledMain(IntPtr owner, IntPtr disabledMain) =>
        owner != IntPtr.Zero && owner == disabledMain;

    private static string GetTitle(IntPtr hWnd)
    {
        var builder = new StringBuilder(256);
        _ = NativeMethods.GetWindowText(hWnd, builder, builder.Capacity);
        return builder.ToString();
    }

    private readonly record struct WindowInfo(IntPtr hWnd, bool Enabled, string Title);

    private static class NativeMethods
    {
        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowEnabled(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    }
}
