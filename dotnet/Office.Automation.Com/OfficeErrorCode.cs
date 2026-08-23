using System.Text.RegularExpressions;

namespace Office.Automation.Com;

/// <summary>
/// Error codes the COM backend surfaces (proposal §20). Wire names are the
/// SCREAMING_SNAKE strings the Rust gateway expects. The machine-readable
/// office-rpc catalog owns the canonical set; Office-free Host tests require
/// this enum and the generated Rust enum to match it exactly.
/// </summary>
public enum OfficeErrorCode
{
    OfficeInvalidRequest,
    OfficeAppNotInstalled,
    OfficeAppVersionUnsupported,
    OfficeAppBusy,
    OfficeModalDialog,
    OfficeProtectedView,
    OfficeAccessDenied,
    OfficeDocumentNotFound,
    OfficeDocumentLocked,
    OfficeDocumentConflict,
    OfficeFileCorrupt,
    OfficeMacroBlocked,
    OfficeExternalLinkBlocked,
    OfficeCapabilityUnsupported,
    OfficeBackendUnavailable,
    OfficeRpcTimeout,
    OfficeRenderTimeout,
    OfficeGraphThrottled,
    OfficeGraphAuthRequired,
    OfficeUserConfirmationRequired,
    OfficePartialSuccess,
    OfficeUnclassified,
}

public static class OfficeErrorCodeExtensions
{
    private static readonly Regex PascalSplit = new("(?<=[a-z0-9])(?=[A-Z])", RegexOptions.Compiled);

    /// <summary>
    /// SCREAMING_SNAKE wire name, matching the Rust protocol crate's
    /// serde(rename_all = "SCREAMING_SNAKE_CASE") on OfficeErrorCode —
    /// member names already carry the OFFICE prefix.
    /// </summary>
    public static string ToWireName(this OfficeErrorCode code) =>
        PascalSplit.Replace(code.ToString(), "_").ToUpperInvariant();
}

/// <summary>
/// Raised by the COM backend for any mapped failure (proposal §20: the RPC
/// error.code carries the OFFICE_* code, error.message stays human-readable).
/// </summary>
public sealed class OfficeComException : Exception
{
    public OfficeComException(OfficeErrorCode code, string message)
        : base(message)
    {
        Code = code;
    }

    public OfficeComException(OfficeErrorCode code, string message, Exception inner)
        : base(message, inner)
    {
        Code = code;
    }

    public OfficeErrorCode Code { get; }
}
