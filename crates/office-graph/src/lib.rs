//! dcc-mcp-office-graph — Microsoft Graph connector (proposal §6.3).
//!
//! Phase 3. Scope: OneDrive/SharePoint files, driveItem format conversion
//! (support matrix must be queried, never assumed — MS-05), Excel Workbook
//! sessions for multi-call workflows (MS-06/07), Outlook/OneNote later.
//!
//! M0 placeholder: the data shapes only. Auth (device code / broker) and
//! HTTP land in Phase 3.

#![forbid(unsafe_code)]

use serde::{Deserialize, Serialize};

/// Graph Workbook session — required for multi-call Excel workflows to
/// unify persistence behaviour (proposal §6.3).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct WorkbookSession {
    pub id: String,
    #[serde(default)]
    pub persist_changes: bool,
}

/// Known-convertible format pair, sourced from a maintained support matrix
/// (never hard-coded as "anything to anything").
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct ConversionPair {
    pub from: String,
    pub to: String,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn workbook_session_defaults_do_not_persist() {
        let s: WorkbookSession = serde_json::from_str(r#"{"id":"wb-1"}"#).unwrap();
        assert!(!s.persist_changes);
    }
}
