//! dcc-mcp-office-jobs — job phases layered on dcc-mcp-job (proposal §16).
//!
//! Batch operations are always jobs: a single MCP request never blocks on a
//! whole batch. dcc-mcp-job owns async tracking + persistence; this crate
//! adds the Office-specific phases (approval gate, validation, publishing)
//! and the per-item result bookkeeping the proposal §16 requires.

#![forbid(unsafe_code)]

use serde::{Deserialize, Serialize};

/// Job phase (proposal §16).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum OfficeJobPhase {
    Queued,
    Planning,
    WaitingForApproval,
    Running,
    Validating,
    Publishing,
    Succeeded,
    PartiallySucceeded,
    Failed,
    Cancelled,
}

impl OfficeJobPhase {
    pub fn is_terminal(self) -> bool {
        matches!(
            self,
            OfficeJobPhase::Succeeded
                | OfficeJobPhase::PartiallySucceeded
                | OfficeJobPhase::Failed
                | OfficeJobPhase::Cancelled
        )
    }

    /// Concurrency rules (proposal §16): Open XML workers parallelise across
    /// files; each COM sidecar is a single STA write queue; same-document
    /// writes are mutually exclusive.
    pub fn allows_parallel_files(self) -> bool {
        matches!(
            self,
            OfficeJobPhase::Running | OfficeJobPhase::Validating | OfficeJobPhase::Publishing
        )
    }
}

/// Per-item execution result — the granularity agents actually read back
/// (proposal §17: never just "succeeded").
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ItemResult {
    pub input_path: String,
    pub ok: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub error: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub backend: Option<String>,
    #[serde(default)]
    pub artefacts: Vec<String>,
    #[serde(default)]
    pub warnings: Vec<String>,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn terminal_phases() {
        assert!(OfficeJobPhase::Succeeded.is_terminal());
        assert!(OfficeJobPhase::PartiallySucceeded.is_terminal());
        assert!(!OfficeJobPhase::WaitingForApproval.is_terminal());
    }

    #[test]
    fn round_trip() {
        let phase = serde_json::to_string(&OfficeJobPhase::WaitingForApproval).unwrap();
        assert_eq!(phase, "\"waiting_for_approval\"");
    }
}
