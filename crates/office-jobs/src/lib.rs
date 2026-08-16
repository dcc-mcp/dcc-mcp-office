//! dcc-mcp-office-jobs — job phases layered on dcc-mcp-job (proposal §16).
//!
//! Batch operations are always jobs: a single MCP request never blocks on a
//! whole batch. dcc-mcp-job owns async tracking + persistence; this crate
//! adds the Office-specific phases (approval gate, validation, publishing)
//! and the per-item result bookkeeping the proposal §16 requires.
//!
//! The aggregation helpers here are pure (dependency-free): layering onto
//! dcc-mcp-job only needs a phase transition + persistence around them, and
//! that crate has not been published yet (see AGENTS.md dependency map).

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

    /// Terminal phase derived from per-item outcomes (proposal §16): all ok
    /// → Succeeded; a mix → PartiallySucceeded; nothing ok → Failed.
    pub fn from_outcomes(succeeded: usize, files: usize) -> Self {
        if files == 0 || succeeded == 0 {
            OfficeJobPhase::Failed
        } else if succeeded >= files {
            OfficeJobPhase::Succeeded
        } else {
            OfficeJobPhase::PartiallySucceeded
        }
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

/// Job-level summary aggregated from per-item results (proposal §16/§17:
/// agents read per-item granularity, never just "succeeded").
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct BatchSummary {
    pub files: usize,
    pub succeeded: usize,
    pub failed: usize,
    /// Dominant backend across items (first non-empty in item order).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub backend: Option<String>,
    #[serde(default)]
    pub artefacts: Vec<String>,
    #[serde(default)]
    pub warnings: Vec<String>,
}

impl BatchSummary {
    /// Phase this summary maps to (proposal §16).
    pub fn phase(&self) -> OfficeJobPhase {
        OfficeJobPhase::from_outcomes(self.succeeded, self.files)
    }
}

/// Aggregates per-item results (pure; dcc-mcp-job layering wraps this with
/// persistence + phase transitions).
pub fn summarize(items: &[ItemResult]) -> BatchSummary {
    let succeeded = items.iter().filter(|item| item.ok).count();
    let backend = items.iter().find_map(|item| item.backend.clone());
    let mut artefacts = Vec::new();
    let mut warnings = Vec::new();
    for item in items {
        artefacts.extend(item.artefacts.iter().cloned());
        warnings.extend(item.warnings.iter().cloned());
        if let Some(error) = &item.error {
            warnings.push(format!("{}: {error}", item.input_path));
        }
    }
    BatchSummary {
        files: items.len(),
        succeeded,
        failed: items.len() - succeeded,
        backend,
        artefacts,
        warnings,
    }
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

    #[test]
    fn phase_derives_from_outcomes() {
        assert_eq!(OfficeJobPhase::from_outcomes(0, 0), OfficeJobPhase::Failed);
        assert_eq!(
            OfficeJobPhase::from_outcomes(2, 2),
            OfficeJobPhase::Succeeded
        );
        assert_eq!(
            OfficeJobPhase::from_outcomes(1, 3),
            OfficeJobPhase::PartiallySucceeded
        );
        assert_eq!(OfficeJobPhase::from_outcomes(0, 3), OfficeJobPhase::Failed);
    }

    #[test]
    fn summarize_aggregates_items_and_derives_phase() {
        let items = vec![
            ItemResult {
                input_path: "a.pptx".into(),
                ok: true,
                error: None,
                backend: Some("desktop_com".into()),
                artefacts: vec!["a.pdf".into()],
                warnings: vec![],
            },
            ItemResult {
                input_path: "b.pptx".into(),
                ok: false,
                error: Some("OFFICE_DOCUMENT_LOCKED".into()),
                backend: Some("desktop_com".into()),
                artefacts: vec![],
                warnings: vec!["retried once".into()],
            },
        ];
        let summary = summarize(&items);
        assert_eq!(summary.files, 2);
        assert_eq!(summary.succeeded, 1);
        assert_eq!(summary.failed, 1);
        assert_eq!(summary.backend.as_deref(), Some("desktop_com"));
        assert_eq!(summary.artefacts, vec!["a.pdf".to_string()]);
        assert!(summary
            .warnings
            .iter()
            .any(|w| w.contains("OFFICE_DOCUMENT_LOCKED")));
        assert_eq!(summary.phase(), OfficeJobPhase::PartiallySucceeded);
        let json = serde_json::to_string(&summary).unwrap();
        let back: BatchSummary = serde_json::from_str(&json).unwrap();
        assert_eq!(back, summary);
    }
}
