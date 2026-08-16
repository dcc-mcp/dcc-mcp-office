//! dcc-mcp-office-protocol — wire schema for the office-rpc/1 sidecar contract.
//!
//! This crate is **schema-only** (no I/O), mirroring the role of
//! dcc-mcp-host-rpc in dcc-mcp-core: everything is serde round-trippable and
//! versioned so the Rust gateway and the C# office-host.exe sidecar can
//! negotiate and evolve independently.
//!
//! Transport: JSON-RPC 2.0 over a Windows named pipe
//! \\.\pipe\dcc-mcp-office-{app}-{user_sid}-{session_id}
//! (see pipe_name). Large payloads travel as artefact IDs, never
//! Base64-inline (proposal §12.1).
//!
//! Reference: proposal §10-§21 in docs/proposals/office-automation-platform-v1.0.md.

#![forbid(unsafe_code)]

use serde::{Deserialize, Serialize};

/// Current wire protocol version (proposal §12.2).
pub const PROTOCOL_VERSION: &str = "office-rpc/1";

/// Named-pipe name prefix (proposal §12.1).
pub const PIPE_PREFIX: &str = "dcc-mcp-office";

/// Builds the canonical named-pipe name for an Office application sidecar.
///
/// ACL is per current-user SID; the gateway is the only remote entry point.
pub fn pipe_name(app: &str, user_sid: &str, session_id: u32) -> String {
    format!(r"\\.\pipe\{PIPE_PREFIX}-{app}-{user_sid}-{session_id}")
}

/// Sidecar lifecycle states (proposal §8.3).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum SidecarState {
    Requested,
    Launching,
    Handshaking,
    Attaching,
    CreatingApplication,
    Ready,
    Busy,
    Degraded,
    Recovering,
    Stopped,
}

/// Standard error codes (proposal §20). RPC error.code carries one of these
/// where applicable; error.message stays human-readable.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "SCREAMING_SNAKE_CASE")]
pub enum OfficeErrorCode {
    OfficeAppNotInstalled,
    OfficeAppVersionUnsupported,
    OfficeAppBusy,
    OfficeModalDialog,
    OfficeProtectedView,
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
}

impl OfficeErrorCode {
    /// Codes for which an automatic retry of the same call is safe
    /// (idempotent or determinable — proposal §20 recovery ladder).
    pub fn is_retryable(self) -> bool {
        matches!(
            self,
            OfficeErrorCode::OfficeAppBusy
                | OfficeErrorCode::OfficeRpcTimeout
                | OfficeErrorCode::OfficeGraphThrottled
                | OfficeErrorCode::OfficeBackendUnavailable
        )
    }
}

/// Application identity reported at handshake (proposal §10.3).
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct ApplicationInfo {
    pub name: String,
    pub version: String,
    pub bitness: String,
    pub language: String,
}

/// Sidecar execution limits (proposal §10.3).
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default)]
pub struct Limits {
    pub max_parallel_writes: u32,
    pub requires_user_session: bool,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_parallel_writes: 1,
            requires_user_session: true,
        }
    }
}

/// Capability manifest a sidecar reports at handshake and re-reports on
/// reconnect (proposal §10.3).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(default)]
pub struct CapabilityManifest {
    pub provider: String,
    pub provider_version: String,
    pub protocol_version: String,
    pub application: Option<ApplicationInfo>,
    pub execution_modes: Vec<String>,
    /// capability name → semantic version.
    pub capabilities: std::collections::BTreeMap<String, String>,
    pub limits: Limits,
}

impl Default for CapabilityManifest {
    fn default() -> Self {
        Self {
            provider: String::new(),
            provider_version: "0.1.0".to_string(),
            protocol_version: PROTOCOL_VERSION.to_string(),
            application: None,
            execution_modes: Vec::new(),
            capabilities: Default::default(),
            limits: Limits::default(),
        }
    }
}

/// Handshake request (proposal §12.2).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct HandshakeRequest {
    pub protocol_versions: Vec<String>,
    pub gateway_version: String,
    pub requested_app: String,
}

/// Handshake response (proposal §12.2).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct HandshakeResponse {
    pub protocol_version: String,
    pub host_id: String,
    #[serde(default)]
    pub capability_manifest: CapabilityManifest,
}

/// Stable reference to an open document plus optimistic-concurrency guard
/// (proposal §12.3 / §14.5).
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct DocumentRef {
    pub document_id: String,
    pub expected_revision: u64,
}

/// office.command.execute params (proposal §12.3). input and policy are
/// capability-specific and validated against the capability schema on both
/// sides.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CommandParams {
    pub capability: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub document: Option<DocumentRef>,
    #[serde(default = "default_input")]
    pub input: serde_json::Value,
    #[serde(default = "default_policy")]
    pub policy: serde_json::Value,
}

fn default_input() -> serde_json::Value {
    serde_json::Value::Object(Default::default())
}

fn default_policy() -> serde_json::Value {
    serde_json::json!({ "checkpoint": true, "render_after": false })
}

/// Command result. indeterminate means the sidecar cannot prove whether the
/// write took effect — the caller must re-inspect, never assume success or
/// failure (proposal §20).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CommandResult {
    pub operation_id: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub revision: Option<u64>,
    /// Change summary (what changed, per scope).
    #[serde(default = "default_input")]
    pub changed: serde_json::Value,
    #[serde(default)]
    pub warnings: Vec<String>,
    #[serde(default)]
    pub artefacts: Vec<ArtifactRecord>,
    #[serde(default)]
    pub validation: serde_json::Value,
    /// Execution path actually used (openxml | desktop_com | graph | office_js | uia | cua).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub backend: Option<String>,
    #[serde(default)]
    pub indeterminate: bool,
}

/// office.job.progress notification (proposal §12.4).
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct JobProgress {
    pub job_id: String,
    pub stage: String,
    pub completed: u64,
    pub total: u64,
}

/// Minimised selection DTO carried by selection_changed events
/// (proposal §12.4 / §21).
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct SelectionInfo {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub slide_id: Option<u64>,
    #[serde(default)]
    pub object_ids: Vec<String>,
}

/// Unified event notification (proposal §21). High-frequency events must be
/// debounced on the sidecar before this envelope crosses the pipe.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct EventNotification {
    pub event: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub document_id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub revision: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub selection: Option<SelectionInfo>,
    /// provider / app instance / correlation id (free-form DTO).
    #[serde(default)]
    pub context: serde_json::Value,
}

/// Sidecar heartbeat payload (proposal §8.3).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct SidecarStatus {
    pub state: SidecarState,
    pub app: String,
    pub pid: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub office_version: Option<String>,
    #[serde(default)]
    pub open_documents: Vec<String>,
    #[serde(default)]
    pub busy: bool,
    #[serde(default)]
    pub modal: bool,
    #[serde(default)]
    pub protected_view: bool,
}

/// Artifact registry record (proposal §17).
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct ArtifactRecord {
    pub artifact_id: String,
    pub kind: String,
    pub path: String,
    pub sha256: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub source_document_id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub revision: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub created_by_job: Option<String>,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn pipe_name_matches_proposal_format() {
        assert_eq!(
            pipe_name("powerpoint", "S-1-5-21-42", 3),
            r"\\.\pipe\dcc-mcp-office-powerpoint-S-1-5-21-42-3"
        );
    }

    #[test]
    fn error_codes_serialize_to_proposal_names() {
        let code = serde_json::to_string(&OfficeErrorCode::OfficeAppBusy).unwrap();
        assert_eq!(code, "\"OFFICE_APP_BUSY\"");
        let code = serde_json::to_string(&OfficeErrorCode::OfficeDocumentConflict).unwrap();
        assert_eq!(code, "\"OFFICE_DOCUMENT_CONFLICT\"");
        assert!(OfficeErrorCode::OfficeAppBusy.is_retryable());
        assert!(!OfficeErrorCode::OfficeDocumentConflict.is_retryable());
    }

    #[test]
    fn handshake_round_trip() {
        let req = HandshakeRequest {
            protocol_versions: vec![PROTOCOL_VERSION.to_string()],
            gateway_version: "1.0.0".to_string(),
            requested_app: "powerpoint".to_string(),
        };
        let json = serde_json::to_string(&req).unwrap();
        let back: HandshakeRequest = serde_json::from_str(&json).unwrap();
        assert_eq!(back.requested_app, "powerpoint");
        assert!(json.contains("office-rpc/1"));
    }

    #[test]
    fn capability_manifest_defaults() {
        let m: CapabilityManifest = serde_json::from_str("{}").unwrap();
        assert_eq!(m.limits.max_parallel_writes, 1);
        assert!(m.limits.requires_user_session);
    }

    #[test]
    fn command_params_default_policy_checkpoints() {
        let p: CommandParams = serde_json::from_str(
            r#"{"capability":"presentation.patch","input":{"operations":[]}}"#,
        )
        .unwrap();
        assert_eq!(p.policy["checkpoint"], true);
        assert_eq!(p.policy["render_after"], false);
    }

    #[test]
    fn command_result_marks_indeterminate() {
        let r = CommandResult {
            operation_id: "op-1".into(),
            revision: None,
            changed: serde_json::json!({}),
            warnings: vec![],
            artefacts: vec![],
            validation: serde_json::json!({}),
            backend: Some("desktop_com".into()),
            indeterminate: true,
        };
        let json = serde_json::to_string(&r).unwrap();
        assert!(json.contains("\"indeterminate\":true"));
    }
}
