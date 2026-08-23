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

include!(concat!(env!("OUT_DIR"), "/office_error_codes.rs"));

/// Machine-readable capability catalog shared by Rust and the C# host.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct OfficeCapabilityCatalog {
    pub schema_version: String,
    pub protocol_version: String,
    pub provider: String,
    pub command_params_schema: String,
    pub security_policy: CatalogSecurityPolicy,
    pub errors: Vec<CatalogError>,
    pub capabilities: Vec<CatalogCapability>,
}

/// Canonical default-deny policy shared by the Rust gateway and C# host.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct CatalogSecurityPolicy {
    pub actions: std::collections::BTreeMap<String, String>,
    pub workspace_only: bool,
    pub execute_mso_allowlist: std::collections::BTreeMap<String, Vec<String>>,
    pub execute_mso_confirm: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct CatalogError {
    pub code: String,
    pub retryable: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct CatalogCapability {
    pub name: String,
    pub version: String,
    pub handler: String,
    pub mcp_tool: String,
    pub input_schema: String,
    pub output_schema: String,
    pub availability: Vec<CatalogAvailability>,
    pub errors: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct CatalogAvailability {
    pub execution_mode: String,
    pub apps: Vec<String>,
}

/// Parsed canonical catalog. Parsing is cached and cannot drift from the
/// file embedded in this crate at build time.
pub fn capability_catalog() -> &'static OfficeCapabilityCatalog {
    static CATALOG: std::sync::OnceLock<OfficeCapabilityCatalog> = std::sync::OnceLock::new();
    CATALOG.get_or_init(|| {
        serde_json::from_str(include_str!("../office-rpc.catalog.json"))
            .expect("embedded office-rpc capability catalog must be valid")
    })
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
    /// Evidence that a human approved one policy action. The host validates
    /// the action and proof before any confirm-gated operation starts.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub confirmation: Option<ConfirmationProof>,
    #[serde(default = "default_input")]
    pub input: serde_json::Value,
    #[serde(default = "default_policy")]
    pub policy: serde_json::Value,
}

/// Structured, auditable human-confirmation evidence (proposal §19.1).
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct ConfirmationProof {
    pub action: String,
    pub confirmed: bool,
    pub confirmed_by: String,
    pub confirmed_at: String,
}

fn default_input() -> serde_json::Value {
    serde_json::Value::Object(Default::default())
}

fn default_policy() -> serde_json::Value {
    let canonical = &capability_catalog().security_policy;
    let mut policy = serde_json::Map::new();
    for (name, action) in &canonical.actions {
        policy.insert(name.clone(), serde_json::Value::String(action.clone()));
    }
    policy.insert(
        "workspace_only".into(),
        serde_json::Value::Bool(canonical.workspace_only),
    );
    policy.insert(
        "execute_mso_allowlist".into(),
        serde_json::to_value(&canonical.execute_mso_allowlist)
            .expect("canonical ExecuteMso allowlist must serialize"),
    );
    policy.insert(
        "execute_mso_confirm".into(),
        serde_json::to_value(&canonical.execute_mso_confirm)
            .expect("canonical ExecuteMso confirmations must serialize"),
    );
    policy.insert(
        "checkpoint".into(),
        serde_json::Value::Bool(
            canonical
                .actions
                .get("overwrite_original")
                .map(String::as_str)
                == Some("checkpoint_and_confirm"),
        ),
    );
    policy.insert("render_after".into(), serde_json::Value::Bool(false));
    serde_json::Value::Object(policy)
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
    /// Audit trail (proposal §20 / §27 criterion 10): policy applied,
    /// security posture, backend, application info, duration — the host
    /// fills this on every command result.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub audit: Option<serde_json::Value>,
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

/// batch.convert input (proposal §15.1). File specs may be plain paths or
/// simple wildcards (* and ?, with **/ for recursion) expanded by the host.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default)]
pub struct BatchConvertParams {
    pub inputs: Vec<String>,
    /// pdf is the only target in M1.
    pub target_format: String,
    /// auto | desktop_com | openxml | graph. The desktop sidecar rejects
    /// non-desktop renderers; cloud routing is a gateway concern.
    pub backend: String,
    pub output_directory: String,
    /// versioned | fail | overwrite.
    pub overwrite: String,
    pub validation: Vec<String>,
}

impl Default for BatchConvertParams {
    fn default() -> Self {
        Self {
            inputs: Vec::new(),
            target_format: "pdf".to_string(),
            backend: "auto".to_string(),
            output_directory: String::new(),
            overwrite: "versioned".to_string(),
            validation: vec![
                "output_openable".to_string(),
                "non_empty".to_string(),
                "page_count_reasonable".to_string(),
            ],
        }
    }
}

/// One replace rule (proposal §15.2). Match modes: literal | case_insensitive.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct ReplaceRule {
    pub find: String,
    pub replace: String,
    #[serde(default = "default_match_mode")]
    pub r#match: String,
}

fn default_match_mode() -> String {
    "literal".to_string()
}

/// batch.replace_text input (proposal §15.2). dry_run defaults to true: a
/// replace without an explicit commit only reports matches.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default)]
pub struct BatchReplaceTextParams {
    pub inputs: Vec<String>,
    pub rules: Vec<ReplaceRule>,
    /// body | headers | footers | notes | comments | charts.
    pub scope: Vec<String>,
    pub dry_run: bool,
}

impl Default for BatchReplaceTextParams {
    fn default() -> Self {
        Self {
            inputs: Vec::new(),
            rules: Vec::new(),
            scope: vec!["body".to_string()],
            dry_run: true,
        }
    }
}

/// slide.render input (PowerPoint only — proposal §27 criterion 6).
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default)]
pub struct SlideRenderParams {
    pub path: String,
    pub output_directory: String,
    pub width: u32,
    pub height: u32,
}

impl Default for SlideRenderParams {
    fn default() -> Self {
        Self {
            path: String::new(),
            output_directory: String::new(),
            width: 1280,
            height: 720,
        }
    }
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
    fn unclassified_com_errors_are_terminal() {
        let code: OfficeErrorCode = serde_json::from_str("\"OFFICE_UNCLASSIFIED\"").unwrap();

        assert_eq!(code, OfficeErrorCode::OfficeUnclassified);
        assert!(!code.is_retryable());
        assert!(!OfficeErrorCode::OfficeAccessDenied.is_retryable());
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
        assert_eq!(p.policy["macros"], "deny");
        assert_eq!(p.policy["overwrite_original"], "checkpoint_and_confirm");
        assert_eq!(p.policy["workspace_only"], true);
    }

    #[test]
    fn command_params_carry_structured_confirmation_proof() {
        let p: CommandParams = serde_json::from_str(
            r#"{
                "capability":"batch.replace_text",
                "confirmation":{
                    "action":"overwrite_original",
                    "confirmed":true,
                    "confirmed_by":"human:reviewer",
                    "confirmed_at":"2026-08-23T14:00:00Z"
                }
            }"#,
        )
        .unwrap();

        let confirmation = p.confirmation.expect("confirmation proof");
        assert_eq!(confirmation.action, "overwrite_original");
        assert!(confirmation.confirmed);
        assert_eq!(confirmation.confirmed_by, "human:reviewer");
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
            audit: None,
        };
        let json = serde_json::to_string(&r).unwrap();
        assert!(json.contains("\"indeterminate\":true"));
    }

    #[test]
    fn batch_convert_params_default_to_pdf_with_validation() {
        let p: BatchConvertParams =
            serde_json::from_str(r#"{"inputs":["a.pptx"],"output_directory":"out"}"#).unwrap();
        assert_eq!(p.target_format, "pdf");
        assert_eq!(p.overwrite, "versioned");
        assert_eq!(p.validation.len(), 3);
    }

    #[test]
    fn batch_replace_text_defaults_to_dry_run() {
        let p: BatchReplaceTextParams = serde_json::from_str(
            r#"{"inputs":["a.docx"],"rules":[{"find":"2025年度","replace":"2026年度"}]}"#,
        )
        .unwrap();
        assert!(p.dry_run);
        assert_eq!(p.rules[0].r#match, "literal");
        assert_eq!(p.scope, vec!["body".to_string()]);
    }

    #[test]
    fn replace_rule_round_trip() {
        let rule = ReplaceRule {
            find: "Old Project Name".into(),
            replace: "DCC-MCP".into(),
            r#match: "case_insensitive".into(),
        };
        let json = serde_json::to_string(&rule).unwrap();
        assert!(json.contains("case_insensitive"));
        let back: ReplaceRule = serde_json::from_str(&json).unwrap();
        assert_eq!(back, rule);
    }

    #[test]
    fn slide_render_defaults_to_720p() {
        let p: SlideRenderParams =
            serde_json::from_str(r#"{"path":"a.pptx","output_directory":"out"}"#).unwrap();
        assert_eq!((p.width, p.height), (1280, 720));
    }

    #[test]
    fn command_result_carries_audit_trail() {
        let r = CommandResult {
            operation_id: "op-2".into(),
            revision: None,
            changed: serde_json::json!({}),
            warnings: vec![],
            artefacts: vec![],
            validation: serde_json::json!({}),
            backend: Some("desktop_com".into()),
            indeterminate: false,
            audit: Some(serde_json::json!({
                "security": {
                    "automation_security": {
                        "applicable": true,
                        "observed": 3,
                        "expected": 3,
                        "enforced": true
                    }
                },
                "duration_ms": 42,
            })),
        };
        let json = serde_json::to_string(&r).unwrap();
        assert!(json.contains("\"enforced\":true"));
        assert!(json.contains("duration_ms"));
        let back: CommandResult = serde_json::from_str(&json).unwrap();
        assert_eq!(back.audit.unwrap()["duration_ms"], 42);
    }
}
