//! dcc-mcp-office-security — default-deny policy (proposal §19).
//!
//! Enforced in **two layers**: the Rust gateway checks policy before
//! dispatching; the C# host re-checks at the COM boundary and forces
//! AutomationSecurity to disable macros while opening untrusted files
//! (MS-23/24/27/28/29). XLM/Excel 4.0 macros need separate detection because
//! msoAutomationSecurityForceDisable does not cover them (MS-27).

#![forbid(unsafe_code)]

use std::collections::{BTreeMap, BTreeSet};

use serde::{Deserialize, Serialize};
use serde_json::Value;

use dcc_mcp_office_protocol::OfficeErrorCode;

/// Policy action for a capability class.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum PolicyAction {
    Deny,
    Confirm,
    CheckpointAndConfirm,
    DenyOrConfirm,
}

impl PolicyAction {
    /// Whether a human confirmation is required before execution.
    pub fn requires_confirmation(self) -> bool {
        matches!(
            self,
            PolicyAction::Confirm
                | PolicyAction::CheckpointAndConfirm
                | PolicyAction::DenyOrConfirm
        )
    }
}

/// Security policy. Defaults match proposal §19.1 exactly.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default)]
pub struct SecurityPolicy {
    pub vba_application_run: PolicyAction,
    pub macros: PolicyAction,
    pub external_links_auto_update: PolicyAction,
    pub ole_activex_activation: PolicyAction,
    pub protected_view_bypass: PolicyAction,
    pub arbitrary_execute_mso: PolicyAction,
    pub print: PolicyAction,
    pub overwrite_original: PolicyAction,
    pub send_email: PolicyAction,
    pub meeting_invite: PolicyAction,
    pub access_macros: PolicyAction,
    pub project_publish: PolicyAction,
    /// Restrict file access to the workspace (proposal §19.3).
    pub workspace_only: bool,
    /// ExecuteMso whitelist per application (proposal §19.2). Empty = deny all.
    pub execute_mso_allowlist: BTreeMap<String, Vec<String>>,
    /// ExecuteMso commands that additionally require confirmation.
    pub execute_mso_confirm: Vec<String>,
}

impl Default for SecurityPolicy {
    fn default() -> Self {
        let canonical = &dcc_mcp_office_protocol::capability_catalog().security_policy;
        Self {
            vba_application_run: canonical_action(canonical, "vba_application_run"),
            macros: canonical_action(canonical, "macros"),
            external_links_auto_update: canonical_action(canonical, "external_links_auto_update"),
            ole_activex_activation: canonical_action(canonical, "ole_activex_activation"),
            protected_view_bypass: canonical_action(canonical, "protected_view_bypass"),
            arbitrary_execute_mso: canonical_action(canonical, "arbitrary_execute_mso"),
            print: canonical_action(canonical, "print"),
            overwrite_original: canonical_action(canonical, "overwrite_original"),
            send_email: canonical_action(canonical, "send_email"),
            meeting_invite: canonical_action(canonical, "meeting_invite"),
            access_macros: canonical_action(canonical, "access_macros"),
            project_publish: canonical_action(canonical, "project_publish"),
            workspace_only: canonical.workspace_only,
            execute_mso_allowlist: canonical.execute_mso_allowlist.clone(),
            execute_mso_confirm: canonical.execute_mso_confirm.clone(),
        }
    }
}

fn canonical_action(
    policy: &dcc_mcp_office_protocol::CatalogSecurityPolicy,
    name: &str,
) -> PolicyAction {
    match policy.actions.get(name).map(String::as_str) {
        Some("deny") => PolicyAction::Deny,
        Some("confirm") => PolicyAction::Confirm,
        Some("checkpoint_and_confirm") => PolicyAction::CheckpointAndConfirm,
        Some("deny_or_confirm") => PolicyAction::DenyOrConfirm,
        value => panic!("canonical security policy action {name} is invalid: {value:?}"),
    }
}

impl SecurityPolicy {
    /// Whether ExecuteMso is allowed for an app; Some(true) means a
    /// confirmation is additionally required (proposal §19.2).
    pub fn execute_mso(&self, app: &str, command: &str) -> Option<bool> {
        let allowed = self
            .execute_mso_allowlist
            .get(app)
            .is_some_and(|list| list.iter().any(|c| c == command));
        if !allowed {
            return None;
        }
        Some(self.execute_mso_confirm.iter().any(|c| c == command))
    }
}

/// First-layer policy rejection returned before a sidecar is contacted.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PolicyViolation {
    /// Closed Office wire error associated with the rejected policy field.
    pub code: OfficeErrorCode,
    /// Human-readable explanation of the attempted relaxation.
    pub message: String,
}

/// Enforces the catalog policy at the MCP boundary. Requests may retain a
/// canonical action or tighten it to `deny`; all other divergence fails closed.
pub fn validate_policy_tightening(requested: &Value) -> Result<(), PolicyViolation> {
    let requested = requested
        .as_object()
        .ok_or_else(|| invalid("policy must be an object"))?;
    let canonical = dcc_mcp_office_protocol::capability_catalog();
    let mut known: BTreeSet<&str> = canonical
        .security_policy
        .actions
        .keys()
        .map(String::as_str)
        .collect();
    known.extend([
        "workspace_only",
        "workspace_root",
        "execute_mso_allowlist",
        "execute_mso_confirm",
        "checkpoint",
        "render_after",
    ]);
    if let Some(name) = requested.keys().find(|name| !known.contains(name.as_str())) {
        return Err(invalid(format!(
            "policy.{name} is not defined by the catalog"
        )));
    }

    for (name, canonical_action) in &canonical.security_policy.actions {
        let Some(value) = requested.get(name) else {
            continue;
        };
        let action = value
            .as_str()
            .ok_or_else(|| invalid(format!("policy.{name} must be a string")))?;
        if action != canonical_action && action != "deny" {
            return Err(PolicyViolation {
                code: policy_error(name),
                message: format!(
                    "policy.{name} cannot relax canonical action '{canonical_action}' to '{action}'"
                ),
            });
        }
    }

    if requested
        .get("workspace_only")
        .is_some_and(|value| value != &Value::Bool(true))
    {
        return Err(PolicyViolation {
            code: OfficeErrorCode::OfficeAccessDenied,
            message: "policy.workspace_only must remain true".into(),
        });
    }
    if requested
        .get("checkpoint")
        .is_some_and(|value| value != &Value::Bool(true))
    {
        return Err(PolicyViolation {
            code: OfficeErrorCode::OfficeCapabilityUnsupported,
            message: "policy.checkpoint must remain true".into(),
        });
    }
    if requested.get("execute_mso_allowlist").is_some_and(|value| {
        value
            .as_object()
            .is_none_or(|allowlist| !allowlist.is_empty())
    }) {
        return Err(PolicyViolation {
            code: OfficeErrorCode::OfficeCapabilityUnsupported,
            message: "policy.execute_mso_allowlist must stay empty".into(),
        });
    }
    if let Some(confirm) = requested.get("execute_mso_confirm") {
        let expected = serde_json::to_value(&canonical.security_policy.execute_mso_confirm)
            .expect("canonical confirmation list must serialize");
        if confirm != &expected {
            return Err(PolicyViolation {
                code: OfficeErrorCode::OfficeCapabilityUnsupported,
                message: "policy.execute_mso_confirm cannot diverge from the catalog".into(),
            });
        }
    }
    for name in ["render_after"] {
        if requested.get(name).is_some_and(|value| !value.is_boolean()) {
            return Err(invalid(format!("policy.{name} must be a boolean")));
        }
    }
    if requested
        .get("workspace_root")
        .is_some_and(|value| value.as_str().is_none_or(str::is_empty))
    {
        return Err(invalid("policy.workspace_root must be a non-empty string"));
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> PolicyViolation {
    PolicyViolation {
        code: OfficeErrorCode::OfficeInvalidRequest,
        message: message.into(),
    }
}

fn policy_error(name: &str) -> OfficeErrorCode {
    match name {
        "vba_application_run" | "macros" | "ole_activex_activation" | "access_macros" => {
            OfficeErrorCode::OfficeMacroBlocked
        }
        "external_links_auto_update" => OfficeErrorCode::OfficeExternalLinkBlocked,
        "protected_view_bypass" => OfficeErrorCode::OfficeProtectedView,
        _ => OfficeErrorCode::OfficeCapabilityUnsupported,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn defaults_deny_vba_macros_and_links() {
        let p = SecurityPolicy::default();
        assert_eq!(p.vba_application_run, PolicyAction::Deny);
        assert_eq!(p.macros, PolicyAction::Deny);
        assert_eq!(p.external_links_auto_update, PolicyAction::Deny);
        assert_eq!(p.arbitrary_execute_mso, PolicyAction::Deny);
        assert!(p.workspace_only);
        assert_eq!(
            dcc_mcp_office_protocol::capability_catalog()
                .security_policy
                .actions["macros"],
            "deny"
        );
    }

    #[test]
    fn defaults_confirm_print_and_send() {
        let p = SecurityPolicy::default();
        assert!(p.print.requires_confirmation());
        assert!(p.send_email.requires_confirmation());
        assert!(!p.vba_application_run.requires_confirmation());
    }

    #[test]
    fn execute_mso_is_deny_by_default() {
        let p = SecurityPolicy::default();
        assert_eq!(p.execute_mso("powerpoint", "Copy"), None);
    }

    #[test]
    fn execute_mso_whitelist_with_confirmation() {
        let mut p = SecurityPolicy::default();
        p.execute_mso_allowlist.insert(
            "powerpoint".into(),
            vec!["Copy".into(), "PrintPreviewAndPrint".into()],
        );
        assert_eq!(p.execute_mso("powerpoint", "Copy"), Some(false));
        assert_eq!(
            p.execute_mso("powerpoint", "PrintPreviewAndPrint"),
            Some(true)
        );
        assert_eq!(p.execute_mso("powerpoint", "PasteAsPicture"), None);
        assert_eq!(p.execute_mso("excel", "Copy"), None);
    }

    #[test]
    fn first_layer_allows_tightening_but_rejects_relaxation() {
        assert!(validate_policy_tightening(&serde_json::json!({
            "print": "deny",
            "workspace_only": true,
            "checkpoint": true
        }))
        .is_ok());

        let violation = validate_policy_tightening(&serde_json::json!({
            "macros": "confirm"
        }))
        .expect_err("macro policy relaxation must fail");
        assert_eq!(violation.code, OfficeErrorCode::OfficeMacroBlocked);
    }
}
