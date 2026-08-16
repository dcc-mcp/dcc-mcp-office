//! dcc-mcp-office-security — default-deny policy (proposal §19).
//!
//! Enforced in **two layers**: the Rust gateway checks policy before
//! dispatching; the C# host re-checks at the COM boundary and forces
//! AutomationSecurity to disable macros while opening untrusted files
//! (MS-23/24/27/28/29). XLM/Excel 4.0 macros need separate detection because
//! msoAutomationSecurityForceDisable does not cover them (MS-27).

#![forbid(unsafe_code)]

use std::collections::BTreeMap;

use serde::{Deserialize, Serialize};

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
        Self {
            vba_application_run: PolicyAction::Deny,
            macros: PolicyAction::Deny,
            external_links_auto_update: PolicyAction::Deny,
            ole_activex_activation: PolicyAction::Deny,
            protected_view_bypass: PolicyAction::Deny,
            arbitrary_execute_mso: PolicyAction::Deny,
            print: PolicyAction::Confirm,
            overwrite_original: PolicyAction::CheckpointAndConfirm,
            send_email: PolicyAction::Confirm,
            meeting_invite: PolicyAction::Confirm,
            access_macros: PolicyAction::DenyOrConfirm,
            project_publish: PolicyAction::Confirm,
            workspace_only: true,
            execute_mso_allowlist: BTreeMap::new(),
            execute_mso_confirm: vec!["PrintPreviewAndPrint".to_string()],
        }
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
}
