//! dcc-mcp-office-testkit — contract-test helpers (proposal §24.1 / §24.4).
//!
//! M1 surface:
//!   - JSON-RPC 2.0 fixtures matching the office-rpc/1 wire shape,
//!   - FakeSidecar: an in-process scripted JSON-RPC responder so gateway
//!     code can be exercised without Office or a live office-host.exe,
//!   - fault-injection hooks for the §24.4 ladder: Busy (OFFICE_APP_BUSY),
//!     modal dialog (OFFICE_MODAL_DIALOG), timeout (no response at all),
//!     Protected View (OFFICE_PROTECTED_VIEW).
//!
//! The Rust ↔ C# end-to-end contract tests live in the office-client crate's
//! pipe_contract test; this crate is the dependency-light half for unit-level
//! gateway tests.

#![forbid(unsafe_code)]

use std::collections::HashMap;

use serde_json::Value;

pub use dcc_mcp_office_protocol as protocol;

/// Contract-test marker: every error code the gateway may surface to agents
/// must round-trip through the protocol crate.
pub fn assert_error_code_round_trip(code: protocol::OfficeErrorCode) {
    let json = serde_json::to_string(&code).expect("serialize");
    let back: protocol::OfficeErrorCode = serde_json::from_str(&json).expect("deserialize");
    assert_eq!(back, code);
}

/// Wire name of an error code — the exact string the host puts in
/// JSON-RPC error.code (e.g. "OFFICE_APP_BUSY").
pub fn wire_name(code: protocol::OfficeErrorCode) -> String {
    serde_json::to_string(&code)
        .expect("serialize")
        .trim_matches('"')
        .to_string()
}

pub mod rpc {
    //! JSON-RPC 2.0 fixture builders (proposal §12).
    use serde_json::{json, Value};

    pub fn request(id: u64, method: &str, params: Value) -> Value {
        json!({ "jsonrpc": "2.0", "id": id, "method": method, "params": params })
    }

    pub fn response(id: u64, result: Value) -> Value {
        json!({ "jsonrpc": "2.0", "id": id, "result": result })
    }

    pub fn error(id: u64, code: impl Into<Value>, message: &str) -> Value {
        json!({ "jsonrpc": "2.0", "id": id, "error": { "code": code.into(), "message": message } })
    }

    /// Error with an OFFICE_* wire code (the code field stays a string).
    pub fn office_error(id: u64, code: super::protocol::OfficeErrorCode, message: &str) -> Value {
        error(id, super::wire_name(code), message)
    }
}

/// Which fault the fake injects next (proposal §24.4).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum InjectedFault {
    Busy,
    ModalDialog,
    Timeout,
    ProtectedView,
}

/// Scripted in-process JSON-RPC sidecar.
///
/// Every method (or command capability) maps to a scripted result; fault
/// counters inject the §24.4 ladder deterministically. A timeout fault means
/// no response line at all — callers must treat it like an RPC timeout.
pub struct FakeSidecar {
    results: HashMap<String, Value>,
    capabilities: HashMap<String, Value>,
    faults: Vec<InjectedFault>,
    served: usize,
}

impl Default for FakeSidecar {
    fn default() -> Self {
        Self::new()
    }
}

impl FakeSidecar {
    pub fn new() -> Self {
        Self {
            results: HashMap::new(),
            capabilities: HashMap::new(),
            faults: Vec::new(),
            served: 0,
        }
    }

    /// Scripts a raw method result (e.g. office.host.handshake).
    pub fn on(mut self, method: &str, result: Value) -> Self {
        self.results.insert(method.to_string(), result);
        self
    }

    /// Scripts an office.command.execute result for one capability.
    pub fn on_capability(mut self, capability: &str, result: Value) -> Self {
        self.capabilities.insert(capability.to_string(), result);
        self
    }

    /// Appends faults served in order before the scripted responses kick in.
    pub fn inject(mut self, faults: impl IntoIterator<Item = InjectedFault>) -> Self {
        self.faults.extend(faults);
        self
    }

    /// Requests served so far (faults included).
    pub fn served(&self) -> usize {
        self.served
    }

    /// Handles one request line; None = timeout fault (no response).
    pub fn handle(&mut self, request: &str) -> Option<String> {
        let parsed: Value = serde_json::from_str(request.trim_end()).ok()?;
        let id = parsed.get("id").cloned().unwrap_or(Value::Null);
        let method = parsed
            .get("method")
            .and_then(Value::as_str)
            .unwrap_or("")
            .to_string();
        self.served += 1;

        if let Some(fault) = self.faults.first().copied() {
            self.faults.remove(0);
            let response = match fault {
                InjectedFault::Timeout => return None,
                InjectedFault::Busy => rpc::office_error(
                    id.as_u64().unwrap_or(0),
                    protocol::OfficeErrorCode::OfficeAppBusy,
                    "injected busy",
                ),
                InjectedFault::ModalDialog => rpc::office_error(
                    id.as_u64().unwrap_or(0),
                    protocol::OfficeErrorCode::OfficeModalDialog,
                    "injected modal dialog",
                ),
                InjectedFault::ProtectedView => rpc::office_error(
                    id.as_u64().unwrap_or(0),
                    protocol::OfficeErrorCode::OfficeProtectedView,
                    "injected protected view",
                ),
            };
            return Some(response.to_string());
        }

        let result = if method == "office.command.execute" {
            let capability = parsed
                .pointer("/params/capability")
                .and_then(Value::as_str)
                .unwrap_or("");
            match self.capabilities.get(capability) {
                Some(result) => result.clone(),
                None => {
                    return Some(
                        rpc::office_error(
                            id.as_u64().unwrap_or(0),
                            protocol::OfficeErrorCode::OfficeCapabilityUnsupported,
                            &format!("OFFICE_CAPABILITY_UNSUPPORTED: {capability}"),
                        )
                        .to_string(),
                    );
                }
            }
        } else {
            match self.results.get(&method) {
                Some(result) => result.clone(),
                None => {
                    return Some(
                        rpc::office_error(
                            id.as_u64().unwrap_or(0),
                            protocol::OfficeErrorCode::OfficeInvalidRequest,
                            &format!("unknown method: {method}"),
                        )
                        .to_string(),
                    );
                }
            }
        };

        Some(rpc::response(id.as_u64().unwrap_or(0), result).to_string())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use dcc_mcp_office_protocol::OfficeErrorCode;
    use serde_json::json;

    #[test]
    fn sample_error_codes_round_trip() {
        assert_error_code_round_trip(OfficeErrorCode::OfficeModalDialog);
        assert_error_code_round_trip(OfficeErrorCode::OfficeGraphAuthRequired);
    }

    #[test]
    fn wire_names_match_the_host_strings() {
        assert_eq!(wire_name(OfficeErrorCode::OfficeAppBusy), "OFFICE_APP_BUSY");
        assert_eq!(
            wire_name(OfficeErrorCode::OfficeInvalidRequest),
            "OFFICE_INVALID_REQUEST"
        );
        assert_eq!(
            wire_name(OfficeErrorCode::OfficeCapabilityUnsupported),
            "OFFICE_CAPABILITY_UNSUPPORTED"
        );
    }

    #[test]
    fn scripted_method_returns_result() {
        let mut sidecar = FakeSidecar::new().on(
            "office.host.ping",
            json!({ "app": "powerpoint", "protocol_version": "office-rpc/1" }),
        );
        let response = sidecar
            .handle(&rpc::request(7, "office.host.ping", json!({})).to_string())
            .expect("response");
        let parsed: Value = serde_json::from_str(&response).unwrap();
        assert_eq!(parsed["id"], 7);
        assert_eq!(parsed["result"]["app"], "powerpoint");
    }

    #[test]
    fn busy_fault_then_success() {
        let mut sidecar = FakeSidecar::new()
            .on_capability("batch.convert", json!({ "backend": "desktop_com" }))
            .inject([InjectedFault::Busy]);
        let first = sidecar
            .handle(
                &rpc::request(
                    1,
                    "office.command.execute",
                    json!({ "capability": "batch.convert", "input": {} }),
                )
                .to_string(),
            )
            .expect("response");
        let parsed: Value = serde_json::from_str(&first).unwrap();
        assert_eq!(parsed["error"]["code"], "OFFICE_APP_BUSY");

        let second = sidecar
            .handle(
                &rpc::request(
                    2,
                    "office.command.execute",
                    json!({ "capability": "batch.convert", "input": {} }),
                )
                .to_string(),
            )
            .expect("response");
        let parsed: Value = serde_json::from_str(&second).unwrap();
        assert_eq!(parsed["result"]["backend"], "desktop_com");
    }

    #[test]
    fn timeout_fault_produces_no_response() {
        let mut sidecar = FakeSidecar::new().inject([InjectedFault::Timeout]);
        assert!(sidecar
            .handle(&rpc::request(1, "office.host.ping", json!({})).to_string())
            .is_none());
        assert_eq!(sidecar.served(), 1);
    }

    #[test]
    fn unknown_capability_carries_wire_code() {
        let mut sidecar = FakeSidecar::new();
        let response = sidecar
            .handle(
                &rpc::request(
                    3,
                    "office.command.execute",
                    json!({ "capability": "nope", "input": {} }),
                )
                .to_string(),
            )
            .expect("response");
        let parsed: Value = serde_json::from_str(&response).unwrap();
        assert_eq!(parsed["error"]["code"], "OFFICE_CAPABILITY_UNSUPPORTED");
    }
}
