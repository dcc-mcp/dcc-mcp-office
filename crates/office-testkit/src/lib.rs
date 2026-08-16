//! dcc-mcp-office-testkit — contract-test helpers (proposal §24.1).
//!
//! Contract surface under test: RPC schema, capability manifests, IR schema,
//! error codes, events, protocol-version negotiation.
//!
//! M0 placeholder: real helpers land in M1 together with the pipe transport
//! (JSON-RPC fixtures, a fake sidecar for gateway tests, fault-injection
//! hooks for Busy / modal / Protected View).

#![forbid(unsafe_code)]

pub use dcc_mcp_office_protocol as protocol;

/// Contract-test marker: every error code the gateway may surface to agents
/// must round-trip through the protocol crate.
pub fn assert_error_code_round_trip(code: protocol::OfficeErrorCode) {
    let json = serde_json::to_string(&code).expect("serialize");
    let back: protocol::OfficeErrorCode = serde_json::from_str(&json).expect("deserialize");
    assert_eq!(back, code);
}

#[cfg(test)]
mod tests {
    use super::*;
    use dcc_mcp_office_protocol::OfficeErrorCode;

    #[test]
    fn sample_error_codes_round_trip() {
        assert_error_code_round_trip(OfficeErrorCode::OfficeModalDialog);
        assert_error_code_round_trip(OfficeErrorCode::OfficeGraphAuthRequired);
    }
}
