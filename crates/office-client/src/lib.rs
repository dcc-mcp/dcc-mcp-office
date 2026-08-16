//! dcc-mcp-office-client — Rust-side client for office-host.exe.
//!
//! M0 scaffold: URI scheme + configuration surface only. M1 wiring plan:
//!
//! * implement dcc_mcp_host_rpc::HostRpcClient from dcc-mcp-core (RFC #998
//!   Phase 2) so the existing dcc-mcp-server sidecar binary manages process
//!   lifecycle / PPID watch for the office host;
//! * register the namedpipe:// URI scheme in the scheme registry;
//! * speak office-rpc/1 JSON-RPC over the named pipe from
//!   dcc_mcp_office_protocol::pipe_name.
//!
//! This crate intentionally carries no I/O dependencies in M0.

#![forbid(unsafe_code)]

use dcc_mcp_office_protocol::SidecarState;

/// URI scheme under which this client registers with dcc-mcp-host-rpc.
pub const URI_SCHEME: &str = "namedpipe";

/// Default handshake timeout.
pub const HANDSHAKE_TIMEOUT_MS: u64 = 30_000;

/// Default namedpipe:// URI for an application sidecar:
/// namedpipe://powerpoint → pipe \\.\pipe\dcc-mcp-office-powerpoint-...
pub fn default_uri(app: &str) -> String {
    format!("{URI_SCHEME}://{app}")
}

/// Placeholder client handle. M1: connection + handshake + command/event
/// channels over the pipe.
pub struct OfficeHostClient {
    app: String,
    state: SidecarState,
}

impl OfficeHostClient {
    pub fn new(app: impl Into<String>) -> Self {
        Self {
            app: app.into(),
            state: SidecarState::Requested,
        }
    }

    pub fn app(&self) -> &str {
        &self.app
    }

    pub fn state(&self) -> SidecarState {
        self.state
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn uri_scheme_is_namedpipe() {
        assert_eq!(default_uri("powerpoint"), "namedpipe://powerpoint");
    }

    #[test]
    fn client_starts_requested() {
        let c = OfficeHostClient::new("word");
        assert_eq!(c.state(), SidecarState::Requested);
    }
}
