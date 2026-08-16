//! dcc-mcp-office-client — Rust-side client for office-host.exe.
//!
//! M1 wiring: the client speaks office-rpc/1 JSON-RPC over the Windows named
//! pipe from dcc_mcp_office_protocol::pipe_name. The transport is
//! std-only ("\\.\pipe\..." opens through "std::fs::OpenOptions", line-framed
//! JSON), keeping this crate dependency-free except serde. On non-Windows
//! hosts the client compiles to an explicit ClientError::UnsupportedPlatform.
//!
//! Remaining gateway integration (not this crate): registering the
//! "namedpipe://" URI scheme with dcc-mcp-host-rpc so the dcc-mcp-server
//! sidecar binary owns process lifecycle / PPID watch for office-host.

#![forbid(unsafe_code)]

use std::fmt;

use dcc_mcp_office_protocol::{
    CommandParams, CommandResult, HandshakeResponse, SidecarState, PROTOCOL_VERSION,
};
use serde_json::{json, Value};

/// URI scheme under which this client registers with dcc-mcp-host-rpc.
pub const URI_SCHEME: &str = "namedpipe";

/// Default handshake timeout.
pub const HANDSHAKE_TIMEOUT_MS: u64 = 30_000;

/// Default namedpipe:// URI for an application sidecar:
/// namedpipe://powerpoint → pipe \\.\pipe\dcc-mcp-office-powerpoint-...
pub fn default_uri(app: &str) -> String {
    format!("{URI_SCHEME}://{app}")
}

/// Client-side failures (proposal §20 error ladder, client half).
#[derive(Debug)]
pub enum ClientError {
    /// connect() on a non-Windows host.
    UnsupportedPlatform,
    Io(std::io::Error),
    Serde(serde_json::Error),
    /// JSON-RPC error response; code is the OFFICE_* wire string.
    Rpc {
        code: Value,
        message: String,
    },
    /// Protocol violations: version mismatch, missing response, bad state.
    Protocol(String),
    /// The host closed the pipe.
    Disconnected,
}

impl fmt::Display for ClientError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            ClientError::UnsupportedPlatform => write!(f, "named pipes require Windows"),
            ClientError::Io(e) => write!(f, "pipe I/O error: {e}"),
            ClientError::Serde(e) => write!(f, "serialization error: {e}"),
            ClientError::Rpc { code, message } => write!(f, "rpc error {code}: {message}"),
            ClientError::Protocol(m) => write!(f, "protocol error: {m}"),
            ClientError::Disconnected => write!(f, "host closed the pipe"),
        }
    }
}

impl std::error::Error for ClientError {}

impl From<std::io::Error> for ClientError {
    fn from(e: std::io::Error) -> Self {
        ClientError::Io(e)
    }
}

impl From<serde_json::Error> for ClientError {
    fn from(e: serde_json::Error) -> Self {
        ClientError::Serde(e)
    }
}

#[cfg(windows)]
mod transport {
    use std::fs::{File, OpenOptions};
    use std::io::{BufRead, BufReader, Write};

    use serde_json::{json, Value};

    use super::ClientError;

    /// Line-framed JSON-RPC 2.0 over one duplex named-pipe handle.
    /// std::fs opens \\.\pipe\... natively on Windows — no extra crates.
    pub struct PipeRpc {
        reader: BufReader<File>,
        writer: File,
        next_id: u64,
    }

    impl PipeRpc {
        pub fn connect(pipe: &str) -> Result<Self, ClientError> {
            let file = OpenOptions::new().read(true).write(true).open(pipe)?;
            let reader = BufReader::new(file.try_clone()?);
            Ok(Self {
                reader,
                writer: file,
                next_id: 0,
            })
        }

        pub fn call(&mut self, method: &str, params: Value) -> Result<Value, ClientError> {
            let id = self.next_id;
            self.next_id += 1;
            let request = json!({ "jsonrpc": "2.0", "id": id, "method": method, "params": params });
            self.writer.write_all(request.to_string().as_bytes())?;
            self.writer.write_all(b"\n")?;
            self.writer.flush()?;
            loop {
                let mut line = String::new();
                let read = self.reader.read_line(&mut line)?;
                if read == 0 {
                    return Err(ClientError::Disconnected);
                }
                let Ok(message) = serde_json::from_str::<Value>(line.trim_end()) else {
                    continue; // tolerate non-JSON noise lines
                };
                // Notifications (no id) are skipped; responses carry our id.
                match message.get("id").and_then(Value::as_u64) {
                    Some(response_id) if response_id == id => {
                        if let Some(error) = message.get("error") {
                            return Err(ClientError::Rpc {
                                code: error.get("code").cloned().unwrap_or(Value::Null),
                                message: error
                                    .get("message")
                                    .and_then(Value::as_str)
                                    .unwrap_or("rpc error")
                                    .to_string(),
                            });
                        }
                        return Ok(message.get("result").cloned().unwrap_or(Value::Null));
                    }
                    _ => continue,
                }
            }
        }
    }
}

/// Client handle: connection + handshake + command execution over the pipe.
pub struct OfficeHostClient {
    app: String,
    state: SidecarState,
    #[cfg(windows)]
    pipe: Option<transport::PipeRpc>,
}

impl OfficeHostClient {
    pub fn new(app: impl Into<String>) -> Self {
        Self {
            app: app.into(),
            state: SidecarState::Requested,
            #[cfg(windows)]
            pipe: None,
        }
    }

    pub fn app(&self) -> &str {
        &self.app
    }

    pub fn state(&self) -> SidecarState {
        self.state
    }

    /// Canonical pipe name for this client's app (matches the host ACL layout).
    pub fn default_pipe_name(&self, user_sid: &str, session_id: u32) -> String {
        dcc_mcp_office_protocol::pipe_name(&self.app, user_sid, session_id)
    }

    /// Opens the named pipe. On non-Windows hosts this fails with
    /// ClientError::UnsupportedPlatform instead of pretending to connect.
    #[cfg(windows)]
    pub fn connect(&mut self, pipe: &str) -> Result<&mut Self, ClientError> {
        self.pipe = Some(transport::PipeRpc::connect(pipe)?);
        self.state = SidecarState::Handshaking;
        Ok(self)
    }

    #[cfg(not(windows))]
    pub fn connect(&mut self, _pipe: &str) -> Result<&mut Self, ClientError> {
        Err(ClientError::UnsupportedPlatform)
    }

    /// office.host.handshake → protocol-version check + capability manifest.
    #[cfg(windows)]
    pub fn handshake(&mut self, gateway_version: &str) -> Result<HandshakeResponse, ClientError> {
        let response = self.call(
            "office.host.handshake",
            json!({
                "protocol_versions": [PROTOCOL_VERSION],
                "gateway_version": gateway_version,
                "requested_app": self.app,
            }),
        )?;
        let handshake: HandshakeResponse = serde_json::from_value(response)?;
        if handshake.protocol_version != PROTOCOL_VERSION {
            return Err(ClientError::Protocol(format!(
                "host speaks '{}', client requires '{PROTOCOL_VERSION}'",
                handshake.protocol_version
            )));
        }
        self.state = SidecarState::Ready;
        Ok(handshake)
    }

    #[cfg(not(windows))]
    pub fn handshake(&mut self, _gateway_version: &str) -> Result<HandshakeResponse, ClientError> {
        Err(ClientError::UnsupportedPlatform)
    }

    /// office.host.ping — liveness plus attach state.
    #[cfg(windows)]
    pub fn ping(&mut self) -> Result<Value, ClientError> {
        self.call("office.host.ping", json!({}))
    }

    #[cfg(not(windows))]
    pub fn ping(&mut self) -> Result<Value, ClientError> {
        Err(ClientError::UnsupportedPlatform)
    }

    /// office.command.execute — the task-level capability surface (§12.3).
    #[cfg(windows)]
    pub fn execute(&mut self, params: &CommandParams) -> Result<CommandResult, ClientError> {
        let response = self.call("office.command.execute", serde_json::to_value(params)?)?;
        Ok(serde_json::from_value(response)?)
    }

    #[cfg(not(windows))]
    pub fn execute(&mut self, _params: &CommandParams) -> Result<CommandResult, ClientError> {
        Err(ClientError::UnsupportedPlatform)
    }

    #[cfg(windows)]
    fn call(&mut self, method: &str, params: Value) -> Result<Value, ClientError> {
        let pipe = self
            .pipe
            .as_mut()
            .ok_or_else(|| ClientError::Protocol("connect() must precede calls".into()))?;
        pipe.call(method, params)
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

    #[test]
    fn default_pipe_name_uses_protocol_layout() {
        let c = OfficeHostClient::new("powerpoint");
        assert_eq!(
            c.default_pipe_name("S-1-5-21-42", 3),
            r"\\.\pipe\dcc-mcp-office-powerpoint-S-1-5-21-42-3"
        );
    }
}
