//! dcc-mcp-office-client — Rust-side client for office-host.exe.
//!
//! M1 wiring: the client speaks office-rpc/1 JSON-RPC over the Windows named
//! pipe from dcc_mcp_office_protocol::pipe_name. Windows uses Tokio's
//! overlapped named-pipe I/O so writes and reads share one enforceable
//! deadline; non-Windows hosts compile to an explicit
//! ClientError::UnsupportedPlatform.
//!
//! Remaining gateway integration (not this crate): registering the
//! "namedpipe://" URI scheme with dcc-mcp-host-rpc so the dcc-mcp-server
//! sidecar binary owns process lifecycle / PPID watch for office-host.

#![forbid(unsafe_code)]

use std::collections::VecDeque;
use std::fmt;
use std::time::Duration;

use dcc_mcp_office_protocol::{
    CommandParams, CommandResult, HandshakeResponse, JobCancelResult, JobStatus, SidecarState,
    SidecarStatus,
};
use serde_json::Value;

// The pipe transport (and everything using these two) is Windows-only:
// non-Windows builds must stay warning-free for the ubuntu CI clippy gate.
#[cfg(windows)]
use dcc_mcp_office_protocol::PROTOCOL_VERSION;
#[cfg(windows)]
use serde_json::json;

/// URI scheme under which this client registers with dcc-mcp-host-rpc.
pub const URI_SCHEME: &str = "namedpipe";

/// Default handshake timeout.
pub const HANDSHAKE_TIMEOUT_MS: u64 = 30_000;

/// Default command timeout: longer than the host's 120-second COM soft timeout.
pub const DEFAULT_CALL_TIMEOUT_MS: u64 = 150_000;

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
        /// Structured JSON-RPC error data, including indeterminate write state.
        data: Value,
    },
    /// Protocol violations: version mismatch, missing response, bad state.
    Protocol(String),
    /// The host closed the pipe.
    Disconnected,
    /// A bounded connect or RPC operation exceeded its deadline.
    Timeout {
        operation: String,
        timeout: Duration,
    },
}

impl fmt::Display for ClientError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            ClientError::UnsupportedPlatform => write!(f, "named pipes require Windows"),
            ClientError::Io(e) => write!(f, "pipe I/O error: {e}"),
            ClientError::Serde(e) => write!(f, "serialization error: {e}"),
            ClientError::Rpc { code, message, .. } => write!(f, "rpc error {code}: {message}"),
            ClientError::Protocol(m) => write!(f, "protocol error: {m}"),
            ClientError::Disconnected => write!(f, "host closed the pipe"),
            ClientError::Timeout { operation, timeout } => {
                write!(f, "{operation} timed out after {timeout:?}")
            }
        }
    }
}

impl std::error::Error for ClientError {}

#[cfg(any(windows, test))]
impl ClientError {
    fn from_rpc_error(error: &Value) -> Self {
        Self::Rpc {
            code: error.get("code").cloned().unwrap_or(Value::Null),
            message: error
                .get("message")
                .and_then(Value::as_str)
                .unwrap_or("rpc error")
                .to_string(),
            data: error.get("data").cloned().unwrap_or(Value::Null),
        }
    }
}

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
    use std::collections::VecDeque;
    use std::future::Future;
    use std::time::Duration;

    use serde_json::{json, Value};
    use tokio::io::{AsyncBufReadExt, AsyncWriteExt, BufReader};
    use tokio::net::windows::named_pipe::{ClientOptions, NamedPipeClient};
    use tokio::runtime::{Builder, Runtime};

    use super::ClientError;

    const MAX_UNMATCHED_MESSAGES: usize = 8;

    fn block_on_compatible<F>(runtime: &Runtime, future: F) -> F::Output
    where
        F: Future + Send,
        F::Output: Send,
    {
        if tokio::runtime::Handle::try_current().is_ok() {
            std::thread::scope(
                |scope| match scope.spawn(move || runtime.block_on(future)).join() {
                    Ok(output) => output,
                    Err(panic) => std::panic::resume_unwind(panic),
                },
            )
        } else {
            runtime.block_on(future)
        }
    }

    /// Line-framed JSON-RPC 2.0 over one duplex overlapped named-pipe handle.
    pub struct PipeRpc {
        runtime: Runtime,
        pipe: BufReader<NamedPipeClient>,
        next_id: u64,
    }

    impl PipeRpc {
        pub fn connect(pipe: &str) -> Result<Self, ClientError> {
            let runtime = Builder::new_current_thread()
                .enable_io()
                .enable_time()
                .build()?;
            let client = {
                let _guard = runtime.enter();
                ClientOptions::new().open(pipe)?
            };
            Ok(Self {
                runtime,
                pipe: BufReader::new(client),
                next_id: 0,
            })
        }

        pub fn call(
            &mut self,
            method: &str,
            params: Value,
            timeout: Duration,
            notifications: &mut VecDeque<Value>,
        ) -> Result<Value, ClientError> {
            let id = self.next_id;
            self.next_id += 1;
            let request = json!({ "jsonrpc": "2.0", "id": id, "method": method, "params": params });
            let operation = method.to_string();
            let runtime = &self.runtime;
            let pipe = &mut self.pipe;
            let result = block_on_compatible(runtime, async {
                tokio::time::timeout(timeout, async {
                    pipe
                        .get_mut()
                        .write_all(request.to_string().as_bytes())
                        .await?;
                    pipe.get_mut().write_all(b"\n").await?;
                    pipe.get_mut().flush().await?;

                    let mut unmatched = 0;
                    loop {
                        let mut line = Vec::new();
                        let read = pipe.read_until(b'\n', &mut line).await?;
                        if read == 0 {
                            return Err(ClientError::Disconnected);
                        }
                        while matches!(line.last(), Some(b'\n' | b'\r')) {
                            line.pop();
                        }
                        let message = match serde_json::from_slice::<Value>(&line) {
                            Ok(message) => message,
                            Err(_) => {
                                unmatched += 1;
                                if unmatched >= MAX_UNMATCHED_MESSAGES {
                                    return Err(ClientError::Protocol(
                                        "too many non-JSON messages while awaiting a response"
                                            .into(),
                                    ));
                                }
                                continue;
                            }
                        };

                        if message.get("id").is_none() && message.get("method").is_some() {
                            notifications.push_back(message);
                            continue;
                        }

                        match message.get("id").and_then(Value::as_u64) {
                            Some(response_id) if response_id == id => {
                                if let Some(error) = message.get("error") {
                                    return Err(ClientError::from_rpc_error(error));
                                }
                                return Ok(
                                    message.get("result").cloned().unwrap_or(Value::Null)
                                );
                            }
                            _ => {
                                unmatched += 1;
                                if unmatched >= MAX_UNMATCHED_MESSAGES {
                                    return Err(ClientError::Protocol(format!(
                                        "too many unmatched response ids while awaiting request {id}"
                                    )));
                                }
                            }
                        }
                    }
                })
                .await
            });

            match result {
                Ok(result) => result,
                Err(_) => Err(ClientError::Timeout { operation, timeout }),
            }
        }
    }
}

/// Client handle: connection + handshake + command execution over the pipe.
pub struct OfficeHostClient {
    app: String,
    state: SidecarState,
    #[cfg(windows)]
    pipe_name: Option<String>,
    #[cfg(windows)]
    gateway_version: Option<String>,
    notifications: VecDeque<Value>,
    #[cfg(windows)]
    pipe: Option<transport::PipeRpc>,
}

impl OfficeHostClient {
    pub fn new(app: impl Into<String>) -> Self {
        Self {
            app: app.into(),
            state: SidecarState::Requested,
            #[cfg(windows)]
            pipe_name: None,
            #[cfg(windows)]
            gateway_version: None,
            notifications: VecDeque::new(),
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
        self.state = SidecarState::Launching;
        match transport::PipeRpc::connect(pipe) {
            Ok(connection) => {
                self.pipe = Some(connection);
                self.pipe_name = Some(pipe.to_string());
                self.state = SidecarState::Handshaking;
                Ok(self)
            }
            Err(error) => {
                self.state = SidecarState::Degraded;
                Err(error)
            }
        }
    }

    #[cfg(not(windows))]
    pub fn connect(&mut self, _pipe: &str) -> Result<&mut Self, ClientError> {
        Err(ClientError::UnsupportedPlatform)
    }

    /// Opens a named pipe, retrying transient not-found/busy errors until timeout.
    #[cfg(windows)]
    pub fn connect_with_retry(
        &mut self,
        pipe: &str,
        timeout: Duration,
    ) -> Result<&mut Self, ClientError> {
        self.state = SidecarState::Launching;
        match connect_pipe_with_retry(pipe, timeout) {
            Ok(connection) => {
                self.pipe = Some(connection);
                self.pipe_name = Some(pipe.to_string());
                self.state = SidecarState::Handshaking;
                Ok(self)
            }
            Err(error) => {
                self.state = SidecarState::Degraded;
                Err(error)
            }
        }
    }

    #[cfg(not(windows))]
    pub fn connect_with_retry(
        &mut self,
        _pipe: &str,
        _timeout: Duration,
    ) -> Result<&mut Self, ClientError> {
        Err(ClientError::UnsupportedPlatform)
    }

    /// office.host.handshake → protocol-version check + capability manifest.
    #[cfg(windows)]
    pub fn handshake(&mut self, gateway_version: &str) -> Result<HandshakeResponse, ClientError> {
        self.handshake_with_timeout(gateway_version, Duration::from_millis(HANDSHAKE_TIMEOUT_MS))
    }

    #[cfg(windows)]
    pub fn handshake_with_timeout(
        &mut self,
        gateway_version: &str,
        timeout: Duration,
    ) -> Result<HandshakeResponse, ClientError> {
        let response = self.call_with_timeout(
            "office.host.handshake",
            json!({
                "protocol_versions": [PROTOCOL_VERSION],
                "gateway_version": gateway_version,
                "requested_app": self.app,
            }),
            timeout,
        )?;
        let handshake: HandshakeResponse = match serde_json::from_value(response) {
            Ok(handshake) => handshake,
            Err(error) => {
                self.degrade();
                return Err(ClientError::Serde(error));
            }
        };
        if handshake.protocol_version != PROTOCOL_VERSION {
            self.degrade();
            return Err(ClientError::Protocol(format!(
                "host speaks '{}', client requires '{PROTOCOL_VERSION}'",
                handshake.protocol_version
            )));
        }
        self.gateway_version = Some(gateway_version.to_string());
        self.state = SidecarState::Ready;
        Ok(handshake)
    }

    #[cfg(not(windows))]
    pub fn handshake(&mut self, _gateway_version: &str) -> Result<HandshakeResponse, ClientError> {
        Err(ClientError::UnsupportedPlatform)
    }

    #[cfg(not(windows))]
    pub fn handshake_with_timeout(
        &mut self,
        _gateway_version: &str,
        _timeout: Duration,
    ) -> Result<HandshakeResponse, ClientError> {
        Err(ClientError::UnsupportedPlatform)
    }

    /// office.host.ping — liveness plus attach state.
    #[cfg(windows)]
    pub fn ping(&mut self) -> Result<Value, ClientError> {
        self.call("office.host.ping", json!({}))
    }

    #[cfg(windows)]
    pub fn ping_with_timeout(&mut self, timeout: Duration) -> Result<Value, ClientError> {
        self.call_with_timeout("office.host.ping", json!({}), timeout)
    }

    #[cfg(not(windows))]
    pub fn ping(&mut self) -> Result<Value, ClientError> {
        Err(ClientError::UnsupportedPlatform)
    }

    #[cfg(not(windows))]
    pub fn ping_with_timeout(&mut self, _timeout: Duration) -> Result<Value, ClientError> {
        Err(ClientError::UnsupportedPlatform)
    }

    /// Typed sidecar heartbeat. Unlike handshake, this never asks the Host to
    /// attach to or launch an Office application.
    #[cfg(windows)]
    pub fn status(&mut self) -> Result<SidecarStatus, ClientError> {
        Ok(serde_json::from_value(self.ping()?)?)
    }

    #[cfg(not(windows))]
    pub fn status(&mut self) -> Result<SidecarStatus, ClientError> {
        Err(ClientError::UnsupportedPlatform)
    }

    /// office.host.shutdown — graceful sidecar stop (quits the Office app).
    #[cfg(windows)]
    pub fn shutdown(&mut self) -> Result<Value, ClientError> {
        self.call("office.host.shutdown", json!({}))
    }

    #[cfg(windows)]
    pub fn shutdown_with_timeout(&mut self, timeout: Duration) -> Result<Value, ClientError> {
        self.call_with_timeout("office.host.shutdown", json!({}), timeout)
    }

    #[cfg(not(windows))]
    pub fn shutdown(&mut self) -> Result<Value, ClientError> {
        Err(ClientError::UnsupportedPlatform)
    }

    #[cfg(not(windows))]
    pub fn shutdown_with_timeout(&mut self, _timeout: Duration) -> Result<Value, ClientError> {
        Err(ClientError::UnsupportedPlatform)
    }

    /// office.command.execute — the task-level capability surface (§12.3).
    #[cfg(windows)]
    pub fn execute(&mut self, params: &CommandParams) -> Result<CommandResult, ClientError> {
        let response = self.call("office.command.execute", serde_json::to_value(params)?)?;
        Ok(serde_json::from_value(response)?)
    }

    #[cfg(windows)]
    pub fn execute_with_timeout(
        &mut self,
        params: &CommandParams,
        timeout: Duration,
    ) -> Result<CommandResult, ClientError> {
        let response = self.call_with_timeout(
            "office.command.execute",
            serde_json::to_value(params)?,
            timeout,
        )?;
        Ok(serde_json::from_value(response)?)
    }

    #[cfg(not(windows))]
    pub fn execute(&mut self, _params: &CommandParams) -> Result<CommandResult, ClientError> {
        Err(ClientError::UnsupportedPlatform)
    }

    #[cfg(not(windows))]
    pub fn execute_with_timeout(
        &mut self,
        _params: &CommandParams,
        _timeout: Duration,
    ) -> Result<CommandResult, ClientError> {
        Err(ClientError::UnsupportedPlatform)
    }

    /// office.job.get — polls an asynchronous batch command.
    #[cfg(windows)]
    pub fn job_get(&mut self, job_id: &str) -> Result<JobStatus, ClientError> {
        self.job_get_with_timeout(job_id, Duration::from_millis(DEFAULT_CALL_TIMEOUT_MS))
    }

    #[cfg(windows)]
    pub fn job_get_with_timeout(
        &mut self,
        job_id: &str,
        timeout: Duration,
    ) -> Result<JobStatus, ClientError> {
        let response =
            self.call_with_timeout("office.job.get", json!({ "job_id": job_id }), timeout)?;
        Ok(serde_json::from_value(response)?)
    }

    #[cfg(not(windows))]
    pub fn job_get(&mut self, _job_id: &str) -> Result<JobStatus, ClientError> {
        Err(ClientError::UnsupportedPlatform)
    }

    #[cfg(not(windows))]
    pub fn job_get_with_timeout(
        &mut self,
        _job_id: &str,
        _timeout: Duration,
    ) -> Result<JobStatus, ClientError> {
        Err(ClientError::UnsupportedPlatform)
    }

    /// office.job.cancel — requests cancellation at the next item boundary.
    #[cfg(windows)]
    pub fn job_cancel(&mut self, job_id: &str) -> Result<JobCancelResult, ClientError> {
        let response = self.call("office.job.cancel", json!({ "job_id": job_id }))?;
        Ok(serde_json::from_value(response)?)
    }

    #[cfg(not(windows))]
    pub fn job_cancel(&mut self, _job_id: &str) -> Result<JobCancelResult, ClientError> {
        Err(ClientError::UnsupportedPlatform)
    }

    /// Drains JSON-RPC notifications observed while calls awaited responses.
    pub fn drain_notifications(&mut self) -> Vec<Value> {
        self.notifications.drain(..).collect()
    }

    /// Reconnects to the last pipe and repeats the last successful handshake.
    #[cfg(windows)]
    pub fn recover(&mut self, timeout: Duration) -> Result<HandshakeResponse, ClientError> {
        let pipe_name = self
            .pipe_name
            .clone()
            .ok_or_else(|| ClientError::Protocol("no previous pipe to recover".into()))?;
        let gateway_version = self
            .gateway_version
            .clone()
            .ok_or_else(|| ClientError::Protocol("no successful handshake to recover".into()))?;
        self.state = SidecarState::Recovering;
        let started = std::time::Instant::now();
        let connection = match connect_pipe_with_retry(&pipe_name, timeout) {
            Ok(connection) => connection,
            Err(error) => {
                self.degrade();
                return Err(error);
            }
        };
        self.pipe = Some(connection);
        self.state = SidecarState::Handshaking;
        let remaining = timeout.saturating_sub(started.elapsed());
        if remaining.is_zero() {
            self.degrade();
            return Err(ClientError::Timeout {
                operation: "recovery handshake".into(),
                timeout,
            });
        }
        self.handshake_with_timeout(
            &gateway_version,
            remaining.min(Duration::from_millis(HANDSHAKE_TIMEOUT_MS)),
        )
    }

    #[cfg(not(windows))]
    pub fn recover(&mut self, _timeout: Duration) -> Result<HandshakeResponse, ClientError> {
        Err(ClientError::UnsupportedPlatform)
    }

    #[cfg(windows)]
    fn call(&mut self, method: &str, params: Value) -> Result<Value, ClientError> {
        self.call_with_timeout(
            method,
            params,
            Duration::from_millis(DEFAULT_CALL_TIMEOUT_MS),
        )
    }

    #[cfg(windows)]
    fn call_with_timeout(
        &mut self,
        method: &str,
        params: Value,
        timeout: Duration,
    ) -> Result<Value, ClientError> {
        let result = match self.pipe.as_mut() {
            Some(pipe) => pipe.call(method, params, timeout, &mut self.notifications),
            None => Err(ClientError::Protocol("connect() must precede calls".into())),
        };
        if matches!(
            result,
            Err(ClientError::Io(_))
                | Err(ClientError::Serde(_))
                | Err(ClientError::Protocol(_))
                | Err(ClientError::Disconnected)
                | Err(ClientError::Timeout { .. })
        ) {
            self.degrade();
        }
        result
    }

    #[cfg(windows)]
    fn degrade(&mut self) {
        self.pipe = None;
        self.state = SidecarState::Degraded;
    }
}

#[cfg(windows)]
fn connect_pipe_with_retry(
    pipe: &str,
    timeout: Duration,
) -> Result<transport::PipeRpc, ClientError> {
    let started = std::time::Instant::now();
    loop {
        match transport::PipeRpc::connect(pipe) {
            Ok(connection) => return Ok(connection),
            Err(ClientError::Io(error)) if matches!(error.raw_os_error(), Some(2) | Some(231)) => {
                let remaining = timeout.saturating_sub(started.elapsed());
                if remaining.is_zero() {
                    return Err(ClientError::Timeout {
                        operation: format!("connect {pipe}"),
                        timeout,
                    });
                }
                std::thread::sleep(remaining.min(Duration::from_millis(200)));
            }
            Err(error) => return Err(error),
        }
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

    #[test]
    fn rpc_error_preserves_indeterminate_recovery_data() {
        let error = ClientError::from_rpc_error(&serde_json::json!({
            "code": "OFFICE_RPC_TIMEOUT",
            "message": "write may have completed",
            "data": { "indeterminate": true }
        }));

        match error {
            ClientError::Rpc { code, data, .. } => {
                assert_eq!(code, "OFFICE_RPC_TIMEOUT");
                assert_eq!(data["indeterminate"], true);
            }
            other => panic!("expected RPC error, got {other:?}"),
        }
    }
}
