use std::collections::BTreeSet;
use std::path::Path;

use dcc_mcp_office_protocol::CommandParams;
use serde_json::Value;

use crate::{BridgeError, OfficeBridge};

#[cfg(windows)]
mod platform {
    use std::path::PathBuf;
    use std::process::{Child, Command, Stdio};
    use std::sync::atomic::{AtomicU64, Ordering};
    use std::time::Duration;

    use dcc_mcp_office_client::{locate_office_host, ClientError, OfficeHostClient};
    use dcc_mcp_office_protocol::OfficeErrorCode;

    use super::*;

    static NEXT_PIPE_INSTANCE: AtomicU64 = AtomicU64::new(0);

    pub(crate) struct ProductionBridge {
        app: String,
        workspace_root: PathBuf,
        capabilities: BTreeSet<String>,
        client: OfficeHostClient,
        child: Child,
    }

    impl ProductionBridge {
        pub(crate) fn start(
            app: &str,
            workspace_root: &Path,
            host_override: Option<&Path>,
        ) -> Result<Self, BridgeError> {
            if !matches!(app, "powerpoint" | "word" | "excel") {
                return Err(BridgeError::office(
                    OfficeErrorCode::OfficeCapabilityUnsupported,
                    format!("reference MCP server does not support app '{app}'"),
                    Value::Null,
                ));
            }
            let workspace_root = workspace_root.canonicalize().map_err(|error| {
                BridgeError::office(
                    OfficeErrorCode::OfficeAccessDenied,
                    format!("workspace root cannot be resolved: {error}"),
                    Value::Null,
                )
            })?;
            if !workspace_root.is_dir() {
                return Err(BridgeError::office(
                    OfficeErrorCode::OfficeAccessDenied,
                    "workspace root must be an existing directory",
                    Value::Null,
                ));
            }
            let host = match host_override {
                Some(path) if path.is_file() => path.canonicalize().map_err(|error| {
                    BridgeError::backend(format!("Office Host path cannot be resolved: {error}"))
                })?,
                Some(path) => {
                    return Err(BridgeError::backend(format!(
                        "explicit Office Host is missing: {}",
                        path.display()
                    )))
                }
                None => {
                    locate_office_host()
                        .map_err(|error| BridgeError::backend(error.to_string()))?
                        .path
                }
            };
            let instance = NEXT_PIPE_INSTANCE.fetch_add(1, Ordering::Relaxed);
            let pipe = format!(
                r"\\.\pipe\dcc-mcp-office-mcp-{}-{instance}-{app}",
                std::process::id()
            );
            let mut command = Command::new(&host);
            command
                .arg(format!("--app={app}"))
                .arg(format!("--pipe-name={pipe}"))
                .arg(format!("--parent-pid={}", std::process::id()))
                .arg(format!(
                    "--workspace-root={}",
                    workspace_root.to_string_lossy()
                ))
                .stdin(Stdio::null())
                .stdout(Stdio::null())
                .stderr(Stdio::inherit());
            hide_console_window(&mut command);
            let mut child = command.spawn().map_err(|error| {
                BridgeError::backend(format!("failed to start {}: {error}", host.display()))
            })?;
            let mut client = OfficeHostClient::new(app);
            let handshake = client
                .connect_with_retry(&pipe, Duration::from_secs(15))
                .and_then(|client| client.handshake(env!("CARGO_PKG_VERSION")));
            let handshake = match handshake {
                Ok(handshake) => handshake,
                Err(error) => {
                    let _ = child.kill();
                    let _ = child.wait();
                    return Err(client_error(error));
                }
            };
            let capabilities = handshake
                .capability_manifest
                .capabilities
                .keys()
                .cloned()
                .collect();
            Ok(Self {
                app: app.to_string(),
                workspace_root,
                capabilities,
                client,
                child,
            })
        }
    }

    impl OfficeBridge for ProductionBridge {
        fn app(&self) -> &str {
            &self.app
        }

        fn workspace_root(&self) -> &Path {
            &self.workspace_root
        }

        fn capabilities(&self) -> BTreeSet<String> {
            self.capabilities.clone()
        }

        fn execute(&mut self, params: CommandParams) -> Result<Value, BridgeError> {
            let result = self.client.execute(&params).map_err(client_error)?;
            serde_json::to_value(result).map_err(|error| {
                BridgeError::backend(format!("result serialization failed: {error}"))
            })
        }

        fn job_get(&mut self, job_id: &str) -> Result<Value, BridgeError> {
            let result = self.client.job_get(job_id).map_err(client_error)?;
            serde_json::to_value(result)
                .map_err(|error| BridgeError::backend(format!("job serialization failed: {error}")))
        }

        fn job_cancel(&mut self, job_id: &str) -> Result<Value, BridgeError> {
            let result = self.client.job_cancel(job_id).map_err(client_error)?;
            serde_json::to_value(result)
                .map_err(|error| BridgeError::backend(format!("job serialization failed: {error}")))
        }
    }

    impl Drop for ProductionBridge {
        fn drop(&mut self) {
            let _ = self.client.shutdown_with_timeout(Duration::from_secs(2));
            for _ in 0..80 {
                match self.child.try_wait() {
                    Ok(Some(_)) => return,
                    Ok(None) => std::thread::sleep(Duration::from_millis(25)),
                    Err(_) => break,
                }
            }
            let _ = self.child.kill();
            let _ = self.child.wait();
        }
    }

    fn client_error(error: ClientError) -> BridgeError {
        match error {
            ClientError::Rpc {
                code,
                message,
                data,
            } => {
                let code = code.as_str().unwrap_or("OFFICE_UNCLASSIFIED").to_string();
                let retryable = dcc_mcp_office_protocol::capability_catalog()
                    .errors
                    .iter()
                    .find(|candidate| candidate.code == code)
                    .is_some_and(|candidate| candidate.retryable);
                BridgeError {
                    code,
                    message,
                    data,
                    retryable,
                }
            }
            ClientError::Timeout { operation, timeout } => BridgeError::office(
                OfficeErrorCode::OfficeRpcTimeout,
                format!("{operation} timed out after {timeout:?}"),
                Value::Null,
            ),
            other => BridgeError::backend(other.to_string()),
        }
    }

    fn hide_console_window(command: &mut Command) {
        use std::os::windows::process::CommandExt;

        const CREATE_NO_WINDOW: u32 = 0x0800_0000;
        command.creation_flags(CREATE_NO_WINDOW);
    }
}

#[cfg(not(windows))]
mod platform {
    use super::*;

    pub(crate) struct ProductionBridge;

    impl ProductionBridge {
        pub(crate) fn start(
            _app: &str,
            _workspace_root: &Path,
            _host_override: Option<&Path>,
        ) -> Result<Self, BridgeError> {
            Err(BridgeError::backend(
                "Office Host lifecycle requires an interactive Windows session",
            ))
        }
    }

    impl OfficeBridge for ProductionBridge {
        fn app(&self) -> &str {
            unreachable!("non-Windows production bridge cannot be constructed")
        }

        fn workspace_root(&self) -> &Path {
            unreachable!("non-Windows production bridge cannot be constructed")
        }

        fn capabilities(&self) -> BTreeSet<String> {
            unreachable!("non-Windows production bridge cannot be constructed")
        }

        fn execute(&mut self, _params: CommandParams) -> Result<Value, BridgeError> {
            unreachable!("non-Windows production bridge cannot be constructed")
        }

        fn job_get(&mut self, _job_id: &str) -> Result<Value, BridgeError> {
            unreachable!("non-Windows production bridge cannot be constructed")
        }

        fn job_cancel(&mut self, _job_id: &str) -> Result<Value, BridgeError> {
            unreachable!("non-Windows production bridge cannot be constructed")
        }
    }
}

pub(crate) use platform::ProductionBridge;
