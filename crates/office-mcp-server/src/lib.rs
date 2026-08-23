//! Interim, contract-first MCP surface for Office task-level capabilities.
//!
//! The server is intentionally thin: the canonical capability catalog owns
//! names and schemas, `office-security` owns the first policy gate, and an
//! injected bridge owns sidecar lifecycle and `office-rpc/1` transport.

#![forbid(unsafe_code)]

mod production;
mod schemas;

use std::collections::{BTreeMap, BTreeSet};
use std::fmt;
use std::path::{Path, PathBuf};
use std::sync::{Arc, Mutex};

use dcc_mcp_office_protocol::{
    capability_catalog, CommandParams, ConfirmationProof, OfficeErrorCode,
};
use dcc_mcp_office_security::validate_policy_tightening;
use jsonschema::Validator;
use rmcp::model::{
    CallToolRequestParams, CallToolResponse, CallToolResult, Implementation, JsonObject,
    ListToolsResult, PaginatedRequestParams, ServerCapabilities, ServerInfo, Tool, ToolAnnotations,
};
use rmcp::service::{RequestContext, RoleServer};
use rmcp::{ErrorData as McpError, ServerHandler};
use serde_json::{json, Map, Value};

use production::ProductionBridge;

const JOB_GET_TOOL: &str = "office.job.get";
const JOB_CANCEL_TOOL: &str = "office.job.cancel";

/// Sidecar-facing error preserved as a caller-visible MCP tool error.
#[derive(Debug, Clone, PartialEq)]
pub struct BridgeError {
    /// Closed `OFFICE_*` wire code.
    pub code: String,
    /// Human-readable failure summary.
    pub message: String,
    /// Structured sidecar context, when available.
    pub data: Value,
    /// Whether the canonical catalog permits a bounded retry.
    pub retryable: bool,
}

impl BridgeError {
    /// Builds an error from the closed Office error enum and catalog metadata.
    pub fn office(code: OfficeErrorCode, message: impl Into<String>, data: Value) -> Self {
        let code = wire_code(code);
        let retryable = capability_catalog()
            .errors
            .iter()
            .find(|error| error.code == code)
            .is_some_and(|error| error.retryable);
        Self {
            code,
            message: message.into(),
            data,
            retryable,
        }
    }

    /// Builds a canonical infrastructure failure.
    pub fn backend(message: impl Into<String>) -> Self {
        Self::office(
            OfficeErrorCode::OfficeBackendUnavailable,
            message,
            Value::Null,
        )
    }

    fn into_tool_result(self) -> CallToolResult {
        CallToolResult::structured_error(json!({
            "code": self.code,
            "message": self.message,
            "data": self.data,
            "retryable": self.retryable,
        }))
    }
}

impl fmt::Display for BridgeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "{}: {}", self.code, self.message)
    }
}

impl std::error::Error for BridgeError {}

/// Narrow application boundary used by the MCP adapter and fake contract tests.
pub trait OfficeBridge: Send {
    /// Selected Office application.
    fn app(&self) -> &str;
    /// Process-bound workspace root.
    fn workspace_root(&self) -> &Path;
    /// Live capability set returned by the sidecar handshake.
    fn capabilities(&self) -> BTreeSet<String>;
    /// Executes one catalog capability.
    fn execute(&mut self, params: CommandParams) -> Result<Value, BridgeError>;
    /// Reads one asynchronous job.
    fn job_get(&mut self, job_id: &str) -> Result<Value, BridgeError>;
    /// Requests cooperative cancellation of one asynchronous job.
    fn job_cancel(&mut self, job_id: &str) -> Result<Value, BridgeError>;
}

/// Failure to construct a schema-valid MCP surface or start its Office Host.
#[derive(Debug)]
pub struct ServerBuildError(String);

impl fmt::Display for ServerBuildError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.0)
    }
}

impl std::error::Error for ServerBuildError {}

#[derive(Clone)]
enum Operation {
    Capability(String),
    JobGet,
    JobCancel,
}

struct Binding {
    operation: Operation,
    input_validator: Validator,
    output_validator: Validator,
}

/// Dynamic MCP handler built from one live sidecar handshake.
pub struct OfficeMcpServer {
    app: String,
    workspace_root: PathBuf,
    bridge: Arc<BridgeCell>,
    tools: Vec<Tool>,
    bindings: BTreeMap<String, Binding>,
}

impl OfficeMcpServer {
    /// Builds a dynamic MCP surface from an injected, already-handshaken bridge.
    pub fn from_bridge(bridge: Box<dyn OfficeBridge>) -> Result<Self, ServerBuildError> {
        let app = bridge.app().to_string();
        let workspace_root = bridge.workspace_root().canonicalize().map_err(|error| {
            ServerBuildError(format!(
                "workspace root cannot be resolved ({}): {error}",
                bridge.workspace_root().display()
            ))
        })?;
        let advertised = bridge.capabilities();
        let mut tools = Vec::new();
        let mut bindings = BTreeMap::new();
        let mut has_jobs = false;

        for capability in &capability_catalog().capabilities {
            let supported_for_app = capability
                .availability
                .iter()
                .any(|availability| availability.apps.iter().any(|candidate| candidate == &app));
            if !supported_for_app || !advertised.contains(&capability.name) {
                continue;
            }
            let capability_schema =
                schemas::load(&capability.input_schema).map_err(ServerBuildError)?;
            let input_schema =
                schemas::command_tool_input(&capability_schema).map_err(ServerBuildError)?;
            let output_schema =
                schemas::load(&capability.output_schema).map_err(ServerBuildError)?;
            let title = capability_schema["title"]
                .as_str()
                .unwrap_or(&capability.mcp_tool);
            let description = format!(
                "{} (office-rpc capability {} v{})",
                capability_schema["description"]
                    .as_str()
                    .unwrap_or("Office task-level capability"),
                capability.name,
                capability.version
            );
            let tool = tool_definition(ToolDefinition {
                name: &capability.mcp_tool,
                title,
                description: &description,
                input_schema: &input_schema,
                output_schema: &output_schema,
                read_only: capability.name == "document.inspect",
                destructive: capability.name == "batch.replace_text",
            })?;
            let binding = Binding {
                operation: Operation::Capability(capability.name.clone()),
                input_validator: compile_schema(&input_schema)?,
                output_validator: compile_schema(&output_schema)?,
            };
            has_jobs |= capability.name.starts_with("batch.");
            if bindings
                .insert(capability.mcp_tool.clone(), binding)
                .is_some()
            {
                return Err(ServerBuildError(format!(
                    "duplicate MCP tool mapping: {}",
                    capability.mcp_tool
                )));
            }
            tools.push(tool);
        }
        if has_jobs {
            add_job_tool(
                JOB_GET_TOOL,
                "Get Office job status",
                "Read the current phase, progress, and terminal result for an Office batch job.",
                "schemas/job-status.schema.json",
                Operation::JobGet,
                &mut tools,
                &mut bindings,
            )?;
            add_job_tool(
                JOB_CANCEL_TOOL,
                "Cancel Office job",
                "Request cooperative cancellation at the next Office batch item boundary.",
                "schemas/job-cancel-result.schema.json",
                Operation::JobCancel,
                &mut tools,
                &mut bindings,
            )?;
        }
        tools.sort_by(|left, right| left.name.cmp(&right.name));

        Ok(Self {
            app,
            workspace_root,
            bridge: Arc::new(BridgeCell::new(bridge)),
            tools,
            bindings,
        })
    }

    /// Locates and starts an Office Host without blocking the MCP async runtime.
    pub async fn start(
        app: &str,
        workspace_root: impl AsRef<Path>,
        host_override: Option<&Path>,
    ) -> Result<Self, ServerBuildError> {
        let app = app.to_string();
        let workspace_root = workspace_root.as_ref().to_path_buf();
        let host_override = host_override.map(Path::to_path_buf);
        tokio::task::spawn_blocking(move || {
            let bridge = ProductionBridge::start(&app, &workspace_root, host_override.as_deref())
                .map_err(|error| ServerBuildError(error.to_string()))?;
            Self::from_bridge(Box::new(bridge))
        })
        .await
        .map_err(|error| ServerBuildError(format!("Office Host startup task failed: {error}")))?
    }

    /// Returns the live, application-filtered MCP tool definitions.
    pub fn tool_definitions(&self) -> &[Tool] {
        &self.tools
    }

    /// Validates and dispatches one MCP tool call.
    pub async fn invoke(
        &self,
        name: &str,
        arguments: JsonObject,
    ) -> Result<CallToolResult, McpError> {
        let value = Value::Object(arguments.clone());
        let operation = {
            let binding = self.bindings.get(name).ok_or_else(|| {
                McpError::invalid_params(
                    format!("unknown or unavailable Office tool: {name}"),
                    Some(json!({"tool": name, "app": self.app})),
                )
            })?;
            validate_instance(&binding.input_validator, &value, "tool arguments")?;
            binding.operation.clone()
        };

        let bridge = Arc::clone(&self.bridge);
        let workspace_root = self.workspace_root.clone();
        let output = tokio::task::spawn_blocking(move || {
            bridge.with_mut(|bridge| match operation {
                Operation::Capability(capability) => {
                    let params = command_params(&capability, arguments, &workspace_root)?;
                    bridge.execute(params)
                }
                Operation::JobGet => bridge.job_get(required_job_id(&arguments)?),
                Operation::JobCancel => bridge.job_cancel(required_job_id(&arguments)?),
            })
        })
        .await
        .map_err(|error| McpError::internal_error(error.to_string(), None))?;

        match output {
            Ok(value) => {
                let binding = self.bindings.get(name).expect("binding was resolved above");
                if let Err(error) = validate_output(&binding.output_validator, &value) {
                    return Ok(error.into_tool_result());
                }
                Ok(CallToolResult::structured(value))
            }
            Err(error) => Ok(error.into_tool_result()),
        }
    }
}

/// Owns the synchronous bridge without dropping its nested pipe runtime on a
/// Tokio worker during MCP service shutdown.
struct BridgeCell(Mutex<Option<Box<dyn OfficeBridge>>>);

impl BridgeCell {
    fn new(bridge: Box<dyn OfficeBridge>) -> Self {
        Self(Mutex::new(Some(bridge)))
    }

    fn with_mut(
        &self,
        operation: impl FnOnce(&mut dyn OfficeBridge) -> Result<Value, BridgeError>,
    ) -> Result<Value, BridgeError> {
        let mut slot = self
            .0
            .lock()
            .map_err(|_| BridgeError::backend("Office bridge lock was poisoned"))?;
        let bridge = slot
            .as_deref_mut()
            .ok_or_else(|| BridgeError::backend("Office bridge is shutting down"))?;
        operation(bridge)
    }
}

impl Drop for BridgeCell {
    fn drop(&mut self) {
        let bridge = match self.0.get_mut() {
            Ok(slot) => slot.take(),
            Err(poisoned) => poisoned.into_inner().take(),
        };
        let Some(bridge) = bridge else {
            return;
        };
        if tokio::runtime::Handle::try_current().is_err() {
            drop(bridge);
            return;
        }
        if std::thread::spawn(move || drop(bridge)).join().is_err() {
            eprintln!("dcc-mcp-office-mcp-server: Office bridge shutdown panicked");
        }
    }
}

impl ServerHandler for OfficeMcpServer {
    fn get_info(&self) -> ServerInfo {
        ServerInfo::new(ServerCapabilities::builder().enable_tools().build())
            .with_server_info(Implementation::new(
                "dcc-mcp-office-mcp-server",
                env!("CARGO_PKG_VERSION"),
            ))
            .with_instructions(format!(
                "Task-level Office tools for the '{}' sidecar. Paths are restricted to {}.",
                self.app,
                self.workspace_root.display()
            ))
    }

    async fn list_tools(
        &self,
        _request: Option<PaginatedRequestParams>,
        _context: RequestContext<RoleServer>,
    ) -> Result<ListToolsResult, McpError> {
        Ok(ListToolsResult::with_all_items(self.tools.clone()))
    }

    fn get_tool(&self, name: &str) -> Option<Tool> {
        self.tools.iter().find(|tool| tool.name == name).cloned()
    }

    async fn call_tool(
        &self,
        request: CallToolRequestParams,
        _context: RequestContext<RoleServer>,
    ) -> Result<CallToolResponse, McpError> {
        Ok(self
            .invoke(request.name.as_ref(), request.arguments.unwrap_or_default())
            .await?
            .into())
    }
}

fn command_params(
    capability: &str,
    mut arguments: JsonObject,
    workspace_root: &Path,
) -> Result<CommandParams, BridgeError> {
    let document = arguments.remove("document");
    let confirmation = arguments.remove("confirmation");
    let policy = arguments.remove("policy");
    let mut envelope = Map::new();
    envelope.insert("capability".into(), Value::String(capability.to_string()));
    envelope.insert("input".into(), Value::Object(arguments));
    if let Some(value) = document {
        envelope.insert("document".into(), value);
    }
    if let Some(value) = confirmation {
        envelope.insert("confirmation".into(), value);
    }
    if let Some(value) = policy {
        envelope.insert("policy".into(), value);
    }
    let params: CommandParams =
        serde_json::from_value(Value::Object(envelope)).map_err(|error| {
            BridgeError::office(
                OfficeErrorCode::OfficeInvalidRequest,
                format!("invalid Office command envelope: {error}"),
                Value::Null,
            )
        })?;
    validate_policy_tightening(&params.policy)
        .map_err(|violation| BridgeError::office(violation.code, violation.message, Value::Null))?;
    validate_workspace_echo(&params.policy, workspace_root)?;
    validate_confirmation(capability, &params)?;
    Ok(params)
}

fn validate_workspace_echo(policy: &Value, workspace_root: &Path) -> Result<(), BridgeError> {
    let Some(requested) = policy.get("workspace_root").and_then(Value::as_str) else {
        return Ok(());
    };
    let requested = Path::new(requested).canonicalize().map_err(|error| {
        BridgeError::office(
            OfficeErrorCode::OfficeAccessDenied,
            format!("policy.workspace_root cannot be resolved: {error}"),
            Value::Null,
        )
    })?;
    if requested != workspace_root {
        return Err(BridgeError::office(
            OfficeErrorCode::OfficeAccessDenied,
            "policy.workspace_root cannot replace the server-bound workspace",
            Value::Null,
        ));
    }
    Ok(())
}

fn validate_confirmation(capability: &str, params: &CommandParams) -> Result<(), BridgeError> {
    let requires_confirmation = match capability {
        "batch.replace_text" => params.input["dry_run"].as_bool() == Some(false),
        "batch.convert" => params.input["overwrite"].as_str() == Some("overwrite"),
        _ => false,
    };
    if !requires_confirmation {
        return Ok(());
    }
    if params.policy["overwrite_original"].as_str() == Some("deny") {
        return Err(BridgeError::office(
            OfficeErrorCode::OfficeAccessDenied,
            "policy.overwrite_original denies this operation",
            Value::Null,
        ));
    }
    let valid = params
        .confirmation
        .as_ref()
        .is_some_and(valid_overwrite_confirmation);
    if !valid {
        return Err(BridgeError::office(
            OfficeErrorCode::OfficeUserConfirmationRequired,
            "overwrite requires action, confirmed=true, confirmed_by='human:<id>', and confirmed_at",
            Value::Null,
        ));
    }
    Ok(())
}

fn valid_overwrite_confirmation(confirmation: &ConfirmationProof) -> bool {
    confirmation.action == "overwrite_original"
        && confirmation.confirmed
        && confirmation.confirmed_by.starts_with("human:")
        && confirmation.confirmed_by.len() > "human:".len()
        && !confirmation.confirmed_at.trim().is_empty()
}

fn required_job_id(arguments: &JsonObject) -> Result<&str, BridgeError> {
    arguments
        .get("job_id")
        .and_then(Value::as_str)
        .ok_or_else(|| {
            BridgeError::office(
                OfficeErrorCode::OfficeInvalidRequest,
                "job_id is required",
                Value::Null,
            )
        })
}

fn add_job_tool(
    name: &str,
    title: &str,
    description: &str,
    output_schema_path: &str,
    operation: Operation,
    tools: &mut Vec<Tool>,
    bindings: &mut BTreeMap<String, Binding>,
) -> Result<(), ServerBuildError> {
    let input_schema = schemas::load("schemas/job-id.schema.json").map_err(ServerBuildError)?;
    let output_schema = schemas::load(output_schema_path).map_err(ServerBuildError)?;
    tools.push(tool_definition(ToolDefinition {
        name,
        title,
        description,
        input_schema: &input_schema,
        output_schema: &output_schema,
        read_only: name == JOB_GET_TOOL,
        destructive: false,
    })?);
    bindings.insert(
        name.to_string(),
        Binding {
            operation,
            input_validator: compile_schema(&input_schema)?,
            output_validator: compile_schema(&output_schema)?,
        },
    );
    Ok(())
}

struct ToolDefinition<'a> {
    name: &'a str,
    title: &'a str,
    description: &'a str,
    input_schema: &'a Value,
    output_schema: &'a Value,
    read_only: bool,
    destructive: bool,
}

fn tool_definition(definition: ToolDefinition<'_>) -> Result<Tool, ServerBuildError> {
    let input = Arc::new(schemas::object(definition.input_schema).map_err(ServerBuildError)?);
    let output = Arc::new(schemas::object(definition.output_schema).map_err(ServerBuildError)?);
    Ok(Tool::new(
        definition.name.to_string(),
        definition.description.to_string(),
        input,
    )
    .with_title(definition.title)
    .with_raw_output_schema(output)
    .with_annotations(
        ToolAnnotations::new()
            .read_only(definition.read_only)
            .destructive(definition.destructive)
            .open_world(false),
    ))
}

fn compile_schema(schema: &Value) -> Result<Validator, ServerBuildError> {
    jsonschema::validator_for(schema)
        .map_err(|error| ServerBuildError(format!("invalid bundled JSON Schema: {error}")))
}

fn validate_instance(validator: &Validator, instance: &Value, label: &str) -> Result<(), McpError> {
    if validator.is_valid(instance) {
        return Ok(());
    }
    let errors = validator
        .iter_errors(instance)
        .take(4)
        .map(|error| error.to_string())
        .collect::<Vec<_>>();
    Err(McpError::invalid_params(
        format!("invalid {label}: {}", errors.join("; ")),
        Some(json!({"validation_errors": errors})),
    ))
}

fn validate_output(validator: &Validator, value: &Value) -> Result<(), BridgeError> {
    if validator.is_valid(value) {
        return Ok(());
    }
    let errors = validator
        .iter_errors(value)
        .take(4)
        .map(|error| error.to_string())
        .collect::<Vec<_>>();
    Err(BridgeError::backend(format!(
        "sidecar result violated the catalog output schema: {}",
        errors.join("; ")
    )))
}

fn wire_code(code: OfficeErrorCode) -> String {
    serde_json::to_value(code)
        .expect("Office error code must serialize")
        .as_str()
        .expect("Office error code must serialize as a string")
        .to_string()
}
