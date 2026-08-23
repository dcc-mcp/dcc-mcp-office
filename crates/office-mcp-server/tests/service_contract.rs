use std::collections::BTreeSet;
use std::path::{Path, PathBuf};
use std::sync::{Arc, Mutex};

use dcc_mcp_office_mcp_server::{BridgeError, OfficeBridge, OfficeMcpServer};
use dcc_mcp_office_protocol::{CommandParams, OfficeErrorCode};
use rmcp::model::{CallToolRequestParams, ClientInfo};
use rmcp::{ClientHandler, ServiceExt};
use serde_json::{json, Map, Value};

#[derive(Clone)]
struct FakeState {
    commands: Arc<Mutex<Vec<CommandParams>>>,
    jobs: Arc<Mutex<Vec<(String, String)>>>,
}

impl FakeState {
    fn new() -> Self {
        Self {
            commands: Arc::new(Mutex::new(Vec::new())),
            jobs: Arc::new(Mutex::new(Vec::new())),
        }
    }
}

struct FakeBridge {
    app: String,
    capabilities: BTreeSet<String>,
    state: FakeState,
    workspace_root: PathBuf,
    execute_result: Result<Value, BridgeError>,
}

impl FakeBridge {
    fn new(state: FakeState, capabilities: &[&str]) -> Self {
        Self {
            app: "powerpoint".into(),
            capabilities: capabilities.iter().map(|value| (*value).into()).collect(),
            state,
            workspace_root: std::env::current_dir().expect("current directory"),
            execute_result: Ok(command_result()),
        }
    }

    fn failing(mut self, error: BridgeError) -> Self {
        self.execute_result = Err(error);
        self
    }
}

impl OfficeBridge for FakeBridge {
    fn app(&self) -> &str {
        &self.app
    }

    fn capabilities(&self) -> BTreeSet<String> {
        self.capabilities.clone()
    }

    fn workspace_root(&self) -> &Path {
        &self.workspace_root
    }

    fn execute(&mut self, params: CommandParams) -> Result<Value, BridgeError> {
        self.state.commands.lock().unwrap().push(params);
        self.execute_result.clone()
    }

    fn job_get(&mut self, job_id: &str) -> Result<Value, BridgeError> {
        self.state
            .jobs
            .lock()
            .unwrap()
            .push(("get".into(), job_id.into()));
        Ok(json!({
            "job_id": job_id,
            "capability": "batch.convert",
            "phase": "running",
            "stage": "convert",
            "completed": 1,
            "total": 2,
            "cancel_requested": false,
            "created_at": "2026-08-24T00:00:00Z",
            "updated_at": "2026-08-24T00:00:01Z",
            "result": null,
            "error": null
        }))
    }

    fn job_cancel(&mut self, job_id: &str) -> Result<Value, BridgeError> {
        self.state
            .jobs
            .lock()
            .unwrap()
            .push(("cancel".into(), job_id.into()));
        Ok(json!({"job_id": job_id, "accepted": true, "phase": "running"}))
    }
}

#[derive(Debug, Clone, Default)]
struct TestClient;

impl ClientHandler for TestClient {
    fn get_info(&self) -> ClientInfo {
        ClientInfo::default()
    }
}

#[tokio::test]
async fn catalog_tools_use_real_schemas_and_map_to_the_wire_capability() {
    let state = FakeState::new();
    let server = server(&state, &["deck.compile", "document.inspect"]);

    let tools = server.tool_definitions();
    let generate = tools
        .iter()
        .find(|tool| tool.name == "powerpoint.deck.generate")
        .expect("deck tool");
    assert_eq!(generate.input_schema["required"], json!(["ir", "output"]));
    assert!(generate.input_schema["properties"]["confirmation"].is_object());
    assert!(generate.output_schema.is_some());
    assert!(tools
        .iter()
        .any(|tool| tool.name == "office.document.inspect"));
    assert!(!tools.iter().any(|tool| tool.name == "office.batch.convert"));

    let result = server
        .invoke("powerpoint.deck.generate", deck_arguments())
        .await
        .expect("tool call");

    assert_eq!(result.is_error, Some(false));
    assert_eq!(
        result.structured_content.as_ref().unwrap()["operation_id"],
        "op-1"
    );
    let calls = state.commands.lock().unwrap();
    assert_eq!(calls.len(), 1);
    assert_eq!(calls[0].capability, "deck.compile");
    assert_eq!(calls[0].input, json!({"ir": "{}", "output": "deck.pptx"}));
}

#[tokio::test]
async fn invalid_schema_policy_and_confirmation_fail_before_sidecar_dispatch() {
    let state = FakeState::new();
    let server = server(&state, &["deck.compile", "batch.replace_text"]);

    let invalid_schema = server
        .invoke("powerpoint.deck.generate", object(json!({"ir": "{}"})))
        .await;
    assert!(invalid_schema.is_err());

    let policy_relaxation = server
        .invoke(
            "powerpoint.deck.generate",
            object(json!({
                "ir": "{}",
                "output": "deck.pptx",
                "policy": {"macros": "confirm"}
            })),
        )
        .await
        .expect("policy rejection is a structured tool error");
    assert_structured_error(&policy_relaxation, "OFFICE_MACRO_BLOCKED", false);

    let missing_confirmation = server
        .invoke(
            "office.batch.replace_text",
            object(json!({
                "inputs": ["deck.pptx"],
                "rules": [{"find": "before", "replace": "after"}],
                "dry_run": false,
                "policy": {"overwrite_original": "checkpoint_and_confirm"}
            })),
        )
        .await
        .expect("confirmation rejection is a structured tool error");
    assert_structured_error(
        &missing_confirmation,
        "OFFICE_USER_CONFIRMATION_REQUIRED",
        false,
    );

    assert!(state.commands.lock().unwrap().is_empty());
}

#[tokio::test]
async fn sidecar_errors_keep_office_code_data_and_retryability() {
    let state = FakeState::new();
    let bridge = FakeBridge::new(state.clone(), &["deck.compile"]).failing(BridgeError::office(
        OfficeErrorCode::OfficeAppBusy,
        "PowerPoint is busy",
        json!({"retry_after_ms": 250}),
    ));
    let server = OfficeMcpServer::from_bridge(Box::new(bridge)).expect("server");

    let result = server
        .invoke("powerpoint.deck.generate", deck_arguments())
        .await
        .expect("tool call");

    assert_structured_error(&result, "OFFICE_APP_BUSY", true);
    assert_eq!(
        result.structured_content.as_ref().unwrap()["data"]["retry_after_ms"],
        250
    );
    assert_eq!(state.commands.lock().unwrap().len(), 1);
}

#[tokio::test]
async fn batch_capabilities_add_job_tools_and_route_job_calls() {
    let state = FakeState::new();
    let server = server(&state, &["batch.convert"]);
    let names = server
        .tool_definitions()
        .iter()
        .map(|tool| tool.name.as_ref())
        .collect::<Vec<_>>();
    assert!(names.contains(&"office.batch.convert"));
    assert!(names.contains(&"office.job.get"));
    assert!(names.contains(&"office.job.cancel"));

    let job_id = "job:0123456789abcdef0123456789abcdef";
    let status = server
        .invoke("office.job.get", object(json!({"job_id": job_id})))
        .await
        .expect("job status");
    let cancel = server
        .invoke("office.job.cancel", object(json!({"job_id": job_id})))
        .await
        .expect("job cancel");

    assert_eq!(status.is_error, Some(false));
    assert_eq!(cancel.is_error, Some(false));
    assert_eq!(
        state.jobs.lock().unwrap().as_slice(),
        [
            ("get".into(), job_id.into()),
            ("cancel".into(), job_id.into())
        ]
    );
}

#[tokio::test]
async fn official_mcp_client_can_initialize_list_and_call_over_stdio_framing() {
    let state = FakeState::new();
    let server = server(&state, &["deck.compile"]);
    let (server_transport, client_transport) = tokio::io::duplex(16 * 1024);
    let server_task = tokio::spawn(async move {
        server
            .serve(server_transport)
            .await
            .expect("serve MCP")
            .waiting()
            .await
            .expect("MCP server lifecycle");
    });
    let client = TestClient
        .serve(client_transport)
        .await
        .expect("initialize MCP client");

    let tools = client.list_all_tools().await.expect("list MCP tools");
    assert_eq!(tools.len(), 1);
    assert_eq!(tools[0].name, "powerpoint.deck.generate");
    let response = client
        .call_tool(
            CallToolRequestParams::new("powerpoint.deck.generate").with_arguments(deck_arguments()),
        )
        .await
        .expect("call MCP tool");
    assert_eq!(response.is_error, Some(false));
    assert_eq!(response.structured_content.unwrap()["operation_id"], "op-1");

    client.cancel().await.expect("close MCP client");
    server_task.await.expect("join MCP server");
    assert_eq!(state.commands.lock().unwrap().len(), 1);
}

fn server(state: &FakeState, capabilities: &[&str]) -> OfficeMcpServer {
    OfficeMcpServer::from_bridge(Box::new(FakeBridge::new(state.clone(), capabilities)))
        .expect("construct reference MCP server")
}

fn command_result() -> Value {
    json!({
        "operation_id": "op-1",
        "changed": {},
        "warnings": [],
        "artefacts": [],
        "validation": {},
        "backend": "openxml",
        "indeterminate": false,
        "audit": {"policy_decision": "allowed"}
    })
}

fn deck_arguments() -> Map<String, Value> {
    object(json!({"ir": "{}", "output": "deck.pptx"}))
}

fn object(value: Value) -> Map<String, Value> {
    value.as_object().expect("JSON object").clone()
}

fn assert_structured_error(result: &rmcp::model::CallToolResult, code: &str, retryable: bool) {
    assert_eq!(result.is_error, Some(true));
    let structured = result
        .structured_content
        .as_ref()
        .expect("structured error");
    assert_eq!(structured["code"], code);
    assert_eq!(structured["retryable"], retryable);
}
