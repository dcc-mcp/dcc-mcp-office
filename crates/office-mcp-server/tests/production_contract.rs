//! Windows contract: the reference MCP server owns a real Office Host process
//! and executes one deterministic Open XML capability without attaching COM.

#![cfg(windows)]

use std::path::PathBuf;

use dcc_mcp_office_mcp_server::OfficeMcpServer;
use rmcp::model::{CallToolRequestParams, ClientInfo};
use rmcp::{ClientHandler, ServiceExt};
use serde_json::json;

#[derive(Debug, Clone, Default)]
struct TestClient;

impl ClientHandler for TestClient {
    fn get_info(&self) -> ClientInfo {
        ClientInfo::default()
    }
}

#[tokio::test]
#[ignore = "requires DCC_OFFICE_HOST_EXE; CI runs this Office-free contract explicitly"]
async fn mcp_server_spawns_host_and_compiles_deck_without_office() {
    let host = std::env::var_os("DCC_OFFICE_HOST_EXE")
        .map(PathBuf::from)
        .filter(|path| path.is_file())
        .expect("DCC_OFFICE_HOST_EXE must point to dcc-office-host.exe");
    let workspace =
        std::env::temp_dir().join(format!("dcc-office-mcp-contract-{}", std::process::id()));
    if workspace.exists() {
        std::fs::remove_dir_all(&workspace).expect("remove stale contract workspace");
    }
    std::fs::create_dir_all(&workspace).expect("create contract workspace");
    let output = workspace.join("deck.pptx");
    let server = OfficeMcpServer::start("powerpoint", &workspace, Some(&host))
        .await
        .expect("start MCP server and Office Host");
    let (server_transport, client_transport) = tokio::io::duplex(64 * 1024);
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
    let response = client
        .call_tool(
            CallToolRequestParams::new("powerpoint.deck.generate").with_arguments(
                json!({
                    "ir": SAMPLE_DECK_IR,
                    "output": output.to_string_lossy()
                })
                .as_object()
                .expect("arguments object")
                .clone(),
            ),
        )
        .await
        .expect("MCP tool call");

    assert_eq!(response.is_error, Some(false), "{response:?}");
    assert_eq!(
        response.structured_content.as_ref().unwrap()["backend"],
        "openxml"
    );
    assert!(output.is_file(), "deck.compile did not create output");

    client.cancel().await.expect("close MCP client");
    server_task.await.expect("join MCP server");
    std::fs::remove_dir_all(&workspace).expect("remove contract workspace");
}

const SAMPLE_DECK_IR: &str = r#"{
  "schema_version": "office-ir/1.0",
  "kind": "presentation",
  "document_id": "draft:mcp-contract",
  "metadata": {"title": "MCP Contract", "language": "en-US"},
  "document": {
    "slides": [{
      "semantic_layout": "title_cover",
      "title": "MCP Contract",
      "content_blocks": [{"type": "text", "paragraphs": ["office-rpc/1"]}]
    }]
  },
  "outputs": ["pptx"]
}"#;
