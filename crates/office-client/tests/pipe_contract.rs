//! Cross-language contract test: the Rust client drives a real C# dcc-office-host
//! over the office-rpc/1 named pipe, exercising the full M1 surface:
//!
//!   handshake → deck.compile (Open XML) → document.inspect (COM) →
//!   batch.convert (COM PDF) → batch.replace_text dry-run/commit (COM) →
//!   slide.render (COM previews) → error contract.
//!
//! Requirements (test skips otherwise):
//!   - env DCC_OFFICE_HOST_EXE pointing at the built dcc-office-host.exe
//!   - Microsoft PowerPoint installed (the COM legs skip if the handshake
//!     manifest lacks desktop_com capabilities — progressive discovery).
#![cfg(windows)]

use std::path::PathBuf;
use std::process::{Child, Command, Stdio};
use std::time::{Duration, Instant};

use dcc_mcp_office_client::OfficeHostClient;
use dcc_mcp_office_protocol::CommandParams;
use serde_json::json;

const PIPE: &str = r"\\.\pipe\dcc-office-contract-test-powerpoint";

struct HostGuard(Child);

impl HostGuard {
    /// Graceful teardown: office.host.shutdown quits the COM app and exits
    /// the host — force-killing a host orphans its Office process (which then
    /// holds document locks). Kill stays as a last-resort fallback in Drop.
    fn shutdown(&mut self, client: &mut OfficeHostClient) {
        let _ = client.shutdown();
        let start = Instant::now();
        while start.elapsed() < Duration::from_secs(30) {
            if let Ok(Some(_)) = self.0.try_wait() {
                return;
            }
            std::thread::sleep(Duration::from_millis(200));
        }
        eprintln!("host did not exit after shutdown; falling back to kill");
    }
}

impl Drop for HostGuard {
    fn drop(&mut self) {
        if self.0.try_wait().ok().flatten().is_none() {
            let _ = self.0.kill();
            let _ = self.0.wait();
        }
    }
}

fn host_exe() -> Option<PathBuf> {
    std::env::var_os("DCC_OFFICE_HOST_EXE")
        .map(PathBuf::from)
        .filter(|p| p.is_file())
}

/// Connects with retry: the host accepts one client at a time, and the
/// server-side may still be draining the previous connection when we arrive
/// (ERROR_PIPE_BUSY). Retrying until the instance frees is the honest client
/// behaviour — the gateway does the same.
fn connect_with_retry(
    client: &mut OfficeHostClient,
    child: &mut Child,
    pipe: &str,
    timeout: Duration,
) -> bool {
    let start = Instant::now();
    while start.elapsed() < timeout {
        match client.connect(pipe) {
            Ok(_) => return true,
            Err(dcc_mcp_office_client::ClientError::Io(e))
                if matches!(
                    e.raw_os_error(),
                    Some(231 /* ERROR_PIPE_BUSY */) | Some(2 /* ERROR_FILE_NOT_FOUND */)
                ) =>
            {
                std::thread::sleep(Duration::from_millis(200));
            }
            Err(error) => panic!("connect failed: {error}"),
        }
        if let Ok(Some(status)) = child.try_wait() {
            eprintln!("host exited early: {status}");
            return false;
        }
    }
    false
}

fn command(capability: &str, input: serde_json::Value) -> CommandParams {
    CommandParams {
        capability: capability.to_string(),
        document: None,
        input,
        policy: json!({}),
    }
}

#[test]
fn office_host_full_contract() {
    let Some(exe) = host_exe() else {
        eprintln!("SKIP: DCC_OFFICE_HOST_EXE not set");
        return;
    };

    let pipe_arg = format!("--pipe-name={PIPE}");
    let child = Command::new(&exe)
        .args(["--app=powerpoint", "--pipe", &pipe_arg])
        .stderr(Stdio::null())
        .spawn()
        .expect("spawn dcc-office-host");
    let mut guard = HostGuard(child);

    let mut client = OfficeHostClient::new("powerpoint");
    if !connect_with_retry(&mut client, &mut guard.0, PIPE, Duration::from_secs(30)) {
        panic!("host never accepted a connection");
    }
    let handshake = client.handshake("contract-test-0.1.0").expect("handshake");
    assert_eq!(handshake.protocol_version, "office-rpc/1");
    assert!(handshake
        .capability_manifest
        .capabilities
        .contains_key("deck.compile"));
    assert_eq!(client.state(), dcc_mcp_office_protocol::SidecarState::Ready);

    let ping = client.ping().expect("ping");
    assert_eq!(ping["app"], "powerpoint");

    if !handshake
        .capability_manifest
        .capabilities
        .contains_key("batch.convert")
    {
        eprintln!("SKIP: PowerPoint desktop COM not available in this session");
        return;
    }

    let temp = std::env::temp_dir().join(format!("dcc-office-contract-{}", std::process::id()));
    let _ = std::fs::remove_dir_all(&temp);
    std::fs::create_dir_all(&temp).expect("create temp dir");

    // 1. deck.compile — Open XML worker, no Office needed
    let pptx = temp.join("deck.pptx");
    let compiled = client
        .execute(&command(
            "deck.compile",
            json!({ "ir": SAMPLE_DECK_IR, "output": pptx.to_string_lossy() }),
        ))
        .expect("deck.compile");
    assert_eq!(compiled.backend.as_deref(), Some("openxml"));
    assert_eq!(compiled.changed["slides"], 3);
    assert!(pptx.exists());

    // 1b. every command result carries the §27 criterion-10 audit trail
    let audit = compiled
        .audit
        .as_ref()
        .expect("deck.compile result must carry an audit trail");
    assert_eq!(audit["security"]["automation_security"], "force_disable");
    assert_eq!(audit["backend"], "openxml");
    assert!(audit["duration_ms"].as_u64().unwrap_or(0) > 0 || audit["duration_ms"].is_null());

    // 1c. §19 second-layer policy gate: relaxing deny-by-default is refused
    let policy = CommandParams {
        capability: "batch.replace_text".to_string(),
        document: None,
        input: json!({}),
        policy: json!({ "macros": "confirm" }),
    };
    match client.execute(&policy) {
        Err(dcc_mcp_office_client::ClientError::Rpc { code, .. }) => {
            assert_eq!(code, json!("OFFICE_MACRO_BLOCKED"));
        }
        other => panic!("macros relaxation must be denied, got {other:?}"),
    }
    let policy = CommandParams {
        capability: "batch.replace_text".to_string(),
        document: None,
        input: json!({}),
        policy: json!({ "arbitrary_execute_mso": "confirm" }),
    };
    match client.execute(&policy) {
        Err(dcc_mcp_office_client::ClientError::Rpc { code, .. }) => {
            assert_eq!(code, json!("OFFICE_CAPABILITY_UNSUPPORTED"));
        }
        other => panic!("ExecuteMso relaxation must be denied, got {other:?}"),
    }

    // 1d. brand template gate — unknown brand:// URIs are refused up front
    let template_error = client
        .execute(&command(
            "deck.compile",
            json!({
                "ir": SAMPLE_DECK_IR,
                "output": pptx.to_string_lossy(),
                "template": "brand://unknown/x",
            }),
        ))
        .expect_err("unknown brand template must fail");
    match template_error {
        dcc_mcp_office_client::ClientError::Rpc { code, .. } => {
            assert_eq!(code, json!("OFFICE_CAPABILITY_UNSUPPORTED"));
        }
        other => panic!("expected Rpc error, got {other:?}"),
    }

    // 2. document.inspect — COM backend (auto) once PowerPoint is attached
    let inspect = client
        .execute(&command(
            "document.inspect",
            json!({ "path": pptx.to_string_lossy() }),
        ))
        .expect("document.inspect");
    assert_eq!(inspect.backend.as_deref(), Some("desktop_com"));
    assert_eq!(inspect.changed["summary"]["slide_count"], 3);

    // 3. batch.convert — high-fidelity PDF via COM (§27 criterion 3)
    let pdf_dir = temp.join("pdf");
    let convert = client
        .execute(&command(
            "batch.convert",
            json!({
                "inputs": [pptx.to_string_lossy()],
                "target_format": "pdf",
                "output_directory": pdf_dir.to_string_lossy(),
            }),
        ))
        .expect("batch.convert");
    assert_eq!(convert.backend.as_deref(), Some("desktop_com"));
    assert_eq!(convert.changed["succeeded"], 1);
    assert_eq!(convert.artefacts.len(), 1);
    let pdf = pdf_dir.join("deck.pdf");
    assert!(pdf.exists(), "expected deck.pdf");
    let head = std::fs::read(&pdf).expect("read pdf");
    assert!(head.starts_with(b"%PDF"), "PDF magic mismatch");

    // 4. batch.replace_text — dry-run first, then commit (§27 criterion 4)
    let dry = client
        .execute(&command(
            "batch.replace_text",
            json!({
                "inputs": [pptx.to_string_lossy()],
                "rules": [{ "find": "Checks", "replace": "Verification", "match": "literal" }],
                "scope": ["body"],
                "dry_run": true,
            }),
        ))
        .expect("replace dry-run");
    assert!(dry.changed["total_matched"].as_u64().unwrap_or(0) >= 1);
    assert_eq!(dry.changed["total_replaced"], 0);

    let commit = client
        .execute(&command(
            "batch.replace_text",
            json!({
                "inputs": [pptx.to_string_lossy()],
                "rules": [{ "find": "Checks", "replace": "Verification", "match": "literal" }],
                "scope": ["body"],
                "dry_run": false,
            }),
        ))
        .expect("replace commit");
    assert!(commit.changed["total_replaced"].as_u64().unwrap_or(0) >= 1);

    // 5. slide.render — per-slide PNGs + overflow report (§27 criterion 6)
    let preview_dir = temp.join("previews");
    let render = client
        .execute(&command(
            "slide.render",
            json!({
                "path": pptx.to_string_lossy(),
                "output_directory": preview_dir.to_string_lossy(),
                "width": 640,
                "height": 360,
            }),
        ))
        .expect("slide.render");
    assert_eq!(render.changed["ok"], 3);
    assert_eq!(render.artefacts.len(), 3);
    for artifact in &render.artefacts {
        assert!(
            std::path::Path::new(&artifact.path).exists(),
            "missing {}",
            artifact.path
        );
    }

    // 6. error contract — unknown capability carries the wire error code
    let error = client
        .execute(&command("no.such.capability", json!({})))
        .expect_err("unknown capability must fail");
    match error {
        dcc_mcp_office_client::ClientError::Rpc { code, .. } => {
            assert_eq!(code, json!("OFFICE_CAPABILITY_UNSUPPORTED"));
        }
        other => panic!("expected Rpc error, got {other:?}"),
    }

    guard.shutdown(&mut client);
    drop(client);
    let _ = std::fs::remove_dir_all(&temp);
}

const SAMPLE_DECK_IR: &str = r#"{
  "schema_version": "office-ir/1.0",
  "kind": "presentation",
  "document_id": "draft:contract-test",
  "metadata": {"title": "Contract Test", "language": "zh-CN"},
  "document": {
    "slides": [
      {"semantic_layout": "title_cover", "title": "Contract Test Deck",
       "content_blocks": [{"type": "text", "paragraphs": ["host contract test"]}],
       "speaker_notes": "cover"},
      {"semantic_layout": "bullets", "title": "Checks",
       "content_blocks": [{"type": "bullets", "items": ["compile ok", "inspect ok", "notes ok"]}],
       "speaker_notes": "bullets"},
      {"semantic_layout": "bullets", "title": "Close",
       "content_blocks": [{"type": "bullets", "items": ["done"]}],
       "speaker_notes": "close"}
    ]
  },
  "outputs": ["pptx"]
}"#;

/// Spawns the host for an app, retries connect, returns guard + client.
/// None when the host binary is unavailable (env not set on non-Windows CI).
fn spawn_and_connect(app: &str, pipe: &str) -> Option<(HostGuard, OfficeHostClient)> {
    let exe = host_exe()?;
    let pipe_arg = format!("--pipe-name={pipe}");
    let child = Command::new(&exe)
        .args([format!("--app={app}"), "--pipe".to_string(), pipe_arg])
        .stderr(Stdio::null())
        .spawn()
        .expect("spawn dcc-office-host");
    let mut guard = HostGuard(child);
    let mut client = OfficeHostClient::new(app);
    if !connect_with_retry(&mut client, &mut guard.0, pipe, Duration::from_secs(30)) {
        panic!("host never accepted a connection");
    }
    let handshake = client.handshake("contract-test-0.1.0").expect("handshake");
    assert!(handshake
        .capability_manifest
        .capabilities
        .contains_key("deck.compile"));
    Some((guard, client))
}

/// Converts + replaces text on a fixture through its own app sidecar,
/// skipping when that app's COM backend is unavailable (progressive
/// discovery — the manifest only lists desktop_com capabilities when the
/// app attached).
fn exercise_com_legs(app: &str, pipe: &str, fixture: &std::path::Path, expect_kind: &str) {
    let Some((mut guard, mut client)) = spawn_and_connect(app, pipe) else {
        eprintln!("SKIP: DCC_OFFICE_HOST_EXE not set");
        return;
    };
    let handshake = client.handshake("contract-test-0.1.0").expect("handshake");
    if !handshake
        .capability_manifest
        .capabilities
        .contains_key("batch.convert")
    {
        eprintln!("SKIP: {app} desktop COM not available in this session");
        return;
    }

    let temp =
        std::env::temp_dir().join(format!("dcc-office-contract-{app}-{}", std::process::id()));
    let _ = std::fs::remove_dir_all(&temp);
    std::fs::create_dir_all(&temp).expect("create temp dir");

    let copy = temp.join(fixture.file_name().unwrap());
    std::fs::copy(fixture, &copy).expect("copy fixture");

    // inspect via COM
    let inspect = client
        .execute(&command(
            "document.inspect",
            json!({ "path": copy.to_string_lossy() }),
        ))
        .expect("document.inspect");
    assert_eq!(inspect.backend.as_deref(), Some("desktop_com"));
    assert_eq!(inspect.changed["summary"]["kind"], expect_kind);

    // batch.convert → PDF
    let pdf_dir = temp.join("pdf");
    let convert = client
        .execute(&command(
            "batch.convert",
            json!({
                "inputs": [copy.to_string_lossy()],
                "target_format": "pdf",
                "output_directory": pdf_dir.to_string_lossy(),
            }),
        ))
        .expect("batch.convert");
    assert_eq!(convert.changed["succeeded"], 1);
    let pdf = pdf_dir.join(format!(
        "{}.pdf",
        fixture.file_stem().unwrap().to_string_lossy()
    ));
    let head = std::fs::read(&pdf).expect("read pdf");
    assert!(head.starts_with(b"%PDF"), "PDF magic mismatch");

    // replace dry-run → commit
    let dry = client
        .execute(&command(
            "batch.replace_text",
            json!({
                "inputs": [copy.to_string_lossy()],
                "rules": [{ "find": "2025年度", "replace": "2026年度", "match": "literal" }],
                "scope": ["body"],
                "dry_run": true,
            }),
        ))
        .expect("replace dry-run");
    assert!(dry.changed["total_matched"].as_u64().unwrap_or(0) >= 1);
    assert_eq!(dry.changed["total_replaced"], 0);

    let commit = client
        .execute(&command(
            "batch.replace_text",
            json!({
                "inputs": [copy.to_string_lossy()],
                "rules": [{ "find": "2025年度", "replace": "2026年度", "match": "literal" }],
                "scope": ["body"],
                "dry_run": false,
            }),
        ))
        .expect("replace commit");
    assert!(commit.changed["total_replaced"].as_u64().unwrap_or(0) >= 1);

    guard.shutdown(&mut client);
    drop(client);
    let _ = std::fs::remove_dir_all(&temp);
}

/// Multi-section header/footer coverage: the fixture has two sections whose
/// headers (and one footer) carry the marker; the NextStoryRange ladder must
/// reach them (proposal §15.2 scope "headers"/"footers").
#[test]
fn office_host_word_headers_contract() {
    let fixture = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/fixture-document.docx");
    if !fixture.exists() {
        eprintln!("SKIP: fixture missing: {}", fixture.display());
        return;
    }
    let Some((mut guard, mut client)) =
        spawn_and_connect("word", r"\\.\pipe\dcc-office-contract-test-word-hdr")
    else {
        eprintln!("SKIP: DCC_OFFICE_HOST_EXE not set");
        return;
    };
    let handshake = client.handshake("contract-test-0.1.0").expect("handshake");
    if !handshake
        .capability_manifest
        .capabilities
        .contains_key("batch.convert")
    {
        eprintln!("SKIP: word desktop COM not available in this session");
        return;
    }

    let temp = std::env::temp_dir().join(format!(
        "dcc-office-contract-word-hdr-{}",
        std::process::id()
    ));
    let _ = std::fs::remove_dir_all(&temp);
    std::fs::create_dir_all(&temp).expect("create temp dir");
    let copy = temp.join("fixture-document.docx");
    std::fs::copy(&fixture, &copy).expect("copy fixture");

    let dry = client
        .execute(&command(
            "batch.replace_text",
            json!({
                "inputs": [copy.to_string_lossy()],
                "rules": [{ "find": "2025年度", "replace": "2026年度", "match": "literal" }],
                "scope": ["headers"],
                "dry_run": true,
            }),
        ))
        .expect("headers dry-run");
    // section 1 header + section 2 header (footer excluded by scope)
    assert!(
        dry.changed["total_matched"].as_u64().unwrap_or(0) >= 2,
        "expected the NextStoryRange ladder to reach both section headers: {}",
        dry.changed
    );

    guard.shutdown(&mut client);
    drop(client);
    let _ = std::fs::remove_dir_all(&temp);
}

#[test]
fn office_host_word_contract() {
    let fixture = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/fixture-document.docx");
    if !fixture.exists() {
        eprintln!("SKIP: fixture missing: {}", fixture.display());
        return;
    }
    exercise_com_legs(
        "word",
        r"\\.\pipe\dcc-office-contract-test-word",
        &fixture,
        "document",
    );
}

#[test]
fn office_host_excel_contract() {
    let fixture = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/fixture-workbook.xlsx");
    if !fixture.exists() {
        eprintln!("SKIP: fixture missing: {}", fixture.display());
        return;
    }
    exercise_com_legs(
        "excel",
        r"\\.\pipe\dcc-office-contract-test-excel",
        &fixture,
        "workbook",
    );
}
