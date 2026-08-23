//! Cross-language contract test: the Rust client drives a real C# dcc-office-host
//! over the office-rpc/1 named pipe, exercising the full M1 surface:
//!
//!   handshake → deck.compile (Open XML) → document.inspect (COM) →
//!   batch.convert (COM PDF) → batch.replace_text dry-run/commit (COM) →
//!   slide.render (COM previews) → error contract.
//!
//! Requirements:
//!   - DCC_OFFICE_HOST_EXE points at the built dcc-office-host.exe.
//!   - Hosted CI explicitly runs the deterministic `--openxml-only` case.
//!   - Desktop COM cases are visible `#[ignore]` tests and fail when invoked
//!     without their required PowerPoint, Word, or Excel installation.
#![cfg(windows)]

use std::path::PathBuf;
use std::process::{Child, Command, Stdio};
use std::sync::{Mutex, MutexGuard};
use std::time::{Duration, Instant};

use dcc_mcp_office_client::OfficeHostClient;
use dcc_mcp_office_protocol::{CommandParams, CommandResult, ConfirmationProof, JobPhase};
use serde_json::json;

const PIPE: &str = r"\\.\pipe\dcc-office-contract-test-powerpoint";
const OPENXML_PIPE: &str = r"\\.\pipe\dcc-office-contract-test-openxml";
static DESKTOP_OFFICE_LOCK: Mutex<()> = Mutex::new(());

fn desktop_office_lock() -> MutexGuard<'static, ()> {
    DESKTOP_OFFICE_LOCK
        .lock()
        .unwrap_or_else(std::sync::PoisonError::into_inner)
}

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

fn connect_or_panic(
    client: &mut OfficeHostClient,
    child: &mut Child,
    pipe: &str,
    timeout: Duration,
) {
    if let Err(error) = client.connect_with_retry(pipe, timeout) {
        if let Ok(Some(status)) = child.try_wait() {
            panic!("host exited before accepting a connection: {status}: {error}");
        }
        panic!("host never accepted a connection: {error}");
    }
}

fn command(capability: &str, input: serde_json::Value) -> CommandParams {
    CommandParams {
        capability: capability.to_string(),
        document: None,
        confirmation: None,
        input,
        policy: json!({
            "workspace_root": std::env::temp_dir().to_string_lossy()
        }),
    }
}

fn confirmed_command(capability: &str, input: serde_json::Value) -> CommandParams {
    let mut params = command(capability, input);
    params.confirmation = Some(ConfirmationProof {
        action: "overwrite_original".into(),
        confirmed: true,
        confirmed_by: "human:contract-test".into(),
        confirmed_at: "2026-08-23T14:00:00Z".into(),
    });
    params
}

fn execute_batch(client: &mut OfficeHostClient, params: &CommandParams) -> CommandResult {
    let submission = client.execute(params).expect("submit batch job");
    assert_eq!(submission.backend.as_deref(), Some("job"));
    let job_id = submission.job_id.expect("batch submission job_id");
    let deadline = Instant::now() + Duration::from_secs(180);
    let terminal = loop {
        let status = client.job_get(&job_id).expect("poll batch job");
        if status.phase.is_terminal() {
            break status;
        }
        assert!(Instant::now() < deadline, "batch job {job_id} timed out");
        std::thread::sleep(Duration::from_millis(20));
    };
    let notifications = collect_job_notifications(client, &job_id, Duration::from_secs(5));
    assert!(
        notifications.iter().any(|message| {
            message["method"] == "office.job.progress" && message["params"]["job_id"] == job_id
        }),
        "job {job_id} emitted no progress notification"
    );
    assert!(
        notifications.iter().any(|message| {
            message["method"] == "office.job.completed"
                && message["params"]["correlation_id"] == job_id
        }),
        "job {job_id} emitted no completion event"
    );
    assert!(
        matches!(
            terminal.phase,
            JobPhase::Succeeded | JobPhase::PartiallySucceeded
        ),
        "job {job_id} ended as {:?}: {:?}",
        terminal.phase,
        terminal.error
    );
    serde_json::from_value(terminal.result.expect("successful batch result"))
        .expect("deserialize terminal command result")
}

fn collect_job_notifications(
    client: &mut OfficeHostClient,
    job_id: &str,
    timeout: Duration,
) -> Vec<serde_json::Value> {
    let deadline = Instant::now() + timeout;
    let mut notifications = Vec::new();
    loop {
        // A request/response read boundary lets the client consume any
        // notifications already written by the independent Host pump.
        let _ = client.ping().expect("flush job notifications");
        notifications.extend(client.drain_notifications());
        if notifications.iter().any(|message| {
            message["method"] == "office.job.completed"
                && message["params"]["correlation_id"] == job_id
        }) {
            return notifications;
        }
        assert!(
            Instant::now() < deadline,
            "job {job_id} emitted no completion event"
        );
        std::thread::sleep(Duration::from_millis(20));
    }
}

#[test]
#[ignore = "requires a built dcc-office-host; CI runs this Office-free contract explicitly"]
fn office_host_openxml_contract() {
    let exe = host_exe().expect("DCC_OFFICE_HOST_EXE must point to a built dcc-office-host.exe");
    let pipe_arg = format!("--pipe-name={OPENXML_PIPE}");
    let workspace_arg = format!("--workspace-root={}", std::env::temp_dir().display());
    let child = Command::new(&exe)
        .args([
            "--app=powerpoint",
            "--pipe",
            "--openxml-only",
            &pipe_arg,
            &workspace_arg,
        ])
        .stderr(Stdio::null())
        .spawn()
        .expect("spawn Office-free dcc-office-host");
    let mut guard = HostGuard(child);
    let mut client = OfficeHostClient::new("powerpoint");
    connect_or_panic(
        &mut client,
        &mut guard.0,
        OPENXML_PIPE,
        Duration::from_secs(30),
    );

    let handshake = client.handshake("contract-test-0.1.0").expect("handshake");
    assert_eq!(handshake.protocol_version, "office-rpc/1");
    assert_eq!(
        handshake.capability_manifest.provider_version,
        env!("CARGO_PKG_VERSION")
    );
    assert_eq!(
        handshake.capability_manifest.execution_modes,
        vec!["openxml"]
    );
    assert!(!handshake
        .capability_manifest
        .capabilities
        .contains_key("batch.convert"));
    let status = client.status().expect("typed sidecar status");
    assert_eq!(status.app, "powerpoint");
    assert_eq!(status.com_attach_state, "unavailable");
    assert!(!status.busy);

    let temp = std::env::temp_dir().join(format!("dcc-office-openxml-{}", std::process::id()));
    let _ = std::fs::remove_dir_all(&temp);
    std::fs::create_dir_all(&temp).expect("create temp dir");
    let pptx = temp.join("deck.pptx");
    let compiled = client
        .execute(&command(
            "deck.compile",
            json!({ "ir": SAMPLE_DECK_IR, "output": pptx.to_string_lossy() }),
        ))
        .expect("deck.compile");
    assert_eq!(compiled.backend.as_deref(), Some("openxml"));
    assert_eq!(compiled.changed["slides"], 3);
    assert_eq!(
        compiled.audit.as_ref().expect("audit")["host_version"],
        env!("CARGO_PKG_VERSION")
    );

    let inspected = client
        .execute(&command(
            "document.inspect",
            json!({ "path": pptx.to_string_lossy(), "backend": "openxml" }),
        ))
        .expect("document.inspect");
    assert_eq!(inspected.backend.as_deref(), Some("openxml"));
    assert_eq!(inspected.changed["summary"]["slide_count"], 3);

    // Even when desktop COM is unavailable, a valid long-running request is
    // accepted immediately and its failure is observable through office.job.get.
    let submission = client
        .execute(&command(
            "batch.convert",
            json!({
                "inputs": [pptx.to_string_lossy()],
                "output_directory": temp.join("pdf").to_string_lossy(),
            }),
        ))
        .expect("submit Office-free batch");
    assert_eq!(submission.backend.as_deref(), Some("job"));
    let job_id = submission.job_id.expect("job_id");
    let deadline = Instant::now() + Duration::from_secs(10);
    let terminal = loop {
        let status = client.job_get(&job_id).expect("poll Office-free batch");
        if status.phase.is_terminal() {
            break status;
        }
        assert!(Instant::now() < deadline, "Office-free batch timed out");
        std::thread::sleep(Duration::from_millis(20));
    };
    assert_eq!(terminal.phase, JobPhase::Failed);
    assert_eq!(
        terminal.error.expect("terminal error").code,
        "OFFICE_BACKEND_UNAVAILABLE"
    );
    assert!(
        collect_job_notifications(&mut client, &job_id, Duration::from_secs(5))
            .iter()
            .any(|message| {
                message["method"] == "office.job.completed"
                    && message["params"]["correlation_id"] == job_id
            })
    );

    let denied = CommandParams {
        capability: "batch.replace_text".to_string(),
        document: None,
        confirmation: None,
        input: json!({}),
        policy: json!({ "macros": "confirm" }),
    };
    match client.execute(&denied) {
        Err(dcc_mcp_office_client::ClientError::Rpc { code, .. }) => {
            assert_eq!(code, json!("OFFICE_MACRO_BLOCKED"));
        }
        other => panic!("macro policy relaxation must fail, got {other:?}"),
    }

    guard.shutdown(&mut client);
    drop(client);
    let _ = std::fs::remove_dir_all(&temp);
}

#[test]
#[ignore = "requires Microsoft PowerPoint desktop; run the explicit real-Office lane"]
fn office_host_full_contract() {
    let _desktop_office = desktop_office_lock();
    let exe = host_exe().expect("DCC_OFFICE_HOST_EXE must point to dcc-office-host.exe");

    let pipe_arg = format!("--pipe-name={PIPE}");
    let workspace_arg = format!("--workspace-root={}", std::env::temp_dir().display());
    let child = Command::new(&exe)
        .args(["--app=powerpoint", "--pipe", &pipe_arg, &workspace_arg])
        .stderr(Stdio::null())
        .spawn()
        .expect("spawn dcc-office-host");
    let mut guard = HostGuard(child);

    let mut client = OfficeHostClient::new("powerpoint");
    connect_or_panic(&mut client, &mut guard.0, PIPE, Duration::from_secs(30));
    let handshake = client.handshake("contract-test-0.1.0").expect("handshake");
    assert_eq!(handshake.protocol_version, "office-rpc/1");
    assert_eq!(
        handshake.capability_manifest.provider_version,
        env!("CARGO_PKG_VERSION")
    );
    assert!(handshake
        .capability_manifest
        .capabilities
        .contains_key("deck.compile"));
    assert_eq!(client.state(), dcc_mcp_office_protocol::SidecarState::Ready);

    let ping = client.ping().expect("ping");
    assert_eq!(ping["app"], "powerpoint");

    assert!(
        handshake
            .capability_manifest
            .capabilities
            .contains_key("batch.convert"),
        "PowerPoint desktop COM is required for this ignored contract"
    );

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
    assert_eq!(audit["security"]["automation_security"]["observed"], 3);
    assert_eq!(audit["security"]["automation_security"]["enforced"], true);
    assert_eq!(audit["backend"], "openxml");
    assert_eq!(audit["host_version"], env!("CARGO_PKG_VERSION"));
    assert!(audit["duration_ms"].as_u64().unwrap_or(0) > 0 || audit["duration_ms"].is_null());

    // 1c. §19 second-layer policy gate: relaxing deny-by-default is refused
    let policy = CommandParams {
        capability: "batch.replace_text".to_string(),
        document: None,
        confirmation: None,
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
        confirmation: None,
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
    std::fs::create_dir_all(&pdf_dir).expect("create PDF directory");
    let pdf = pdf_dir.join("deck.pdf");
    let previous_pdf = b"previous output";
    std::fs::write(&pdf, previous_pdf).expect("seed existing PDF");
    let convert = execute_batch(
        &mut client,
        &confirmed_command(
            "batch.convert",
            json!({
                "inputs": [pptx.to_string_lossy()],
                "target_format": "pdf",
                "output_directory": pdf_dir.to_string_lossy(),
                "overwrite": "overwrite",
            }),
        ),
    );
    assert_eq!(convert.backend.as_deref(), Some("desktop_com"));
    assert_eq!(convert.changed["succeeded"], 1);
    let checkpoint = convert
        .artefacts
        .iter()
        .find(|artifact| artifact.kind == "checkpoint")
        .expect("overwritten PDF checkpoint");
    assert_eq!(
        std::fs::read(&checkpoint.path).expect("read PDF checkpoint"),
        previous_pdf
    );
    assert!(pdf.exists(), "expected deck.pdf");
    let head = std::fs::read(&pdf).expect("read pdf");
    assert!(head.starts_with(b"%PDF"), "PDF magic mismatch");

    // 4. batch.replace_text — dry-run first, then commit (§27 criterion 4)
    let dry = execute_batch(
        &mut client,
        &command(
            "batch.replace_text",
            json!({
                "inputs": [pptx.to_string_lossy()],
                "rules": [{ "find": "Checks", "replace": "Verification", "match": "literal" }],
                "scope": ["body"],
                "dry_run": true,
            }),
        ),
    );
    assert!(dry.changed["total_matched"].as_u64().unwrap_or(0) >= 1);
    assert_eq!(dry.changed["total_replaced"], 0);

    let commit = execute_batch(
        &mut client,
        &confirmed_command(
            "batch.replace_text",
            json!({
                "inputs": [pptx.to_string_lossy()],
                "rules": [{ "find": "Checks", "replace": "Verification", "match": "literal" }],
                "scope": ["body"],
                "dry_run": false,
            }),
        ),
    );
    assert!(commit.changed["total_replaced"].as_u64().unwrap_or(0) >= 1);
    assert!(commit
        .artefacts
        .iter()
        .any(|artifact| artifact.kind == "checkpoint"));

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

/// Spawns the host for an app and retries connect. Ignored COM contracts fail
/// loudly when their explicitly provisioned host is unavailable.
fn spawn_and_connect(app: &str, pipe: &str) -> (HostGuard, OfficeHostClient) {
    let exe = host_exe().expect("DCC_OFFICE_HOST_EXE must point to dcc-office-host.exe");
    let pipe_arg = format!("--pipe-name={pipe}");
    let workspace_arg = format!("--workspace-root={}", std::env::temp_dir().display());
    let child = Command::new(&exe)
        .args([
            format!("--app={app}"),
            "--pipe".to_string(),
            pipe_arg,
            workspace_arg,
        ])
        .stderr(Stdio::null())
        .spawn()
        .expect("spawn dcc-office-host");
    let mut guard = HostGuard(child);
    let mut client = OfficeHostClient::new(app);
    connect_or_panic(&mut client, &mut guard.0, pipe, Duration::from_secs(30));
    let handshake = client.handshake("contract-test-0.1.0").expect("handshake");
    assert!(
        !handshake.capability_manifest.capabilities.is_empty(),
        "{app} host advertised no catalog capabilities"
    );
    (guard, client)
}

/// Converts + replaces text on a fixture through its own app sidecar. The
/// ignored real-Office lane treats a missing desktop_com capability as a
/// failure rather than a passing early return.
fn exercise_com_legs(app: &str, pipe: &str, fixture: &std::path::Path, expect_kind: &str) {
    let (mut guard, mut client) = spawn_and_connect(app, pipe);
    let handshake = client.handshake("contract-test-0.1.0").expect("handshake");
    assert!(
        handshake
            .capability_manifest
            .capabilities
            .contains_key("batch.convert"),
        "{app} desktop COM is required for this ignored contract"
    );

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
    std::fs::create_dir_all(&pdf_dir).expect("create PDF directory");
    let pdf = pdf_dir.join(format!(
        "{}.pdf",
        fixture.file_stem().unwrap().to_string_lossy()
    ));
    let previous_pdf = b"previous output";
    std::fs::write(&pdf, previous_pdf).expect("seed existing PDF");
    let convert = execute_batch(
        &mut client,
        &confirmed_command(
            "batch.convert",
            json!({
                "inputs": [copy.to_string_lossy()],
                "target_format": "pdf",
                "output_directory": pdf_dir.to_string_lossy(),
                "overwrite": "overwrite",
            }),
        ),
    );
    assert_eq!(convert.changed["succeeded"], 1);
    let checkpoint = convert
        .artefacts
        .iter()
        .find(|artifact| artifact.kind == "checkpoint")
        .expect("overwritten PDF checkpoint");
    assert_eq!(
        std::fs::read(&checkpoint.path).expect("read PDF checkpoint"),
        previous_pdf
    );
    let head = std::fs::read(&pdf).expect("read pdf");
    assert!(head.starts_with(b"%PDF"), "PDF magic mismatch");

    // replace dry-run → commit
    let dry = execute_batch(
        &mut client,
        &command(
            "batch.replace_text",
            json!({
                "inputs": [copy.to_string_lossy()],
                "rules": [{ "find": "2025年度", "replace": "2026年度", "match": "literal" }],
                "scope": ["body"],
                "dry_run": true,
            }),
        ),
    );
    assert!(dry.changed["total_matched"].as_u64().unwrap_or(0) >= 1);
    assert_eq!(dry.changed["total_replaced"], 0);

    let commit = execute_batch(
        &mut client,
        &confirmed_command(
            "batch.replace_text",
            json!({
                "inputs": [copy.to_string_lossy()],
                "rules": [{ "find": "2025年度", "replace": "2026年度", "match": "literal" }],
                "scope": ["body"],
                "dry_run": false,
            }),
        ),
    );
    assert!(commit.changed["total_replaced"].as_u64().unwrap_or(0) >= 1);
    assert!(commit
        .artefacts
        .iter()
        .any(|artifact| artifact.kind == "checkpoint"));

    guard.shutdown(&mut client);
    drop(client);
    let _ = std::fs::remove_dir_all(&temp);
}

/// Multi-section header/footer coverage: the fixture has two sections whose
/// headers (and one footer) carry the marker; the NextStoryRange ladder must
/// reach them (proposal §15.2 scope "headers"/"footers").
#[test]
#[ignore = "requires Microsoft Word desktop; run the explicit real-Office lane"]
fn office_host_word_headers_contract() {
    let _desktop_office = desktop_office_lock();
    let fixture = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/fixture-document.docx");
    assert!(fixture.exists(), "fixture missing: {}", fixture.display());
    let (mut guard, mut client) =
        spawn_and_connect("word", r"\\.\pipe\dcc-office-contract-test-word-hdr");
    let handshake = client.handshake("contract-test-0.1.0").expect("handshake");
    assert!(
        handshake
            .capability_manifest
            .capabilities
            .contains_key("batch.convert"),
        "Word desktop COM is required for this ignored contract"
    );

    let temp = std::env::temp_dir().join(format!(
        "dcc-office-contract-word-hdr-{}",
        std::process::id()
    ));
    let _ = std::fs::remove_dir_all(&temp);
    std::fs::create_dir_all(&temp).expect("create temp dir");
    let copy = temp.join("fixture-document.docx");
    std::fs::copy(&fixture, &copy).expect("copy fixture");

    let dry = execute_batch(
        &mut client,
        &command(
            "batch.replace_text",
            json!({
                "inputs": [copy.to_string_lossy()],
                "rules": [{ "find": "2025年度", "replace": "2026年度", "match": "literal" }],
                "scope": ["headers"],
                "dry_run": true,
            }),
        ),
    );
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
#[ignore = "requires Microsoft Word desktop; run the explicit real-Office lane"]
fn office_host_word_contract() {
    let _desktop_office = desktop_office_lock();
    let fixture = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/fixture-document.docx");
    assert!(fixture.exists(), "fixture missing: {}", fixture.display());
    exercise_com_legs(
        "word",
        r"\\.\pipe\dcc-office-contract-test-word",
        &fixture,
        "document",
    );
}

#[test]
#[ignore = "requires Microsoft Excel desktop; run the explicit real-Office lane"]
fn office_host_excel_contract() {
    let _desktop_office = desktop_office_lock();
    let fixture = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/fixture-workbook.xlsx");
    assert!(fixture.exists(), "fixture missing: {}", fixture.display());
    exercise_com_legs(
        "excel",
        r"\\.\pipe\dcc-office-contract-test-excel",
        &fixture,
        "workbook",
    );
}
