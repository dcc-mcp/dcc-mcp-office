#![cfg(windows)]

use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::mpsc;
use std::thread;
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};

use dcc_mcp_office_client::{ClientError, OfficeHostClient};
use dcc_mcp_office_protocol::{CommandParams, SidecarState};
use serde_json::{json, Value};
use tokio::io::{AsyncBufReadExt, AsyncReadExt, AsyncWriteExt, BufReader};
use tokio::net::windows::named_pipe::{NamedPipeServer, ServerOptions};
use tokio::runtime::Builder;

static PIPE_SEQUENCE: AtomicU64 = AtomicU64::new(0);

fn unique_pipe(label: &str) -> String {
    let stamp = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system clock")
        .as_nanos();
    let sequence = PIPE_SEQUENCE.fetch_add(1, Ordering::Relaxed);
    format!(
        r"\\.\pipe\dcc-office-client-{label}-{}-{stamp}-{sequence}",
        std::process::id()
    )
}

fn runtime() -> tokio::runtime::Runtime {
    Builder::new_current_thread()
        .enable_io()
        .enable_time()
        .build()
        .expect("Tokio runtime")
}

async fn read_request(server: &mut BufReader<NamedPipeServer>) -> Value {
    let mut line = String::new();
    let read = server.read_line(&mut line).await.expect("read request");
    assert!(read > 0, "client closed before sending a request");
    serde_json::from_str(line.trim_end()).expect("JSON-RPC request")
}

async fn try_write_message(
    server: &mut BufReader<NamedPipeServer>,
    message: &Value,
) -> std::io::Result<()> {
    let writer = server.get_mut();
    writer.write_all(message.to_string().as_bytes()).await?;
    writer.write_all(b"\n").await?;
    writer.flush().await
}

async fn write_message(server: &mut BufReader<NamedPipeServer>, message: &Value) {
    try_write_message(server, message)
        .await
        .expect("write response");
}

fn spawn_scripted_server(pipe: &str, messages: Vec<Value>) -> thread::JoinHandle<()> {
    let pipe = pipe.to_string();
    let (ready_tx, ready_rx) = mpsc::sync_channel(1);
    let handle = thread::spawn(move || {
        runtime().block_on(async move {
            let server = ServerOptions::new()
                .first_pipe_instance(true)
                .create(&pipe)
                .expect("create named-pipe server");
            ready_tx.send(()).expect("signal server ready");
            server.connect().await.expect("accept client");
            let mut server = BufReader::new(server);
            let _ = read_request(&mut server).await;
            for message in messages {
                if try_write_message(&mut server, &message).await.is_err() {
                    break;
                }
            }
            let mut remainder = Vec::new();
            let _ = server.read_to_end(&mut remainder).await;
        });
    });
    ready_rx.recv().expect("server ready");
    handle
}

#[test]
fn connect_with_retry_waits_for_a_server_that_is_starting() {
    let pipe = unique_pipe("connect-retry");
    let server_pipe = pipe.clone();
    let server = thread::spawn(move || {
        thread::sleep(Duration::from_millis(150));
        runtime().block_on(async move {
            let server = ServerOptions::new()
                .first_pipe_instance(true)
                .create(&server_pipe)
                .expect("create delayed server");
            server.connect().await.expect("accept client");
            let mut server = BufReader::new(server);
            let mut remainder = Vec::new();
            let _ = server.read_to_end(&mut remainder).await;
        });
    });

    let mut client = OfficeHostClient::new("powerpoint");
    client
        .connect_with_retry(&pipe, Duration::from_secs(2))
        .expect("client retries until server appears");
    assert_eq!(client.state(), SidecarState::Handshaking);

    drop(client);
    server.join().expect("server thread");
}

#[test]
fn call_timeout_is_bounded_and_degrades_the_connection() {
    let pipe = unique_pipe("timeout");
    let server = spawn_scripted_server(&pipe, Vec::new());
    let mut client = OfficeHostClient::new("powerpoint");
    client
        .connect_with_retry(&pipe, Duration::from_secs(1))
        .expect("connect");

    let started = Instant::now();
    let error = client
        .ping_with_timeout(Duration::from_millis(100))
        .expect_err("silent server must time out");

    assert!(matches!(error, ClientError::Timeout { .. }));
    assert!(started.elapsed() < Duration::from_secs(1));
    assert_eq!(client.state(), SidecarState::Degraded);
    server.join().expect("server observes client close");
}

#[test]
fn write_timeout_is_bounded_when_the_server_never_reads() {
    let pipe = unique_pipe("write-timeout");
    let server_pipe = pipe.clone();
    let (ready_tx, ready_rx) = mpsc::sync_channel(1);
    let server = thread::spawn(move || {
        runtime().block_on(async move {
            let server = ServerOptions::new()
                .first_pipe_instance(true)
                .create(&server_pipe)
                .expect("create named-pipe server");
            ready_tx.send(()).expect("signal server ready");
            server.connect().await.expect("accept client");
            tokio::time::sleep(Duration::from_millis(500)).await;
        });
    });
    ready_rx.recv().expect("server ready");

    let mut client = OfficeHostClient::new("powerpoint");
    client
        .connect_with_retry(&pipe, Duration::from_secs(1))
        .expect("connect");
    let params = CommandParams {
        capability: "deck.compile".into(),
        document: None,
        input: json!({"payload": "x".repeat(8 * 1024 * 1024)}),
        policy: json!({}),
    };

    let started = Instant::now();
    let error = client
        .execute_with_timeout(&params, Duration::from_millis(100))
        .expect_err("non-reading server must time out the write");

    assert!(matches!(error, ClientError::Timeout { .. }));
    assert!(started.elapsed() < Duration::from_secs(1));
    assert_eq!(client.state(), SidecarState::Degraded);
    server.join().expect("server thread");
}

#[test]
fn notifications_are_buffered_while_waiting_for_the_matching_response() {
    let pipe = unique_pipe("notifications");
    let server = spawn_scripted_server(
        &pipe,
        vec![
            json!({
                "jsonrpc": "2.0",
                "method": "office.job.progress",
                "params": {"job_id": "job-1", "progress": 0.5}
            }),
            json!({"jsonrpc": "2.0", "id": 99, "result": {"ignored": true}}),
            json!({"jsonrpc": "2.0", "id": 0, "result": {"ok": true}}),
        ],
    );
    let mut client = OfficeHostClient::new("powerpoint");
    client
        .connect_with_retry(&pipe, Duration::from_secs(1))
        .expect("connect");

    let result = client
        .ping_with_timeout(Duration::from_secs(1))
        .expect("matching response");
    assert_eq!(result["ok"], true);

    let notifications = client.drain_notifications();
    assert_eq!(notifications.len(), 1);
    assert_eq!(notifications[0]["method"], "office.job.progress");
    assert!(client.drain_notifications().is_empty());

    drop(client);
    server.join().expect("server thread");
}

#[test]
fn synchronous_client_is_safe_inside_an_existing_tokio_runtime() {
    let pipe = unique_pipe("nested-runtime");
    let server = spawn_scripted_server(
        &pipe,
        vec![json!({"jsonrpc": "2.0", "id": 0, "result": {"ok": true}})],
    );
    let mut client = OfficeHostClient::new("powerpoint");
    client
        .connect_with_retry(&pipe, Duration::from_secs(1))
        .expect("connect");

    let result = runtime().block_on(async { client.ping_with_timeout(Duration::from_secs(1)) });

    assert_eq!(result.expect("nested runtime call")["ok"], true);
    drop(client);
    server.join().expect("server thread");
}

#[test]
fn repeated_wrong_response_ids_fail_as_a_protocol_error() {
    let pipe = unique_pipe("wrong-id");
    let messages = (0..16)
        .map(|offset| json!({"jsonrpc": "2.0", "id": 100 + offset, "result": null}))
        .collect();
    let server = spawn_scripted_server(&pipe, messages);
    let mut client = OfficeHostClient::new("powerpoint");
    client
        .connect_with_retry(&pipe, Duration::from_secs(1))
        .expect("connect");

    let error = client
        .ping_with_timeout(Duration::from_secs(1))
        .expect_err("wrong ids must be bounded");
    assert!(matches!(error, ClientError::Protocol(_)));
    assert_eq!(client.state(), SidecarState::Degraded);

    drop(client);
    server.join().expect("server thread");
}

#[test]
fn disconnected_client_reconnects_and_repeats_the_handshake() {
    let pipe = unique_pipe("recover");
    let server_pipe = pipe.clone();
    let (ready_tx, ready_rx) = mpsc::sync_channel(1);
    let server = thread::spawn(move || {
        runtime().block_on(async move {
            let first_server = ServerOptions::new()
                .first_pipe_instance(true)
                .create(&server_pipe)
                .expect("create first server");
            ready_tx.send(()).expect("signal first server ready");
            first_server.connect().await.expect("accept first client");
            let mut first = BufReader::new(first_server);
            let handshake = read_request(&mut first).await;
            write_message(
                &mut first,
                &json!({
                    "jsonrpc": "2.0",
                    "id": handshake["id"],
                    "result": {
                        "protocol_version": "office-rpc/1",
                        "host_id": "first-host",
                        "capability_manifest": {}
                    }
                }),
            )
            .await;
            let _ = read_request(&mut first).await;
            drop(first);

            let replacement = loop {
                match ServerOptions::new().create(&server_pipe) {
                    Ok(server) => break server,
                    Err(error) if error.raw_os_error() == Some(231) => {
                        tokio::time::sleep(Duration::from_millis(20)).await;
                    }
                    Err(error) => panic!("create replacement server: {error}"),
                }
            };
            replacement.connect().await.expect("accept recovery client");
            let mut replacement = BufReader::new(replacement);
            let handshake = read_request(&mut replacement).await;
            write_message(
                &mut replacement,
                &json!({
                    "jsonrpc": "2.0",
                    "id": handshake["id"],
                    "result": {
                        "protocol_version": "office-rpc/1",
                        "host_id": "replacement-host",
                        "capability_manifest": {}
                    }
                }),
            )
            .await;
            let mut remainder = Vec::new();
            let _ = replacement.read_to_end(&mut remainder).await;
        });
    });
    ready_rx.recv().expect("first server ready");

    let mut client = OfficeHostClient::new("powerpoint");
    client
        .connect_with_retry(&pipe, Duration::from_secs(1))
        .expect("connect");
    client.handshake("test-gateway").expect("first handshake");
    assert_eq!(client.state(), SidecarState::Ready);

    let error = client
        .ping_with_timeout(Duration::from_secs(1))
        .expect_err("first server disconnects");
    assert!(matches!(error, ClientError::Disconnected));
    assert_eq!(client.state(), SidecarState::Degraded);

    let recovered = client
        .recover(Duration::from_secs(2))
        .expect("reconnect and handshake");
    assert_eq!(recovered.host_id, "replacement-host");
    assert_eq!(client.state(), SidecarState::Ready);

    drop(client);
    server.join().expect("server thread");
}
