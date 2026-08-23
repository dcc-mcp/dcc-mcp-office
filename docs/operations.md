# Office Host operations

The Host is configured in four deterministic layers, with later layers
winning: built-in defaults, optional `dcc-office-host.json` beside the
executable, `DCC_OFFICE_*` environment variables, then command-line flags.
Use `--config=<path>` to require a particular settings file. Invalid explicit
configuration fails before the pipe or Office starts.

| Setting | JSON | Environment | CLI |
|---|---|---|---|
| COM attach timeout | `attach_timeout_seconds` | `DCC_OFFICE_ATTACH_TIMEOUT_SECONDS` | `--attach-timeout-seconds=<n>` |
| request soft timeout | `request_timeout_seconds` | `DCC_OFFICE_REQUEST_TIMEOUT_SECONDS` | `--request-timeout-seconds=<n>` |
| timeout recovery streak | `recovery_timeout_streak` | `DCC_OFFICE_RECOVERY_TIMEOUT_STREAK` | `--recovery-timeout-streak=<n>` |
| COM busy retry count | `busy_retry_count` | `DCC_OFFICE_BUSY_RETRY_COUNT` | `--busy-retry-count=<n>` |
| pipe buffer bytes | `pipe_buffer_bytes` | `DCC_OFFICE_PIPE_BUFFER_BYTES` | `--pipe-buffer-bytes=<n>` |
| log level | `log_level` | `DCC_OFFICE_LOG_LEVEL` | `--log-level=<level>` |
| optional log file | `log_path` | `DCC_OFFICE_LOG_PATH` | `--log-path=<path>` |
| template roots | `template_directories` | `DCC_OFFICE_TEMPLATE_DIRS` | `--template-dir=<path>` |

The JSON schema and a copyable example live under `manifests/`. Config-file
template paths are relative to that file; environment and CLI paths are
relative to the launching process unless absolute. Template options append so
operators can combine managed and per-task roots.

## Lifecycle

`office.host.ping` and `office.host.handshake` never launch Office. The
handshake performs a registry-only installation probe and advertises
`desktop_com` when installed; the first command that actually needs COM starts
the application. This avoids the empty dark PowerPoint window previously
created during every handshake. `com_attach_state=available` means installed
but not started; `attached` means the live security posture was applied and
read back.

Pass `--parent-pid=<gateway-pid>` when the gateway owns the sidecar. If that
process exits, the Host cancels its pipe loop, disposes the COM backend, and
quits its Office instance. Ctrl+C uses the same cancellation path. Named-pipe
accept and reads are asynchronous, so an idle process stops promptly.

## Diagnostics

Every JSON-RPC request emits one JSON line to stderr with timestamp, level,
application, PID, correlation ID, capability, duration, outcome code, and the
operation ID when available. Request payloads and document contents are never
logged. `log_path` adds an append-only JSONL file while retaining stderr.

The Host logs startup and pipe failures in the same shape. Use the
`correlation_id` to join a request with events and the `operation_id` to join
successful command results.

## Input globs

`*` and `?` match only the directory named by the fixed prefix. Recursive
walking occurs only when the specification explicitly contains `**`.
Inaccessible directories and unmatched literal/glob inputs are reported as
warnings; one protected descendant no longer aborts an entire batch.
