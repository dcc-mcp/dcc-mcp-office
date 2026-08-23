use std::fs;
use std::path::{Path, PathBuf};
#[cfg(windows)]
use std::process::Command;

use serde_json::Value;

const EXPECTED_VERSION: &str = env!("CARGO_PKG_VERSION");

fn repository_root() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .ancestors()
        .nth(2)
        .expect("office-client must live under <repo>/crates")
        .to_path_buf()
}

fn text_between<'a>(text: &'a str, start: &str, end: &str) -> &'a str {
    text.split_once(start)
        .and_then(|(_, rest)| rest.split_once(end))
        .map(|(value, _)| value.trim())
        .unwrap_or_else(|| panic!("missing value between {start:?} and {end:?}"))
}

#[test]
fn repository_versions_stay_in_lockstep() {
    let root = repository_root();

    let manifest: Value = serde_json::from_str(
        &fs::read_to_string(root.join(".release-please-manifest.json"))
            .expect("read release-please manifest"),
    )
    .expect("parse release-please manifest");
    assert_eq!(manifest["."], EXPECTED_VERSION);

    let props = fs::read_to_string(root.join("dotnet/Directory.Build.props"))
        .expect("read Directory.Build.props");
    assert_eq!(
        text_between(&props, "<Version>", "</Version>"),
        EXPECTED_VERSION,
        ".NET assembly metadata must match the Rust workspace version"
    );

    let release_config: Value = serde_json::from_str(
        &fs::read_to_string(root.join("release-please-config.json"))
            .expect("read release-please config"),
    )
    .expect("parse release-please config");
    let extra_files = release_config["packages"]["."]["extra-files"]
        .as_array()
        .expect("release-please extra-files array");
    assert!(
        extra_files.iter().any(|entry| {
            entry["type"] == "generic" && entry["path"] == "dotnet/Directory.Build.props"
        }),
        "release-please must stamp dotnet/Directory.Build.props"
    );

    let lock = fs::read_to_string(root.join("Cargo.lock")).expect("read Cargo.lock");
    for package in lock.split("[[package]]").skip(1) {
        let name = text_between(package, "name = \"", "\"");
        if name.starts_with("dcc-mcp-office-") {
            assert_eq!(
                text_between(package, "version = \"", "\""),
                EXPECTED_VERSION,
                "Cargo.lock entry for {name} must match the workspace version"
            );
        }
    }
}

#[cfg(windows)]
#[test]
fn office_host_version_is_machine_readable() {
    let Some(exe) = std::env::var_os("DCC_OFFICE_HOST_EXE") else {
        eprintln!("SKIP: DCC_OFFICE_HOST_EXE not set");
        return;
    };
    let output = Command::new(exe)
        .arg("--version")
        .output()
        .expect("run dcc-office-host --version");
    assert!(output.status.success(), "--version must exit successfully");
    assert_eq!(
        String::from_utf8(output.stdout)
            .expect("--version stdout must be UTF-8")
            .trim(),
        format!("dcc-office-host {EXPECTED_VERSION} (office-rpc/1)")
    );
    assert!(output.stderr.is_empty(), "--version must not write stderr");
}
