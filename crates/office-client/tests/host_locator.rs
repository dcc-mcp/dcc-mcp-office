use std::path::PathBuf;

use dcc_mcp_office_client::{
    locate_host, provider_versions_compatible, HostLocationSource, HostLocatorContext,
    HostLocatorError, OFFICE_HOST_EXE,
};

fn temp_directory(label: &str) -> PathBuf {
    let directory = std::env::temp_dir().join(format!(
        "dcc-office-locator-{label}-{}-{}",
        std::process::id(),
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .expect("system clock")
            .as_nanos()
    ));
    std::fs::create_dir_all(&directory).expect("create temp directory");
    directory
}

fn touch(path: &std::path::Path) {
    std::fs::create_dir_all(path.parent().expect("parent")).expect("create parent");
    std::fs::write(path, b"host-fixture").expect("write fixture");
}

#[test]
fn explicit_environment_override_wins() {
    let root = temp_directory("environment");
    let explicit = root.join("custom-host.exe");
    let gateway = root.join("gateway").join("dcc-mcp-server.exe");
    let sibling = gateway.parent().unwrap().join(OFFICE_HOST_EXE);
    let installed = root
        .join("local")
        .join("dcc-mcp")
        .join("office-host")
        .join("0.2.2")
        .join(OFFICE_HOST_EXE);
    touch(&explicit);
    touch(&sibling);
    touch(&installed);

    let located = locate_host(&HostLocatorContext {
        env_override: Some(explicit.clone()),
        gateway_executable: Some(gateway),
        local_app_data: Some(root.join("local")),
        expected_version: "0.2.2".into(),
    })
    .expect("locate explicit host");

    assert_eq!(located.path, explicit.canonicalize().unwrap());
    assert_eq!(located.source, HostLocationSource::Environment);
    std::fs::remove_dir_all(root).expect("remove temp directory");
}

#[test]
fn configured_missing_path_fails_closed() {
    let root = temp_directory("missing-explicit");
    let gateway = root.join("gateway").join("dcc-mcp-server.exe");
    touch(&gateway.parent().unwrap().join(OFFICE_HOST_EXE));
    let missing = root.join("missing.exe");

    let error = locate_host(&HostLocatorContext {
        env_override: Some(missing.clone()),
        gateway_executable: Some(gateway),
        local_app_data: None,
        expected_version: "0.2.2".into(),
    })
    .expect_err("explicit configuration must not silently fall back");

    assert!(matches!(
        error,
        HostLocatorError::ConfiguredPathMissing(path) if path == missing
    ));
    std::fs::remove_dir_all(root).expect("remove temp directory");
}

#[test]
fn gateway_sibling_precedes_versioned_install() {
    let root = temp_directory("sibling");
    let gateway = root.join("gateway").join("dcc-mcp-server.exe");
    let sibling = gateway.parent().unwrap().join(OFFICE_HOST_EXE);
    let installed = root
        .join("local")
        .join("dcc-mcp")
        .join("office-host")
        .join("0.2.2")
        .join(OFFICE_HOST_EXE);
    touch(&sibling);
    touch(&installed);

    let located = locate_host(&HostLocatorContext {
        env_override: None,
        gateway_executable: Some(gateway),
        local_app_data: Some(root.join("local")),
        expected_version: "0.2.2".into(),
    })
    .expect("locate sibling host");

    assert_eq!(located.path, sibling.canonicalize().unwrap());
    assert_eq!(located.source, HostLocationSource::GatewaySibling);
    std::fs::remove_dir_all(root).expect("remove temp directory");
}

#[test]
fn versioned_install_is_the_final_candidate() {
    let root = temp_directory("installed");
    let installed = root
        .join("dcc-mcp")
        .join("office-host")
        .join("0.2.2")
        .join(OFFICE_HOST_EXE);
    touch(&installed);

    let located = locate_host(&HostLocatorContext {
        env_override: None,
        gateway_executable: None,
        local_app_data: Some(root.clone()),
        expected_version: "0.2.2".into(),
    })
    .expect("locate versioned install");

    assert_eq!(located.path, installed.canonicalize().unwrap());
    assert_eq!(located.source, HostLocationSource::VersionedInstall);
    std::fs::remove_dir_all(root).expect("remove temp directory");
}

#[test]
fn provider_compatibility_follows_pre_one_semver_boundaries() {
    assert!(provider_versions_compatible("0.2.2", "0.2.9"));
    assert!(!provider_versions_compatible("0.2.2", "0.3.0"));
    assert!(provider_versions_compatible("1.4.0", "1.9.1"));
    assert!(!provider_versions_compatible("1.4.0", "2.0.0"));
    assert!(!provider_versions_compatible("not-semver", "0.2.2"));
}
