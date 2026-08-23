use std::collections::BTreeSet;
use std::fs;
use std::path::Path;

use dcc_mcp_office_protocol::{capability_catalog, OfficeErrorCode, PROTOCOL_VERSION};

#[test]
fn catalog_is_the_complete_wire_contract() {
    let catalog = capability_catalog();
    assert_eq!(catalog.schema_version, "office-capability-catalog/1.2");
    assert_eq!(catalog.protocol_version, PROTOCOL_VERSION);
    assert_eq!(catalog.security_policy.actions.len(), 12);
    assert_eq!(
        catalog.security_policy.actions["overwrite_original"],
        "checkpoint_and_confirm"
    );
    assert!(catalog.security_policy.workspace_only);

    let rpc_methods: BTreeSet<_> = catalog
        .rpc_methods
        .iter()
        .map(|method| method.name.as_str())
        .collect();
    assert_eq!(
        rpc_methods,
        BTreeSet::from([
            "office.command.execute",
            "office.host.handshake",
            "office.host.ping",
            "office.host.shutdown",
            "office.job.cancel",
            "office.job.get",
        ])
    );

    let repository = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let command_schema = repository
        .join("manifests")
        .join(&catalog.command_params_schema);
    assert!(command_schema.is_file(), "{}", command_schema.display());
    let command_schema_json: serde_json::Value = serde_json::from_str(
        &fs::read_to_string(&command_schema)
            .unwrap_or_else(|error| panic!("read {}: {error}", command_schema.display())),
    )
    .unwrap_or_else(|error| panic!("parse {}: {error}", command_schema.display()));
    assert_eq!(command_schema_json["type"], "object");
    assert_eq!(command_schema_json["additionalProperties"], false);

    let names: BTreeSet<_> = catalog
        .capabilities
        .iter()
        .map(|capability| capability.name.as_str())
        .collect();
    assert_eq!(
        names,
        BTreeSet::from([
            "batch.convert",
            "batch.replace_text",
            "deck.compile",
            "document.inspect",
            "slide.render",
        ])
    );
    assert_eq!(names.len(), catalog.capabilities.len());

    let mcp_tools: BTreeSet<_> = catalog
        .capabilities
        .iter()
        .map(|capability| capability.mcp_tool.as_str())
        .collect();
    assert_eq!(mcp_tools.len(), catalog.capabilities.len());

    let catalog_errors: BTreeSet<_> = catalog
        .errors
        .iter()
        .map(|error| error.code.clone())
        .collect();
    let rust_errors: BTreeSet<_> = OfficeErrorCode::ALL
        .iter()
        .map(|code| serde_json::to_string(code).unwrap())
        .map(|json| json.trim_matches('"').to_string())
        .collect();
    assert_eq!(rust_errors, catalog_errors);
    assert!(catalog_errors.contains("OFFICE_INVALID_REQUEST"));
    for error in &catalog.errors {
        let code: OfficeErrorCode = serde_json::from_str(&format!("\"{}\"", error.code)).unwrap();
        assert_eq!(
            code.is_retryable(),
            error.retryable,
            "retry policy drifted for {}",
            error.code
        );
    }

    for capability in &catalog.capabilities {
        assert!(
            capability
                .errors
                .iter()
                .all(|code| catalog_errors.contains(code)),
            "{} references an unknown error code",
            capability.name
        );
        for schema_ref in [&capability.input_schema, &capability.output_schema] {
            let schema_path = repository.join("manifests").join(schema_ref);
            let schema: serde_json::Value = serde_json::from_str(
                &fs::read_to_string(&schema_path)
                    .unwrap_or_else(|error| panic!("read {}: {error}", schema_path.display())),
            )
            .unwrap_or_else(|error| panic!("parse {}: {error}", schema_path.display()));
            assert_eq!(schema["type"], "object", "{}", schema_path.display());
            assert_eq!(
                schema["additionalProperties"],
                false,
                "{}",
                schema_path.display()
            );
        }
    }
    for method in &catalog.rpc_methods {
        for schema_ref in [&method.params_schema, &method.result_schema] {
            let schema_path = repository.join("manifests").join(schema_ref);
            let schema: serde_json::Value = serde_json::from_str(
                &fs::read_to_string(&schema_path)
                    .unwrap_or_else(|error| panic!("read {}: {error}", schema_path.display())),
            )
            .unwrap_or_else(|error| panic!("parse {}: {error}", schema_path.display()));
            assert_eq!(schema["type"], "object", "{}", schema_path.display());
            assert_eq!(
                schema["additionalProperties"],
                false,
                "{}",
                schema_path.display()
            );
        }
    }
}
