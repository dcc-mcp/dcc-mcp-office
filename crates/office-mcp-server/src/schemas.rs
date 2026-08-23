use serde_json::{Map, Value};

const COMMAND_PARAMS: &str = include_str!("../../../manifests/schemas/command-params.schema.json");
const COMMAND_RESULT: &str = include_str!("../../../manifests/schemas/command-result.schema.json");
const DECK_COMPILE: &str = include_str!("../../../manifests/schemas/deck-compile.schema.json");
const DOCUMENT_INSPECT: &str =
    include_str!("../../../manifests/schemas/document-inspect.schema.json");
const BATCH_CONVERT: &str = include_str!("../../../manifests/schemas/batch-convert.schema.json");
const BATCH_REPLACE_TEXT: &str =
    include_str!("../../../manifests/schemas/batch-replace-text.schema.json");
const SLIDE_RENDER: &str = include_str!("../../../manifests/schemas/slide-render.schema.json");
const JOB_ID: &str = include_str!("../../../manifests/schemas/job-id.schema.json");
const JOB_STATUS: &str = include_str!("../../../manifests/schemas/job-status.schema.json");
const JOB_CANCEL_RESULT: &str =
    include_str!("../../../manifests/schemas/job-cancel-result.schema.json");

pub(crate) fn load(path: &str) -> Result<Value, String> {
    let source = match path {
        "schemas/command-params.schema.json" => COMMAND_PARAMS,
        "schemas/command-result.schema.json" => COMMAND_RESULT,
        "schemas/deck-compile.schema.json" => DECK_COMPILE,
        "schemas/document-inspect.schema.json" => DOCUMENT_INSPECT,
        "schemas/batch-convert.schema.json" => BATCH_CONVERT,
        "schemas/batch-replace-text.schema.json" => BATCH_REPLACE_TEXT,
        "schemas/slide-render.schema.json" => SLIDE_RENDER,
        "schemas/job-id.schema.json" => JOB_ID,
        "schemas/job-status.schema.json" => JOB_STATUS,
        "schemas/job-cancel-result.schema.json" => JOB_CANCEL_RESULT,
        _ => return Err(format!("catalog references an unbundled schema: {path}")),
    };
    serde_json::from_str(source).map_err(|error| format!("invalid bundled schema {path}: {error}"))
}

/// Task-level tools flatten the capability input and the safe outer command
/// envelope into one MCP arguments object. The wire capability is implied by
/// the MCP tool name and is never accepted from callers.
pub(crate) fn command_tool_input(capability_schema: &Value) -> Result<Value, String> {
    let mut schema = capability_schema
        .as_object()
        .cloned()
        .ok_or_else(|| "capability input schema must be an object".to_string())?;
    schema.remove("$id");
    let properties = schema
        .get_mut("properties")
        .and_then(Value::as_object_mut)
        .ok_or_else(|| "capability input schema requires object properties".to_string())?;
    let envelope = load("schemas/command-params.schema.json")?;
    let envelope_properties = envelope["properties"]
        .as_object()
        .ok_or_else(|| "command params schema requires object properties".to_string())?;
    for name in ["document", "confirmation", "policy"] {
        properties.insert(
            name.to_string(),
            envelope_properties
                .get(name)
                .cloned()
                .ok_or_else(|| format!("command params schema is missing {name}"))?,
        );
    }
    Ok(Value::Object(schema))
}

pub(crate) fn object(value: &Value) -> Result<Map<String, Value>, String> {
    value
        .as_object()
        .cloned()
        .ok_or_else(|| "MCP tool schema must be a JSON object".to_string())
}
