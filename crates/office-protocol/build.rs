use std::collections::BTreeSet;
use std::env;
use std::fs;
use std::path::PathBuf;

fn rust_variant(wire_name: &str) -> String {
    wire_name
        .split('_')
        .map(|segment| {
            let mut characters = segment.chars();
            match characters.next() {
                Some(first) => {
                    first.to_ascii_uppercase().to_string()
                        + &characters.as_str().to_ascii_lowercase()
                }
                None => String::new(),
            }
        })
        .collect()
}

fn main() {
    let catalog_path =
        PathBuf::from(env::var_os("CARGO_MANIFEST_DIR").expect("CARGO_MANIFEST_DIR"))
            .join("office-rpc.catalog.json");
    println!("cargo:rerun-if-changed={}", catalog_path.display());
    let catalog: serde_json::Value = serde_json::from_str(
        &fs::read_to_string(&catalog_path).expect("read office-rpc capability catalog"),
    )
    .expect("parse office-rpc capability catalog");
    let errors = catalog["errors"]
        .as_array()
        .expect("catalog errors must be an array");
    let mut seen = BTreeSet::new();
    let mut variants = Vec::new();
    let mut retryable = Vec::new();
    for error in errors {
        let code = error["code"]
            .as_str()
            .expect("catalog error code must be a string");
        assert!(seen.insert(code), "duplicate catalog error code: {code}");
        let variant = rust_variant(code);
        if error["retryable"].as_bool() == Some(true) {
            retryable.push(variant.clone());
        }
        variants.push(variant);
    }

    let declarations = variants
        .iter()
        .map(|variant| format!("    {variant},"))
        .collect::<Vec<_>>()
        .join("\n");
    let all = variants
        .iter()
        .map(|variant| format!("        Self::{variant},"))
        .collect::<Vec<_>>()
        .join("\n");
    let retryable_match = retryable
        .iter()
        .map(|variant| format!("Self::{variant}"))
        .collect::<Vec<_>>()
        .join(" | ");
    let generated = format!(
        r#"/// Standard error codes generated from office-rpc.catalog.json.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "SCREAMING_SNAKE_CASE")]
pub enum OfficeErrorCode {{
{declarations}
}}

impl OfficeErrorCode {{
    /// Complete canonical error set in catalog order.
    pub const ALL: &'static [Self] = &[
{all}
    ];

    /// Whether repeating the same call is safe under the recovery policy.
    pub fn is_retryable(self) -> bool {{
        matches!(self, {retryable_match})
    }}
}}
"#
    );
    let output = PathBuf::from(env::var_os("OUT_DIR").expect("OUT_DIR"));
    fs::write(output.join("office_error_codes.rs"), generated)
        .expect("write generated Office error codes");
}
