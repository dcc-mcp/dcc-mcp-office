use dcc_mcp_office_protocol::{CapabilityManifest, TemplatePackageCapability};

#[test]
fn capability_manifest_types_materializable_template_packages() {
    let manifest: CapabilityManifest = serde_json::from_str(
        r#"{
          "template_packages": {
            "brand://example/studio-light": {
              "version": "1.0.0",
              "kind": "presentation",
              "source_kind": "file",
              "layouts": ["title_cover", "bullets"]
            }
          }
        }"#,
    )
    .expect("template package manifest");

    assert_eq!(
        manifest.template_packages["brand://example/studio-light"],
        TemplatePackageCapability {
            version: "1.0.0".into(),
            kind: "presentation".into(),
            source_kind: "file".into(),
            layouts: vec!["title_cover".into(), "bullets".into()],
        }
    );
}
