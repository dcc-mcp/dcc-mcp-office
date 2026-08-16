//! dcc-mcp-office-ir — application-specific document IRs (proposal §13).
//!
//! One common envelope, one IR per application. Application semantics are
//! intentionally **not** unified across PowerPoint / Word / Excel (proposal
//! §3.3): the document payload of the envelope always uses the app-specific
//! schema.

#![forbid(unsafe_code)]

use serde::{Deserialize, Serialize};

/// IR schema version written into every envelope (proposal §13.1).
pub const IR_VERSION: &str = "office-ir/1.0";

/// Document kinds carried by the envelope.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum DocumentKind {
    Presentation,
    WordDocument,
    Workbook,
}

/// Template reference: brand:// registry URI plus pinned version
/// (proposal §15.4).
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct TemplateRef {
    pub uri: String,
    pub version: String,
}

/// External resource (image, media, data source) referenced by the document.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct Resource {
    pub id: String,
    pub uri: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub mime: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct Metadata {
    pub title: String,
    #[serde(default)]
    pub author: String,
    #[serde(default)]
    pub language: String,
}

/// Common envelope (proposal §13.1). The generic T is the app-specific
/// document schema.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Envelope<T> {
    pub schema_version: String,
    pub kind: DocumentKind,
    pub document_id: String,
    pub metadata: Metadata,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub template: Option<TemplateRef>,
    #[serde(default)]
    pub resources: Vec<Resource>,
    pub document: T,
    #[serde(default)]
    pub validation: Vec<String>,
    #[serde(default)]
    pub outputs: Vec<String>,
}

pub mod presentation {
    //! PowerPoint IR (proposal §13.2).
    //!
    //! Slides reference **semantic layouts** (title_cover, kpi_dashboard,
    //! technical_architecture, ...) instead of raw coordinates; the Open XML
    //! compiler + COM finalizer materialise them.

    use super::*;

    #[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
    pub struct PresentationIr {
        #[serde(default, skip_serializing_if = "Option::is_none")]
        pub theme: Option<String>,
        #[serde(default)]
        pub master: Option<TemplateRef>,
        #[serde(default)]
        pub layouts: Vec<TemplateRef>,
        pub slides: Vec<Slide>,
        #[serde(default)]
        pub export_policy: ExportPolicy,
    }

    #[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
    pub struct Slide {
        /// Native SlideID once materialised by Office; none while drafting.
        #[serde(default, skip_serializing_if = "Option::is_none")]
        pub id: Option<u32>,
        pub semantic_layout: String,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        pub title: Option<String>,
        #[serde(default)]
        pub content_blocks: Vec<ContentBlock>,
        #[serde(default)]
        pub images: Vec<Resource>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        pub speaker_notes: Option<String>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        pub animation_timeline: Option<serde_json::Value>,
    }

    #[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
    #[serde(tag = "type", rename_all = "snake_case")]
    pub enum ContentBlock {
        Text {
            paragraphs: Vec<String>,
            #[serde(default, skip_serializing_if = "Option::is_none")]
            style: Option<String>,
        },
        Bullets {
            items: Vec<String>,
        },
        Table {
            header: bool,
            rows: Vec<Vec<String>>,
        },
        Chart {
            chart_type: String,
            data: serde_json::Value,
        },
        Image {
            resource: String,
            #[serde(default, skip_serializing_if = "Option::is_none")]
            fit: Option<String>,
        },
    }

    #[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
    pub struct ExportPolicy {
        #[serde(default)]
        pub include_speaker_notes: bool,
        #[serde(default = "default_true")]
        pub pdf: bool,
        #[serde(default = "default_true")]
        pub slide_previews: bool,
    }

    impl Default for ExportPolicy {
        fn default() -> Self {
            Self {
                include_speaker_notes: false,
                pdf: true,
                slide_previews: true,
            }
        }
    }
}

pub mod word {
    //! Word IR (proposal §13.3) — Phase 2. Shape planned: styles / sections /
    //! paragraphs / lists / tables / figures+captions / content_controls /
    //! headers+footers / fields+TOC / review_policy.
    use serde::{Deserialize, Serialize};

    /// Placeholder outlining the planned structure; filled in Phase 2
    /// alongside the dcc-mcp-word adapter.
    #[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
    pub struct WordDocumentIr {
        pub styles: Vec<String>,
        pub content_controls: Vec<String>,
    }
}

pub mod workbook {
    //! Excel IR (proposal §13.4) — Phase 2. Shape planned: worksheets /
    //! tables / named_ranges / formulas / validations / conditional_formats /
    //! charts / pivots / calculation_policy.
    use serde::{Deserialize, Serialize};

    /// Placeholder outlining the planned structure; filled in Phase 2
    /// alongside the dcc-mcp-excel adapter.
    #[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
    pub struct WorkbookIr {
        pub worksheets: Vec<String>,
        pub named_ranges: Vec<String>,
    }
}

fn default_true() -> bool {
    true
}

#[cfg(test)]
mod tests {
    use super::presentation::*;
    use super::*;

    #[test]
    fn presentation_envelope_round_trip() {
        let deck = Envelope {
            schema_version: IR_VERSION.into(),
            kind: DocumentKind::Presentation,
            document_id: "draft:review-deck".into(),
            metadata: Metadata {
                title: "DCC-MCP Production Review".into(),
                author: "DCC-MCP Agent".into(),
                language: "zh-CN".into(),
            },
            template: Some(TemplateRef {
                uri: "brand://studio/review-v3".into(),
                version: "3.0.0".into(),
            }),
            resources: vec![],
            document: PresentationIr {
                theme: None,
                master: None,
                layouts: vec![],
                slides: vec![Slide {
                    id: None,
                    semantic_layout: "technical_architecture".into(),
                    title: Some("Architecture".into()),
                    content_blocks: vec![ContentBlock::Bullets {
                        items: vec!["Rust control plane".into(), "C# COM data plane".into()],
                    }],
                    images: vec![],
                    speaker_notes: None,
                    animation_timeline: None,
                }],
                export_policy: ExportPolicy::default(),
            },
            validation: vec!["no_text_overflow".into(), "no_out_of_bounds".into()],
            outputs: vec!["pptx".into(), "pdf".into(), "slide-previews".into()],
        };
        let json = serde_json::to_string(&deck).unwrap();
        let back: Envelope<PresentationIr> = serde_json::from_str(&json).unwrap();
        assert_eq!(
            back.document.slides[0].semantic_layout,
            "technical_architecture"
        );
        assert!(json.contains("\"technical_architecture\""));
        assert!(back.document.export_policy.slide_previews);
    }
}
