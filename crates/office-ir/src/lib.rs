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
    //! Word IR (proposal §13.3): styles / sections / paragraphs / lists /
    //! tables / figures+captions / content_controls / headers+footers /
    //! fields+TOC / review_policy.
    use serde::{Deserialize, Serialize};

    #[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
    #[serde(default)]
    pub struct WordDocumentIr {
        /// Style names the compiler must materialise.
        pub styles: Vec<String>,
        pub sections: Vec<Section>,
        /// Main-body paragraphs (before the first structured section).
        pub paragraphs: Vec<Paragraph>,
        pub lists: Vec<ListBlock>,
        pub tables: Vec<TableBlock>,
        pub figures: Vec<Figure>,
        pub content_controls: Vec<ContentControl>,
        pub headers: Vec<HeaderFooterBlock>,
        pub footers: Vec<HeaderFooterBlock>,
        /// Fields and generated tables (TOC) to insert.
        pub fields: Vec<FieldSpec>,
        pub review_policy: ReviewPolicy,
    }

    #[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
    #[serde(default)]
    pub struct Section {
        #[serde(default, skip_serializing_if = "Option::is_none")]
        pub title: Option<String>,
        pub paragraphs: Vec<Paragraph>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        pub page_break_before: Option<bool>,
    }

    #[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
    pub struct Paragraph {
        pub text: String,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        pub style: Option<String>,
    }

    #[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
    #[serde(default)]
    pub struct ListBlock {
        pub items: Vec<String>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        pub style: Option<String>,
    }

    #[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
    #[serde(default)]
    pub struct TableBlock {
        pub header: bool,
        pub rows: Vec<Vec<String>>,
    }

    #[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
    pub struct Figure {
        /// Resource id resolved through the envelope's resources list.
        pub resource: String,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        pub caption: Option<String>,
    }

    /// Anchored Content Control the compiler must fill (proposal §15.5).
    #[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
    pub struct ContentControl {
        pub tag: String,
        pub value: String,
    }

    /// Per-section header/footer block (0-based section index).
    #[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
    pub struct HeaderFooterBlock {
        pub section_index: usize,
        pub text: String,
    }

    #[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
    pub struct FieldSpec {
        /// toc | page | date | custom.
        pub kind: String,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        pub code: Option<String>,
    }

    #[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
    #[serde(default)]
    pub struct ReviewPolicy {
        pub track_changes: bool,
        pub comments_locked: bool,
    }
}

pub mod workbook {
    //! Excel IR (proposal §13.4): worksheets / tables / named_ranges /
    //! formulas / validations / conditional_formats / charts / pivots /
    //! calculation_policy.
    use serde::{Deserialize, Serialize};

    #[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
    #[serde(default)]
    pub struct WorkbookIr {
        pub worksheets: Vec<Worksheet>,
        pub tables: Vec<TableSpec>,
        pub named_ranges: Vec<NamedRange>,
        pub formulas: Vec<FormulaSpec>,
        pub validations: Vec<ValidationSpec>,
        pub conditional_formats: Vec<ConditionalFormatSpec>,
        pub charts: Vec<ChartSpec>,
        pub pivots: Vec<PivotSpec>,
        pub calculation_policy: CalculationPolicy,
    }

    #[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
    pub struct Worksheet {
        pub name: String,
        #[serde(default)]
        pub rows: Vec<Vec<CellValue>>,
    }

    /// Untagged cell value: text, number, boolean, or a formula.
    #[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
    #[serde(untagged)]
    pub enum CellValue {
        Text(String),
        Number(f64),
        Bool(bool),
        Formula { formula: String },
    }

    #[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
    #[serde(default)]
    pub struct TableSpec {
        pub worksheet: String,
        /// A1-style range.
        pub range: String,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        pub name: Option<String>,
    }

    #[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
    pub struct NamedRange {
        pub name: String,
        /// A1-style reference (sheet-qualified).
        pub refers_to: String,
    }

    #[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
    pub struct FormulaSpec {
        pub worksheet: String,
        /// A1-style cell address.
        pub cell: String,
        pub formula: String,
    }

    #[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
    pub struct ValidationSpec {
        pub worksheet: String,
        pub range: String,
        /// list | whole | decimal | date | custom.
        pub kind: String,
        #[serde(default)]
        pub params: serde_json::Value,
    }

    #[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
    pub struct ConditionalFormatSpec {
        pub worksheet: String,
        pub range: String,
        /// cell_value | color_scale | data_bar | formula.
        pub kind: String,
        #[serde(default)]
        pub params: serde_json::Value,
    }

    #[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
    #[serde(default)]
    pub struct ChartSpec {
        pub worksheet: String,
        /// bar | line | pie | scatter | column.
        pub kind: String,
        pub data_range: String,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        pub title: Option<String>,
    }

    impl Default for ChartSpec {
        fn default() -> Self {
            Self {
                worksheet: String::new(),
                kind: "column".to_string(),
                data_range: String::new(),
                title: None,
            }
        }
    }

    #[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
    pub struct PivotSpec {
        pub worksheet: String,
        pub name: String,
        /// Source table or range.
        pub source: String,
    }

    #[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
    #[serde(default)]
    pub struct CalculationPolicy {
        /// auto | manual.
        pub mode: String,
        pub full_calc_on_load: bool,
    }

    impl Default for CalculationPolicy {
        fn default() -> Self {
            Self {
                mode: "auto".to_string(),
                full_calc_on_load: false,
            }
        }
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

    #[test]
    fn word_document_ir_round_trip() {
        let doc = Envelope {
            schema_version: IR_VERSION.into(),
            kind: DocumentKind::WordDocument,
            document_id: "draft:report".into(),
            metadata: Metadata {
                title: "Technical Report".into(),
                author: "DCC-MCP Agent".into(),
                language: "en-US".into(),
            },
            template: Some(TemplateRef {
                uri: "brand://internal/technical-report-v1".into(),
                version: "1.2.0".into(),
            }),
            resources: vec![],
            document: word::WordDocumentIr {
                styles: vec!["Heading 1".into(), "Body".into()],
                sections: vec![word::Section {
                    title: Some("Findings".into()),
                    paragraphs: vec![word::Paragraph {
                        text: "The pipeline holds.".into(),
                        style: Some("Body".into()),
                    }],
                    page_break_before: Some(true),
                }],
                paragraphs: vec![word::Paragraph {
                    text: "Abstract paragraph.".into(),
                    style: None,
                }],
                lists: vec![word::ListBlock {
                    items: vec!["one".into(), "two".into()],
                    style: None,
                }],
                tables: vec![word::TableBlock {
                    header: true,
                    rows: vec![vec!["a".into(), "b".into()]],
                }],
                figures: vec![word::Figure {
                    resource: "img:1".into(),
                    caption: Some("Figure 1".into()),
                }],
                content_controls: vec![word::ContentControl {
                    tag: "customer_name".into(),
                    value: "DCC-MCP".into(),
                }],
                headers: vec![word::HeaderFooterBlock {
                    section_index: 0,
                    text: "Confidential".into(),
                }],
                footers: vec![],
                fields: vec![word::FieldSpec {
                    kind: "toc".into(),
                    code: None,
                }],
                review_policy: word::ReviewPolicy {
                    track_changes: true,
                    comments_locked: false,
                },
            },
            validation: vec![],
            outputs: vec!["docx".into()],
        };
        let json = serde_json::to_string(&doc).unwrap();
        let back: Envelope<word::WordDocumentIr> = serde_json::from_str(&json).unwrap();
        assert_eq!(back.document.sections[0].title.as_deref(), Some("Findings"));
        assert_eq!(back.document.content_controls[0].tag, "customer_name");
        assert!(back.document.review_policy.track_changes);
        assert!(json.contains("\"toc\""));
    }

    #[test]
    fn workbook_ir_round_trip_and_defaults() {
        let book = Envelope {
            schema_version: IR_VERSION.into(),
            kind: DocumentKind::Workbook,
            document_id: "draft:dashboard".into(),
            metadata: Metadata {
                title: "Capability Dashboard".into(),
                author: "DCC-MCP Agent".into(),
                language: "zh-CN".into(),
            },
            template: None,
            resources: vec![],
            document: workbook::WorkbookIr {
                worksheets: vec![workbook::Worksheet {
                    name: "Summary".into(),
                    rows: vec![vec![
                        workbook::CellValue::Text("Total".into()),
                        workbook::CellValue::Number(42.0),
                        workbook::CellValue::Formula {
                            formula: "=SUM(B2:B10)".into(),
                        },
                        workbook::CellValue::Bool(true),
                    ]],
                }],
                tables: vec![workbook::TableSpec {
                    worksheet: "Summary".into(),
                    range: "A1:C10".into(),
                    name: Some("CapTable".into()),
                }],
                named_ranges: vec![workbook::NamedRange {
                    name: "Totals".into(),
                    refers_to: "Summary!$B$2:$B$10".into(),
                }],
                formulas: vec![workbook::FormulaSpec {
                    worksheet: "Summary".into(),
                    cell: "B11".into(),
                    formula: "=SUM(B2:B10)".into(),
                }],
                validations: vec![workbook::ValidationSpec {
                    worksheet: "Summary".into(),
                    range: "C2:C10".into(),
                    kind: "list".into(),
                    params: serde_json::json!({ "source": "A,B,C" }),
                }],
                conditional_formats: vec![workbook::ConditionalFormatSpec {
                    worksheet: "Summary".into(),
                    range: "B2:B10".into(),
                    kind: "color_scale".into(),
                    params: serde_json::json!({}),
                }],
                charts: vec![workbook::ChartSpec {
                    worksheet: "Summary".into(),
                    kind: "bar".into(),
                    data_range: "A1:B10".into(),
                    title: Some("Capabilities".into()),
                }],
                pivots: vec![],
                calculation_policy: workbook::CalculationPolicy::default(),
            },
            validation: vec![],
            outputs: vec!["xlsx".into()],
        };
        let json = serde_json::to_string(&book).unwrap();
        let back: Envelope<workbook::WorkbookIr> = serde_json::from_str(&json).unwrap();
        assert_eq!(back.document.worksheets[0].name, "Summary");
        assert_eq!(back.document.charts[0].kind, "bar");
        assert_eq!(back.document.calculation_policy.mode, "auto");
        assert!(!back.document.calculation_policy.full_calc_on_load);

        // Partial IRs deserialize with defaults (progressive authoring).
        let partial: workbook::WorkbookIr =
            serde_json::from_str(r#"{"worksheets":[{"name":"S","rows":[]}]}"#).unwrap();
        assert_eq!(partial.calculation_policy.mode, "auto");
        assert!(partial.charts.is_empty());
    }
}
