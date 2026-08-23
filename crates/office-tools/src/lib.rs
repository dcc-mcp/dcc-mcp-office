//! dcc-mcp-office-tools — task-level MCP tool registry (proposal §11).
//!
//! Agents see task-level capabilities, never raw COM members (proposal §3.2 /
//! ADR-004). M0 records the registry; JSON schemas + gateway registration
//! land in M1 via the dcc-mcp-gateway capability index.

#![forbid(unsafe_code)]

pub mod common {
    //! Office-wide tools (proposal §11.1).
    pub const CAPABILITIES_SEARCH: &str = "office.capabilities.search";
    pub const APPLICATION_LIST: &str = "office.application.list";
    pub const SESSION_ATTACH: &str = "office.session.attach";
    pub const DOCUMENT_INSPECT: &str = "office.document.inspect";
    pub const DOCUMENT_GENERATE: &str = "office.document.generate";
    pub const DOCUMENT_PATCH: &str = "office.document.patch";
    pub const DOCUMENT_RENDER: &str = "office.document.render";
    pub const DOCUMENT_VALIDATE: &str = "office.document.validate";
    pub const DOCUMENT_EXPORT: &str = "office.document.export";
    pub const BATCH_CONVERT: &str = "office.batch.convert";
    pub const BATCH_REPLACE_TEXT: &str = "office.batch.replace_text";
    pub const BATCH_APPLY_TEMPLATE: &str = "office.batch.apply_template";
    pub const JOB_GET: &str = "office.job.get";
    pub const JOB_CANCEL: &str = "office.job.cancel";
}

pub mod powerpoint {
    //! PowerPoint tools (proposal §11.2) — implemented in dcc-mcp-PowerPoint.
    pub const DECK_GENERATE: &str = "powerpoint.deck.generate";
    pub const SLIDE_COMPOSE: &str = "powerpoint.slide.compose";
    pub const SLIDE_RENDER: &str = "powerpoint.slide.render";
    pub const ANIMATION_APPLY: &str = "powerpoint.animation.apply";
    pub const SLIDESHOW_CONTROL: &str = "powerpoint.slideshow.control";
}

pub mod word {
    //! Word tools (proposal §11.2) — Phase 2.
    pub const DOCUMENT_REFLOW: &str = "word.document.reflow";
    pub const FIELDS_UPDATE: &str = "word.fields.update";
    pub const TOC_REBUILD: &str = "word.toc.rebuild";
    pub const TRACK_CHANGES_INSPECT: &str = "word.track_changes.inspect";
}

pub mod excel {
    //! Excel tools (proposal §11.2) — Phase 2.
    pub const WORKBOOK_CALCULATE: &str = "excel.workbook.calculate";
    pub const TABLE_UPDATE: &str = "excel.table.update";
    pub const CHART_GENERATE: &str = "excel.chart.generate";
    pub const PIVOT_REFRESH: &str = "excel.pivot.refresh";
}

pub mod outlook {
    //! Outlook tools (proposal §11.2) — Phase 3.
    pub const MESSAGE_CREATE_DRAFT: &str = "outlook.message.create_draft";
    pub const CALENDAR_PREPARE_EVENT: &str = "outlook.calendar.prepare_event";
}

pub mod visio {
    //! Visio tools (proposal §11.2) — Phase 4.
    pub const DIAGRAM_LAYOUT: &str = "visio.diagram.layout";
    pub const DIAGRAM_CONNECT: &str = "visio.diagram.connect";
}

pub mod project {
    //! Project tools (proposal §11.2) — Phase 4.
    pub const PLAN_GENERATE: &str = "project.plan.generate";
    pub const RESOURCES_ASSIGN: &str = "project.resources.assign";
}

pub mod access {
    //! Access tools (proposal §11.2) — Phase 4.
    pub const QUERY_EXECUTE: &str = "access.query.execute";
    pub const REPORT_EXPORT: &str = "access.report.export";
}

/// Tool name → owning app + phase (proposal §7 support matrix).
pub struct RegistryEntry {
    pub name: &'static str,
    pub app: &'static str,
    /// P0 (PowerPoint/Word/Excel), P1 (Outlook/OneNote/Visio), P2 (Project/Access).
    pub phase: &'static str,
}

/// Mapping from an agent-facing task-level MCP tool to the sidecar wire
/// capability it invokes. The machine-readable catalog owns these mappings.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WireCapabilityMapping {
    pub wire_capability: String,
    pub mcp_tool: String,
}

pub fn implemented_wire_mappings() -> Vec<WireCapabilityMapping> {
    dcc_mcp_office_protocol::capability_catalog()
        .capabilities
        .iter()
        .map(|capability| WireCapabilityMapping {
            wire_capability: capability.name.clone(),
            mcp_tool: capability.mcp_tool.clone(),
        })
        .collect()
}

/// Registry of every task-level Office capability (proposal §11.1/§11.2).
pub const REGISTRY: &[RegistryEntry] = &[
    RegistryEntry {
        name: common::CAPABILITIES_SEARCH,
        app: "office",
        phase: "P0",
    },
    RegistryEntry {
        name: common::APPLICATION_LIST,
        app: "office",
        phase: "P0",
    },
    RegistryEntry {
        name: common::SESSION_ATTACH,
        app: "office",
        phase: "P0",
    },
    RegistryEntry {
        name: common::DOCUMENT_INSPECT,
        app: "office",
        phase: "P0",
    },
    RegistryEntry {
        name: common::DOCUMENT_GENERATE,
        app: "office",
        phase: "P0",
    },
    RegistryEntry {
        name: common::DOCUMENT_PATCH,
        app: "office",
        phase: "P0",
    },
    RegistryEntry {
        name: common::DOCUMENT_RENDER,
        app: "office",
        phase: "P0",
    },
    RegistryEntry {
        name: common::DOCUMENT_VALIDATE,
        app: "office",
        phase: "P0",
    },
    RegistryEntry {
        name: common::DOCUMENT_EXPORT,
        app: "office",
        phase: "P0",
    },
    RegistryEntry {
        name: common::BATCH_CONVERT,
        app: "office",
        phase: "P0",
    },
    RegistryEntry {
        name: common::BATCH_REPLACE_TEXT,
        app: "office",
        phase: "P0",
    },
    RegistryEntry {
        name: common::BATCH_APPLY_TEMPLATE,
        app: "office",
        phase: "P0",
    },
    RegistryEntry {
        name: common::JOB_GET,
        app: "office",
        phase: "P0",
    },
    RegistryEntry {
        name: common::JOB_CANCEL,
        app: "office",
        phase: "P0",
    },
    RegistryEntry {
        name: powerpoint::DECK_GENERATE,
        app: "powerpoint",
        phase: "P0",
    },
    RegistryEntry {
        name: powerpoint::SLIDE_COMPOSE,
        app: "powerpoint",
        phase: "P0",
    },
    RegistryEntry {
        name: powerpoint::SLIDE_RENDER,
        app: "powerpoint",
        phase: "P0",
    },
    RegistryEntry {
        name: powerpoint::ANIMATION_APPLY,
        app: "powerpoint",
        phase: "P0",
    },
    RegistryEntry {
        name: powerpoint::SLIDESHOW_CONTROL,
        app: "powerpoint",
        phase: "P1",
    },
    RegistryEntry {
        name: word::DOCUMENT_REFLOW,
        app: "word",
        phase: "P0",
    },
    RegistryEntry {
        name: word::FIELDS_UPDATE,
        app: "word",
        phase: "P0",
    },
    RegistryEntry {
        name: word::TOC_REBUILD,
        app: "word",
        phase: "P0",
    },
    RegistryEntry {
        name: word::TRACK_CHANGES_INSPECT,
        app: "word",
        phase: "P0",
    },
    RegistryEntry {
        name: excel::WORKBOOK_CALCULATE,
        app: "excel",
        phase: "P0",
    },
    RegistryEntry {
        name: excel::TABLE_UPDATE,
        app: "excel",
        phase: "P0",
    },
    RegistryEntry {
        name: excel::CHART_GENERATE,
        app: "excel",
        phase: "P0",
    },
    RegistryEntry {
        name: excel::PIVOT_REFRESH,
        app: "excel",
        phase: "P0",
    },
    RegistryEntry {
        name: outlook::MESSAGE_CREATE_DRAFT,
        app: "outlook",
        phase: "P1",
    },
    RegistryEntry {
        name: outlook::CALENDAR_PREPARE_EVENT,
        app: "outlook",
        phase: "P1",
    },
    RegistryEntry {
        name: visio::DIAGRAM_LAYOUT,
        app: "visio",
        phase: "P1",
    },
    RegistryEntry {
        name: visio::DIAGRAM_CONNECT,
        app: "visio",
        phase: "P1",
    },
    RegistryEntry {
        name: project::PLAN_GENERATE,
        app: "project",
        phase: "P2",
    },
    RegistryEntry {
        name: project::RESOURCES_ASSIGN,
        app: "project",
        phase: "P2",
    },
    RegistryEntry {
        name: access::QUERY_EXECUTE,
        app: "access",
        phase: "P2",
    },
    RegistryEntry {
        name: access::REPORT_EXPORT,
        app: "access",
        phase: "P2",
    },
];

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn registry_names_are_unique() {
        let names: Vec<_> = REGISTRY.iter().map(|e| e.name).collect();
        let mut sorted = names.clone();
        sorted.sort_unstable();
        sorted.dedup();
        assert_eq!(names.len(), sorted.len());
    }

    #[test]
    fn p0_tools_cover_mvp_surface() {
        let p0: Vec<_> = REGISTRY
            .iter()
            .filter(|e| e.phase == "P0")
            .map(|e| e.name)
            .collect();
        for required in [
            common::BATCH_CONVERT,
            common::BATCH_REPLACE_TEXT,
            common::DOCUMENT_INSPECT,
            powerpoint::DECK_GENERATE,
        ] {
            assert!(p0.contains(&required), "missing {required}");
        }
    }

    #[test]
    fn catalog_mappings_target_registered_mcp_tools() {
        let registered: std::collections::HashSet<_> =
            REGISTRY.iter().map(|entry| entry.name).collect();
        for mapping in implemented_wire_mappings() {
            assert!(
                registered.contains(mapping.mcp_tool.as_str()),
                "catalog maps {} to unregistered MCP tool {}",
                mapping.wire_capability,
                mapping.mcp_tool
            );
        }
    }
}
