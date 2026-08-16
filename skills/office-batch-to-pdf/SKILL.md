---
name: office-batch-to-pdf
description: "Batch-convert PPTX/DOCX/XLSX to high-fidelity PDF. Local files render through the Office COM sidecar (native export); OneDrive/SharePoint files convert through Microsoft Graph. Open XML performs preflight only — it never claims to render PDF."
dcc: office
version: "0.1.0"
license: "MIT"
compatibility: "Windows, Office 2019+ / Microsoft 365; Graph scenarios need tenant auth"
tags: ["office", "convert", "pdf", "batch", "job"]
capabilities:
  - office.batch.convert
  - office.job.get
  - office.job.cancel
  - office.document.validate
---

# office-batch-to-pdf

Batch-convert office documents to PDF with per-file validation and a full
artifact report (proposal §15.1).

## Input contract

- `inputs.glob` — file pattern (pptx/docx/xlsx)
- `target_format` — "pdf"
- `backend` — "auto" | "desktop_com" | "graph"
- `output.directory` + `output.mode` ("mirror_tree") + `output.overwrite` ("versioned")
- `validation` — e.g. ["output_openable", "non_empty", "page_count_reasonable"]

## Planning steps

1. Resolve inputs → inspect formats (Open XML preflight).
2. Select backend per file: local + high fidelity → COM native export;
   OneDrive/SharePoint → Graph conversion; unsupported → explicit error,
   never a silent low-fidelity substitute.
3. Submit as one Job (`office.batch.convert`); poll `office.job.get`.
4. Validate PDFs; publish artifacts.

## Provider choice

Local files: `desktop_com`. Cloud files: `graph`. Open XML is preflight
only. Unsupported formats return `OFFICE_CAPABILITY_UNSUPPORTED` with a reason.

## Safety confirmation

- Read-only conversion → no confirmation required.
- Overwriting existing outputs in `versioned` mode → automatic
  (never overwrites originals); `overwrite.original` → checkpoint +
  confirm policy.

## Validation rules

- output openable, non-empty, page-count reasonable;
- per-file result in the job report (ok / error / warnings / backend used).

## Failure compensation

- Per-file failures do not abort the job → `partially_succeeded`.
- Busy/timeout → auto-retry (policy-defined); modal dialog → report and wait.
- Determinable failures surface `OFFICE_PARTIAL_SUCCESS` with the item list.

## Artifact naming

- Mirror tree: `<output>/<relative-path>.pdf`
- Versioned overwrite: `.v2.pdf`, `.v3.pdf` (never clobber).

## Agent-visible summary

The job result lists: files converted, files failed (with reason), warnings,
artifact paths + sha256, backend per file, and whether human review is needed.
