---
name: office-global-text-replace
description: "Batch replace text across PPTX/DOCX/XLSX with dry-run first. Literal or case-insensitive rules; scope covers body, headers, footers, notes; charts and comments are reported, never silently modified. Default is dry-run — committing is a separate explicit call."
license: "MIT"
compatibility: "Windows, Office 2019+ / Microsoft 365"
metadata:
  dcc-mcp:
    dcc: office
    version: "0.1.0"
    tags: ["office", "replace", "text", "batch", "job"]
    capabilities:
      - office.batch.replace_text
      - office.job.get
      - office.job.cancel
      - office.document.validate
---

# office-global-text-replace

Batch replace text across Office documents — always dry-run first, commit
only after the caller confirms the diff report (proposal §15.2).

## Input contract

- inputs — file paths or globs (pptx / docx / xlsx; each file routes to
  its own application sidecar, one sidecar per app)
- rules — [{ "find": ..., "replace": ..., "match": "literal" | "case_insensitive" }]
- scope — body, headers, footers, notes, comments, charts
- dry_run — boolean, defaults to true; a commit is a separate call with
  dry_run: false after human confirmation

## Planning steps

1. Resolve inputs; route each file to its app sidecar (per-app COM backend).
2. Dry-run first: matched files, per-file/per-rule match counts, scopes that
   cannot be safely modified.
3. Show the diff report; wait for explicit confirmation.
4. Commit (dry_run: false), then re-inspect to verify the result.

## Provider choice

Desktop COM for text replacement: Word Find across runs and multi-section
headers/footers, PowerPoint text frames/tables/notes, Excel cell values.
Open XML is never used for blind archive-level string replacement (it would
corrupt shared strings and splits). Unsupported scopes return warnings, never
silent skips.

## Safety confirmation

Dry-run requires no confirmation. Commit requires human confirmation
(checkpoint + confirm policy). The host refuses any policy that relaxes the
deny-by-default items (macros, VBA, external links, ExecuteMso).

## Validation rules

- dry-run: total_replaced == 0 and per-file matches reported;
- commit: replacements match the dry-run counts (or are re-counted per file);
- per-file result reports backend used and warnings for unsupported scopes.

## Failure compensation

Per-file failures do not abort the batch (partial success). Busy / modal /
timeout follow the retry ladder; dry-run opens documents read-only so a
failed scan never mutates anything.

## Artifact naming

No new artifacts for dry-run. Commit changes files in place; the checkpoint
policy may pin a pre-image copy before writing.

## Agent-visible summary

Report: files scanned, matches/replacements per rule, backend per file,
scopes skipped with reasons, and which files need re-inspection after commit.
