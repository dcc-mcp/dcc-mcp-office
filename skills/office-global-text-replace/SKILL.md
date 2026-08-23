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
- policy.workspace_root — absolute workspace boundary containing every input
- confirmation — commit-only proof:
  `{ "action": "overwrite_original", "confirmed": true,
  "confirmed_by": "human:<id>", "confirmed_at": "<RFC3339>" }`

## Planning steps

1. Resolve inputs; route each file to its app sidecar (per-app COM backend).
2. Submit the dry-run and poll its returned `job_id` through `office.job.get`;
   read matched files, per-file/per-rule counts, and unsafe scopes from the
   terminal `result`.
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
(checkpoint + confirm policy). The host validates the structured confirmation
proof before attaching COM and refuses any policy that relaxes catalog defaults
(macros, VBA, external links, ExecuteMso, workspace confinement).

## Validation rules

- dry-run: total_replaced == 0 and per-file matches reported;
- commit: replacements match the dry-run counts (or are re-counted per file);
- per-file result reports backend used and warnings for unsupported scopes.

## Failure compensation

Per-file failures do not abort the batch (partial success). Busy / modal /
timeout follow the retry ladder; dry-run opens documents read-only so a
failed scan never mutates anything. A timed-out commit reports
`indeterminate: true`; re-inspect the source before deciding whether to retry.
`office.job.cancel` is cooperative: it is accepted immediately, becomes
`cancelled` at the next file boundary, and never interrupts an in-flight COM
write.

## Artifact naming

No new artifacts for dry-run. Before each commit changes a file in place, the
Host creates a byte-exact `*.dcc-checkpoint-<operation_id>.*` pre-image beside
the source and returns it as a `checkpoint` artifact with SHA-256.

## Agent-visible summary

Report: files scanned, matches/replacements per rule, backend per file,
scopes skipped with reasons, and which files need re-inspection after commit.
