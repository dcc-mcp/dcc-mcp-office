"""office-generate-production-dashboard / generate_dashboard — Workbook IR → XLSX.

Parameter resolution order (dcc-mcp-core execute_script convention):
1. stdin JSON: {"input": ..., "output_dir": ...}
2. CLI flags: --input --out
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

IR_VERSION = "office-ir/1.0"
HEADER_FILL = PatternFill("solid", fgColor="1E2A3D")
HEADER_FONT = Font(color="FFFFFF", bold=True)
ALT_FILL = PatternFill("solid", fgColor="EAF1F8")


class WorkbookIrError(ValueError):
    """Workbook IR contract violation with a json-path hint."""


def _require(mapping: dict, key: str, path: str) -> Any:
    if key not in mapping:
        raise WorkbookIrError(f"[{path}] missing required key '{key}'")
    return mapping[key]


def load_workbook_ir(path: str | Path) -> dict:
    raw_path = Path(path)
    if not raw_path.is_file():
        raise WorkbookIrError(f"[$] input file not found: {raw_path}")
    raw = json.loads(raw_path.read_text(encoding="utf-8"))
    if raw.get("schema_version") != IR_VERSION:
        raise WorkbookIrError(f"[$] expected schema_version '{IR_VERSION}'")
    if raw.get("kind") != "workbook":
        raise WorkbookIrError("[$.kind] expected 'workbook'")
    document = _require(raw, "document", "$")
    worksheets = document.get("worksheets", [])
    if not worksheets:
        raise WorkbookIrError("[$.document.worksheets] must be a non-empty list")
    for wi, ws in enumerate(worksheets):
        path = f"$.document.worksheets[{wi}]"
        if not ws.get("name"):
            raise WorkbookIrError(f"[{path}] missing worksheet 'name'")
        for ti, table in enumerate(ws.get("tables", [])):
            tpath = f"{path}.tables[{ti}]"
            headers = table.get("headers")
            rows = table.get("rows")
            if not isinstance(headers, list) or not headers:
                raise WorkbookIrError(f"[{tpath}] 'headers' must be a non-empty list")
            if not isinstance(rows, list) or not all(isinstance(r, list) for r in rows):
                raise WorkbookIrError(f"[{tpath}] 'rows' must be a list of lists")
    return raw


class DashboardBuilder:
    """Builds one XLSX from a Workbook IR. SRP: layout, formulas, chart."""

    def __init__(self, ir: dict) -> None:
        self.ir = ir
        self.wb = Workbook()
        self.wb.remove(self.wb.active)
        self.summary: dict[str, Any] = {"sheets": [], "rows": 0}

    def build(self) -> Workbook:
        for ws_spec in self.ir["document"]["worksheets"]:
            self._build_sheet(ws_spec)
        for chart_spec in self.ir["document"].get("charts", []):
            self._build_chart(chart_spec)
        return self.wb

    def _build_sheet(self, spec: dict) -> None:
        ws = self.wb.create_sheet(title=spec["name"])
        sheet_rows = 0
        for table in spec.get("tables", []):
            self._write_table(ws, table)
            sheet_rows += len(table["rows"])
        ws.freeze_panes = "A2"
        self.summary["sheets"].append({"name": spec["name"], "tables": len(spec.get("tables", [])), "rows": sheet_rows})
        self.summary["rows"] += sheet_rows

    def _write_table(self, ws, table: dict) -> None:
        start = ws.max_row + 1 if ws.max_row > 1 else 1
        headers = table["headers"]
        for col, header in enumerate(headers, start=1):
            cell = ws.cell(row=start, column=col, value=header)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center")
        for row_idx, row in enumerate(table["rows"], start=start + 1):
            for col_idx, value in enumerate(row, start=1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                if row_idx % 2 == 0:
                    cell.fill = ALT_FILL
        for col in range(1, len(headers) + 1):
            ws.column_dimensions[get_column_letter(col)].width = 22

    def _build_chart(self, spec: dict) -> None:
        if spec.get("type") != "bar":
            raise WorkbookIrError(f"[$.document.charts] unsupported chart type '{spec.get('type')}'")
        if spec["sheet"] not in self.wb.sheetnames:
            raise WorkbookIrError(f"[$.document.charts] unknown sheet '{spec['sheet']}'")
        ws = self.wb[spec["sheet"]]
        chart = BarChart()
        chart.title = spec.get("title", "Chart")
        data = Reference(ws, min_col=_ref_col(spec["values"]), min_row=_ref_row(spec["values"]), max_row=ws.max_row)
        cats = Reference(ws, min_col=_ref_col(spec["categories"]), min_row=_ref_row(spec["categories"]), max_row=ws.max_row)
        chart.add_data(data, titles_from_data=False)
        chart.set_categories(cats)
        ws.add_chart(chart, spec.get("anchor", "E2"))


def _ref_col(ref: str) -> int:
    """A1-style column letters → index (A=1)."""
    letters = "".join(ch for ch in ref if ch.isalpha())
    value = 0
    for ch in letters.upper():
        value = value * 26 + (ord(ch) - ord("A") + 1)
    return value


def _ref_row(ref: str) -> int:
    """A1-style row digits → int."""
    return int("".join(ch for ch in ref if ch.isdigit()))


def artifact_stem(document_id: str) -> str:
    """Safe filesystem stem for a document id (ids may contain ':')."""
    return re.sub(r"[^A-Za-z0-9._-]+", "-", document_id).strip("-") or "dashboard"


def run(params: dict) -> None:
    ir = load_workbook_ir(params["input"])
    builder = DashboardBuilder(ir)
    workbook = builder.build()
    out_dir = Path(params.get("output_dir", "output"))
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"{artifact_stem(ir.get('document_id', 'dashboard'))}.xlsx"
    workbook.save(str(out_path))
    ok = out_path.is_file() and out_path.stat().st_size > 0
    print(
        json.dumps(
            {
                "success": ok,
                "message": f"dashboard '{ir.get('document_id', 'dashboard')}' generated",
                "context": {"artifact": str(out_path), "summary": builder.summary},
            },
            ensure_ascii=False,
        )
    )


def _force_utf8_stdio() -> None:
    """Deterministic output contract: stdout/stderr are always UTF-8.

    On Windows, a piped subprocess stdout defaults to the ANSI codepage
    (charmap) and fails on CJK text. The gateway reads JSON from stdout, so
    the encoding is part of the script contract.
    """
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


def main() -> None:
    _force_utf8_stdio()
    params: dict = {}
    if not sys.stdin.isatty():
        raw = sys.stdin.read()
        if raw.strip():
            try:
                params = json.loads(raw)
            except json.JSONDecodeError:
                params = {}
    if not params:
        parser = argparse.ArgumentParser(description="Workbook IR → styled XLSX dashboard")
        parser.add_argument("--input", required=True)
        parser.add_argument("--out", dest="output_dir", default="output")
        params = vars(parser.parse_args())
    try:
        run(params)
    except Exception as exc:  # noqa: BLE001
        print(json.dumps({"success": False, "message": str(exc), "context": {}}, ensure_ascii=False))
        sys.exit(1)


if __name__ == "__main__":
    main()
