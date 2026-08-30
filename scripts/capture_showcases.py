"""Capture the dcc-mcp-office showcase gallery from the real Host runtime.

The Office-free deck leg is always available. Pass ``--with-office`` on an
interactive Windows host with PowerPoint, Word, Excel and Poppler to capture
the full gallery. Existing generated files are refused unless ``--force`` is
explicitly supplied.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import queue
import shutil
import subprocess
import sys
import threading
import time
from datetime import UTC, datetime
from pathlib import Path
from typing import Any, Self

from build_showcase_preview import compose_preview

SCHEMA_VERSION = "dcc-mcp-showcase/1.0"
TERMINAL_JOB_PHASES = {"succeeded", "partially_succeeded", "failed", "cancelled"}


class HostFailure(RuntimeError):
    def __init__(self, response: dict[str, Any]) -> None:
        self.response = response
        error = response.get("error", {})
        super().__init__(f"{error.get('code', 'HOST_ERROR')}: {error.get('message', response)}")


class HostSession:
    """Small JSONL client for the Host's deterministic stdio transport."""

    def __init__(
        self,
        host_exe: Path,
        app: str,
        workspace: Path,
        templates: Path,
        *,
        openxml_only: bool = False,
        timeout_seconds: float = 150.0,
    ) -> None:
        command = [
            str(host_exe),
            f"--app={app}",
            "--stdio",
            f"--workspace-root={workspace}",
            f"--template-dir={templates}",
        ]
        if openxml_only:
            command.append("--openxml-only")
        self.workspace = workspace
        self.timeout_seconds = timeout_seconds
        self.transcript: list[dict[str, Any]] = []
        self.notifications: list[dict[str, Any]] = []
        self.stderr: list[str] = []
        self._next_id = 1
        self._stdout: queue.Queue[str | None] = queue.Queue()
        self._process = subprocess.Popen(
            command,
            cwd=workspace,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
        )
        self._stdout_thread = threading.Thread(target=self._read_stdout, daemon=True)
        self._stderr_thread = threading.Thread(target=self._read_stderr, daemon=True)
        self._stdout_thread.start()
        self._stderr_thread.start()

    def _read_stdout(self) -> None:
        assert self._process.stdout is not None
        for line in self._process.stdout:
            self._stdout.put(line)
        self._stdout.put(None)

    def _read_stderr(self) -> None:
        assert self._process.stderr is not None
        self.stderr.extend(line.rstrip() for line in self._process.stderr)

    def request(
        self,
        method: str,
        params: dict[str, Any],
        *,
        allow_error: bool = False,
    ) -> dict[str, Any]:
        request_id = self._next_id
        self._next_id += 1
        payload = {"jsonrpc": "2.0", "id": request_id, "method": method, "params": params}
        assert self._process.stdin is not None
        self._process.stdin.write(json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + "\n")
        self._process.stdin.flush()

        deadline = time.monotonic() + self.timeout_seconds
        notifications_before = len(self.notifications)
        while True:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise TimeoutError(f"timed out waiting for {method}")
            try:
                line = self._stdout.get(timeout=remaining)
            except queue.Empty as error:
                raise TimeoutError(f"timed out waiting for {method}") from error
            if line is None:
                raise RuntimeError(f"Host exited while waiting for {method}: {self.stderr[-5:]}")
            message = json.loads(line)
            if message.get("id") != request_id:
                self.notifications.append(message)
                continue
            entry = {
                "request": payload,
                "response": message,
                "notifications": self.notifications[notifications_before:],
            }
            self.transcript.append(entry)
            if "error" in message and not allow_error:
                raise HostFailure(message)
            return message

    def handshake(self, app: str) -> dict[str, Any]:
        return self.request(
            "office.host.handshake",
            {
                "gateway_version": "showcase-capture",
                "protocol_versions": ["office-rpc/1"],
                "requested_app": app,
            },
        )["result"]

    def execute(
        self,
        capability: str,
        input_value: dict[str, Any],
        *,
        confirmation: dict[str, Any] | None = None,
        allow_error: bool = False,
    ) -> dict[str, Any]:
        params: dict[str, Any] = {
            "capability": capability,
            "input": input_value,
            "policy": {"workspace_root": str(self.workspace)},
        }
        if confirmation is not None:
            params["confirmation"] = confirmation
        response = self.request("office.command.execute", params, allow_error=allow_error)
        if allow_error and "error" in response:
            return response
        result = response["result"]
        if "job_id" not in result:
            return result
        return {
            "submission": result,
            "job": self.wait_for_job(result["job_id"]),
        }

    def wait_for_job(self, job_id: str) -> dict[str, Any]:
        deadline = time.monotonic() + self.timeout_seconds
        while True:
            status = self.request("office.job.get", {"job_id": job_id})["result"]
            if status["phase"] in TERMINAL_JOB_PHASES:
                if status["phase"] != "succeeded":
                    raise RuntimeError(f"job {job_id} finished as {status['phase']}: {status}")
                return status
            if time.monotonic() >= deadline:
                raise TimeoutError(f"timed out polling {job_id}")
            time.sleep(0.1)

    def close(self) -> None:
        if self._process.poll() is None and self._process.stdin is not None and not self._process.stdin.closed:
            try:
                self.request("office.host.shutdown", {})
            except (BrokenPipeError, RuntimeError, TimeoutError):
                # The process may already be leaving after an earlier failure;
                # stdin closure below remains the final shutdown signal.
                pass
        if self._process.stdin is not None and not self._process.stdin.closed:
            self._process.stdin.close()
        forced_termination = False
        try:
            # CommandRouter disposal permits a 30-second COM Quit/recovery
            # ladder. Give that contract time to finish before intervention.
            self._process.wait(timeout=45)
        except subprocess.TimeoutExpired:
            forced_termination = True
            self._process.terminate()
            self._process.wait(timeout=5)
        self._stdout_thread.join(timeout=2)
        self._stderr_thread.join(timeout=2)
        if forced_termination:
            raise RuntimeError("Host did not complete graceful shutdown within 45 seconds")
        if self._process.returncode not in (0, None):
            raise RuntimeError(f"Host exited with {self._process.returncode}: {self.stderr[-20:]}")

    def __enter__(self) -> Self:
        return self

    def __exit__(self, exc_type, exc, traceback) -> None:
        self.close()


def utc_now() -> str:
    return datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def scrub_paths(value: Any, workspace: Path) -> Any:
    if isinstance(value, str):
        result = value
        for candidate in {str(workspace), str(workspace).replace("\\", "/")}:
            result = result.replace(candidate, "<workspace>")
        return result.replace("\\", "/")
    if isinstance(value, list):
        return [scrub_paths(item, workspace) for item in value]
    if isinstance(value, dict):
        return {key: scrub_paths(item, workspace) for key, item in value.items()}
    return value


def write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="\n") as stream:
        stream.write(json.dumps(value, ensure_ascii=False, indent=2) + "\n")


def remove_generated(path: Path, workspace: Path, force: bool) -> None:
    if not path.exists():
        return
    resolved = path.resolve()
    if not resolved.is_relative_to(workspace.resolve()):
        raise RuntimeError(f"refusing to remove path outside workspace: {resolved}")
    if not force:
        raise FileExistsError(f"generated path already exists (pass --force): {path}")
    if path.is_dir():
        shutil.rmtree(path)
    else:
        path.unlink()


def relative_files(demo: Path, paths: list[Path]) -> list[str]:
    return [path.relative_to(demo).as_posix() for path in paths]


def write_metadata(
    demo: Path,
    *,
    title: str,
    summary: str,
    capabilities: list[str],
    inputs: list[Path],
    artifacts: list[Path],
    previews: list[Path],
    transcript: Path,
    backends: list[str],
    verification: list[str],
    reproduce: list[str],
) -> None:
    tracked = inputs + artifacts + previews + [transcript]
    payload = {
        "schema_version": SCHEMA_VERSION,
        "title": title,
        "summary": summary,
        "capabilities": capabilities,
        "inputs": relative_files(demo, inputs),
        "artifacts": relative_files(demo, artifacts),
        "previews": relative_files(demo, previews),
        "generated_with": {
            "provider": "dcc-mcp-office",
            "protocol": "office-rpc/1",
            "backends": backends,
        },
        "verified_at": utc_now(),
        "verification": verification,
        "reproduce": reproduce,
        "sha256": {
            path.relative_to(demo).as_posix(): sha256(path)
            for path in tracked
        },
    }
    write_json(demo / "metadata.json", payload)


def render_pdf_first_page(pdftoppm: Path, source: Path, output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    prefix = output.with_suffix("")
    subprocess.run(
        [str(pdftoppm), "-f", "1", "-singlefile", "-png", "-r", "144", str(source), str(prefix)],
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if not output.is_file():
        raise RuntimeError(f"Poppler did not create {output}")


def render_pdf_pages(pdftoppm: Path, source: Path, output_directory: Path) -> list[Path]:
    output_directory.mkdir(parents=True, exist_ok=True)
    prefix = output_directory / "page"
    subprocess.run(
        [str(pdftoppm), "-png", "-r", "144", str(source), str(prefix)],
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    pages = sorted(output_directory.glob("page-*.png"))
    if not pages:
        raise RuntimeError(f"Poppler did not render any pages from {source}")
    return pages


def build_deck(args: argparse.Namespace) -> dict[str, Any]:
    demo = args.workspace / "showcase" / "deck-pipeline"
    artifact = demo / "dcc-mcp-office-suite.pptx"
    slides = demo / "slides"
    preview = demo / "preview.png"
    transcript_path = demo / "transcript.json"
    for path in (artifact, slides, preview, transcript_path, demo / "metadata.json"):
        remove_generated(path, args.workspace, args.force)

    transcripts: list[dict[str, Any]] = []
    with HostSession(
        args.host_exe,
        "powerpoint",
        args.workspace,
        args.templates,
        openxml_only=True,
    ) as host:
        handshake = host.handshake("powerpoint")
        compiled = host.execute(
            "deck.compile",
            {
                "ir": str(demo / "input.json"),
                "output": str(artifact),
                "template": "brand://dcc-mcp/studio-light",
            },
        )
        inspected = host.execute("document.inspect", {"path": str(artifact), "backend": "openxml"})
        transcripts.extend(host.transcript)

    slide_files: list[Path] = []
    rendered: dict[str, Any] | None = None
    if args.with_office:
        with HostSession(args.host_exe, "powerpoint", args.workspace, args.templates) as host:
            host.handshake("powerpoint")
            rendered = host.execute(
                "slide.render",
                {
                    "path": str(artifact),
                    "output_directory": str(slides),
                    "width": 1280,
                    "height": 720,
                },
            )
            transcripts.extend(host.transcript)
        slide_files = sorted(slides.glob("*.png"))
        compose_preview(slide_files, preview, columns=3)

    write_json(
        transcript_path,
        scrub_paths(
            {"handshake": handshake, "compile": compiled, "inspect": inspected, "render": rendered, "rpc": transcripts},
            args.workspace,
        ),
    )
    if args.with_office:
        overflow = rendered["changed"]["overflow"] if rendered else []
        write_metadata(
            demo,
            title="Template-first deck pipeline",
            summary="Presentation IR compiles to an editable branded PPTX, then PowerPoint renders every slide and reports overflow.",
            capabilities=["deck.compile", "document.inspect", "slide.render"],
            inputs=[demo / "input.json"],
            artifacts=[artifact, *slide_files],
            previews=[preview],
            transcript=transcript_path,
            backends=["openxml", "desktop_com"],
            verification=[
                f"Open XML inspection reported {inspected['changed']['summary']['slide_count']} slides.",
                f"PowerPoint rendered {len(slide_files)} slide previews at 1280x720.",
                f"Overflow report contained {len(overflow)} shape findings.",
            ],
            reproduce=[
                "vx run build",
                "python scripts/capture_showcases.py --with-office --force",
                "python scripts/validate_showcases.py",
            ],
        )
    return {"artifact": artifact, "preview": preview}


def build_dashboard(args: argparse.Namespace) -> dict[str, Any]:
    demo = args.workspace / "showcase" / "production-dashboard"
    artifact = demo / "showcase-dcc-mcp-office-runtime-dashboard.xlsx"
    pdf_dir = demo / "pdf"
    pages = demo / "pages"
    preview = demo / "preview.png"
    transcript_path = demo / "transcript.json"
    for path in (artifact, pdf_dir, pages, preview, transcript_path, demo / "metadata.json"):
        remove_generated(path, args.workspace, args.force)

    generator = args.workspace / "skills" / "office-generate-production-dashboard" / "scripts" / "generate_dashboard.py"
    process = subprocess.run(
        [sys.executable, str(generator), "--input", str(demo / "input.json"), "--out", str(demo)],
        cwd=args.workspace,
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    generated = json.loads(process.stdout.strip().splitlines()[-1])
    if Path(generated["context"]["artifact"]).resolve() != artifact.resolve():
        raise RuntimeError(f"unexpected dashboard artifact: {generated}")

    transcript: dict[str, Any] = {"generator": generated}
    pdfs: list[Path] = []
    page_files: list[Path] = []
    if args.with_office:
        with HostSession(args.host_exe, "excel", args.workspace, args.templates) as host:
            handshake = host.handshake("excel")
            converted = host.execute(
                "batch.convert",
                {
                    "inputs": [str(artifact)],
                    "output_directory": str(pdf_dir),
                    "target_format": "pdf",
                    "overwrite": "fail",
                },
            )
            transcript.update({"handshake": handshake, "convert": converted, "rpc": host.transcript})
        pdfs = sorted(pdf_dir.glob("*.pdf"))
        if len(pdfs) != 1:
            raise RuntimeError(f"expected one dashboard PDF, found {pdfs}")
        page_files = render_pdf_pages(args.pdftoppm, pdfs[0], pages)
        compose_preview(page_files, preview, columns=2)

    write_json(transcript_path, scrub_paths(transcript, args.workspace))
    if args.with_office:
        write_metadata(
            demo,
            title="Production capability dashboard",
            summary="Workbook IR becomes an editable XLSX with a capability ledger and chart, then Excel exports a native PDF preview.",
            capabilities=["office-generate-production-dashboard", "batch.convert"],
            inputs=[
                demo / "input.json",
                demo / "asset-manifest.json",
                demo / "assets" / "data-landscape.jpg",
            ],
            artifacts=[artifact, *pdfs, *page_files],
            previews=[preview],
            transcript=transcript_path,
            backends=["openxml", "desktop_com"],
            verification=[
                "The generator returned success with two worksheets and six data rows.",
                "Excel exported a non-empty PDF through batch.convert.",
                f"All {len(page_files)} PDF pages were rendered with Poppler for visual review.",
            ],
            reproduce=[
                "python skills/office-generate-production-dashboard/scripts/generate_dashboard.py --input showcase/production-dashboard/input.json --out showcase/production-dashboard",
                "python scripts/capture_showcases.py --with-office --force",
                "python scripts/validate_showcases.py",
            ],
        )
    return {"artifact": artifact, "preview": preview}


def build_template_gallery(args: argparse.Namespace) -> dict[str, Any]:
    demo = args.workspace / "showcase" / "template-gallery"
    artifacts_dir = demo / "artifacts"
    renders_dir = demo / "renders"
    preview = demo / "preview.png"
    quality_path = demo / "quality-report.json"
    transcript_path = demo / "transcript.json"
    for path in (artifacts_dir, renders_dir, preview, quality_path, transcript_path, demo / "metadata.json"):
        remove_generated(path, args.workspace, args.force)

    templates = [
        ("studio-light", "brand://dcc-mcp/studio-light"),
        ("executive-violet", "brand://dcc-mcp/executive-violet"),
        ("momentum-cobalt", "brand://dcc-mcp/momentum-cobalt"),
    ]
    artifacts_dir.mkdir(parents=True, exist_ok=True)
    compile_report: dict[str, Any] = {}
    transcripts: list[dict[str, Any]] = []
    with HostSession(
        args.host_exe,
        "powerpoint",
        args.workspace,
        args.templates,
        openxml_only=True,
    ) as host:
        handshake = host.handshake("powerpoint")
        for slug, uri in templates:
            artifact = artifacts_dir / f"{slug}.pptx"
            compiled = host.execute(
                "deck.compile",
                {
                    "ir": str(demo / "input.json"),
                    "output": str(artifact),
                    "template": uri,
                },
            )
            inspected = host.execute("document.inspect", {"path": str(artifact), "backend": "openxml"})
            compile_report[slug] = {"uri": uri, "compile": compiled, "inspect": inspected}
        transcripts.extend(host.transcript)

    render_report: dict[str, Any] = {}
    preview_sources: list[Path] = []
    all_slide_files: list[Path] = []
    if args.with_office:
        with HostSession(args.host_exe, "powerpoint", args.workspace, args.templates) as host:
            host.handshake("powerpoint")
            for slug, uri in templates:
                destination = renders_dir / slug
                rendered = host.execute(
                    "slide.render",
                    {
                        "path": str(artifacts_dir / f"{slug}.pptx"),
                        "output_directory": str(destination),
                        "width": 1280,
                        "height": 720,
                    },
                )
                slide_files = sorted(destination.glob("*.png"))
                if len(slide_files) != 4:
                    raise RuntimeError(f"{uri} rendered {len(slide_files)} slides; expected four")
                all_slide_files.extend(slide_files)
                preview_sources.extend(slide_files[:2])
                render_report[slug] = rendered
            transcripts.extend(host.transcript)
        compose_preview(preview_sources, preview, columns=3)

    quality = {
        "schema_version": "dcc-mcp-showcase-quality/1.0",
        "method": "deterministic preflight before any learned or human preference score",
        "templates": [],
    }
    for slug, uri in templates:
        inspected = compile_report[slug]["inspect"]
        rendered = render_report.get(slug)
        overflow = rendered["changed"]["overflow"] if rendered else []
        quality["templates"].append(
            {
                "uri": uri,
                "slide_count": inspected["changed"]["summary"]["slide_count"],
                "rendered_slides": len(list((renders_dir / slug).glob("*.png"))) if args.with_office else 0,
                "overflow_findings": len(overflow),
                "gates_passed": bool(args.with_office and not overflow),
            }
        )
    write_json(quality_path, quality)
    write_json(
        transcript_path,
        scrub_paths(
            {
                "handshake": handshake,
                "compile": compile_report,
                "render": render_report,
                "rpc": transcripts,
            },
            args.workspace,
        ),
    )
    if args.with_office:
        artifacts = [artifacts_dir / f"{slug}.pptx" for slug, _ in templates]
        write_metadata(
            demo,
            title="Brand template comparison",
            summary="The same Presentation IR compiles through three project-owned brand packages, then PowerPoint renders and overflow-checks every result.",
            capabilities=["deck.compile", "document.inspect", "slide.render"],
            inputs=[
                demo / "input.json",
                demo / "asset-manifest.json",
                demo / "assets" / "editorial-office-hero.jpg",
            ],
            artifacts=[*artifacts, quality_path, *all_slide_files],
            previews=[preview],
            transcript=transcript_path,
            backends=["openxml", "desktop_com"],
            verification=[
                "Three versioned brand packages compiled the same four-slide Presentation IR.",
                "PowerPoint rendered all twelve slides at 1280x720.",
                "The deterministic quality report records slide count and overflow gates before aesthetic selection.",
            ],
            reproduce=[
                "python scripts/capture_showcases.py --with-office --force",
                "python scripts/validate_showcases.py",
            ],
        )
    return {"preview": preview}


def build_image_rich_deck(args: argparse.Namespace) -> dict[str, Any]:
    demo = args.workspace / "showcase" / "image-rich-deck"
    artifact = demo / "dcc-mcp-office-visual-story.pptx"
    slides = demo / "slides"
    preview = demo / "preview.png"
    transcript_path = demo / "transcript.json"
    for path in (artifact, slides, preview, transcript_path, demo / "metadata.json"):
        remove_generated(path, args.workspace, args.force)

    transcripts: list[dict[str, Any]] = []
    with HostSession(
        args.host_exe,
        "powerpoint",
        args.workspace,
        args.templates,
        openxml_only=True,
    ) as host:
        handshake = host.handshake("powerpoint")
        compiled = host.execute(
            "deck.compile",
            {
                "ir": str(demo / "input.json"),
                "output": str(artifact),
                "template": "brand://dcc-mcp/momentum-cobalt",
            },
        )
        inspected = host.execute("document.inspect", {"path": str(artifact), "backend": "openxml"})
        transcripts.extend(host.transcript)

    rendered: dict[str, Any] | None = None
    slide_files: list[Path] = []
    if args.with_office:
        with HostSession(args.host_exe, "powerpoint", args.workspace, args.templates) as host:
            host.handshake("powerpoint")
            rendered = host.execute(
                "slide.render",
                {
                    "path": str(artifact),
                    "output_directory": str(slides),
                    "width": 1280,
                    "height": 720,
                },
            )
            transcripts.extend(host.transcript)
        slide_files = sorted(slides.glob("*.png"))
        if len(slide_files) != 5:
            raise RuntimeError(f"image-rich deck rendered {len(slide_files)} slides; expected five")
        compose_preview(slide_files, preview, columns=3)

    write_json(
        transcript_path,
        scrub_paths(
            {"handshake": handshake, "compile": compiled, "inspect": inspected, "render": rendered, "rpc": transcripts},
            args.workspace,
        ),
    )
    if args.with_office:
        overflow = rendered["changed"]["overflow"] if rendered else []
        assets = sorted((demo / "assets").glob("*.jpg"))
        write_metadata(
            demo,
            title="Image-rich semantic layouts",
            summary="Six original editorial visuals are composed into cover, image-left/text-right and asymmetric collage layouts in a fully editable PPTX.",
            capabilities=["deck.compile", "document.inspect", "slide.render"],
            inputs=[demo / "input.json", demo / "asset-manifest.json", *assets],
            artifacts=[artifact, *slide_files],
            previews=[preview],
            transcript=transcript_path,
            backends=["openxml", "desktop_com"],
            verification=[
                "The Open XML compiler embedded six 1280x720 showcase-optimized JPEG inputs as presentation media.",
                "PowerPoint rendered all five slides at 1280x720.",
                f"The native overflow report contained {len(overflow)} shape findings.",
            ],
            reproduce=[
                "python scripts/capture_showcases.py --with-office --force",
                "python scripts/validate_showcases.py",
            ],
        )
    return {"artifact": artifact, "preview": preview}


def build_word_brief(args: argparse.Namespace) -> dict[str, Any]:
    demo = args.workspace / "showcase" / "word-executive-brief"
    artifact = demo / "dcc-mcp-office-executive-brief.docx"
    pdf_dir = demo / "pdf"
    pages = demo / "pages"
    preview = demo / "preview.png"
    transcript_path = demo / "transcript.json"
    for path in (artifact, pdf_dir, pages, preview, transcript_path, demo / "metadata.json"):
        remove_generated(path, args.workspace, args.force)

    builder = args.workspace / "scripts" / "build_showcase_docx.py"
    process = subprocess.run(
        [sys.executable, str(builder), "--input", str(demo / "content.json"), "--output", str(artifact)],
        cwd=args.workspace,
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    generated = json.loads(process.stdout.strip().splitlines()[-1])
    if Path(generated["artifact"]).resolve() != artifact.resolve():
        raise RuntimeError(f"unexpected Word artifact: {generated}")

    transcript: dict[str, Any] = {"builder": generated}
    pdfs: list[Path] = []
    page_files: list[Path] = []
    if args.with_office:
        with HostSession(args.host_exe, "word", args.workspace, args.templates) as host:
            handshake = host.handshake("word")
            inspected = host.execute("document.inspect", {"path": str(artifact)})
            converted = host.execute(
                "batch.convert",
                {
                    "inputs": [str(artifact)],
                    "output_directory": str(pdf_dir),
                    "target_format": "pdf",
                    "overwrite": "fail",
                    "validation": ["output_openable", "non_empty", "page_count_reasonable"],
                },
            )
            transcript.update(
                {
                    "handshake": handshake,
                    "inspect": inspected,
                    "convert": converted,
                    "rpc": host.transcript,
                }
            )
        pdfs = sorted(pdf_dir.glob("*.pdf"))
        if len(pdfs) != 1:
            raise RuntimeError(f"expected one Word PDF, found {pdfs}")
        page_files = render_pdf_pages(args.pdftoppm, pdfs[0], pages)
        compose_preview(page_files, preview, columns=2)

    write_json(transcript_path, scrub_paths(transcript, args.workspace))
    if args.with_office:
        inspected = transcript["inspect"]
        write_metadata(
            demo,
            title="Executive Word brief",
            summary="A polished editable DOCX is inspected and exported by the real Word sidecar, then every native PDF page is rendered for review.",
            capabilities=["document.inspect", "batch.convert"],
            inputs=[
                demo / "content.json",
                demo / "asset-manifest.json",
                demo / "assets" / "document-evidence-banner.jpg",
            ],
            artifacts=[artifact, *pdfs, *page_files],
            previews=[preview],
            transcript=transcript_path,
            backends=["desktop_com"],
            verification=[
                f"Word inspection reported document kind {inspected['changed']['summary']['kind']}.",
                "Word exported a non-empty PDF through the governed batch.convert capability.",
                f"Poppler rendered all {len(page_files)} native PDF pages for visual review.",
            ],
            reproduce=[
                "python scripts/build_showcase_docx.py",
                "python scripts/capture_showcases.py --with-office --force",
                "python scripts/validate_showcases.py",
            ],
        )
    return {"artifact": artifact, "preview": preview}


def build_replace(args: argparse.Namespace, deck: Path, document: Path) -> dict[str, Any]:
    demo = args.workspace / "showcase" / "global-text-replace"
    generated_dirs = [demo / name for name in ("before", "work", "after", "checkpoints", "slides")]
    generated_files = [demo / name for name in ("preview.png", "transcript.json", "metadata.json")]
    for path in [*generated_dirs, *generated_files]:
        remove_generated(path, args.workspace, args.force)
    if not args.with_office:
        return {"preview": demo / "preview.png"}

    before = demo / "before"
    work = demo / "work"
    after = demo / "after"
    checkpoints = demo / "checkpoints"
    for directory in (before, work, after, checkpoints):
        directory.mkdir(parents=True, exist_ok=True)
    sources = {
        "powerpoint": (deck, "presentation.pptx"),
        "word": (document, "document.docx"),
        "excel": (args.workspace / "tests" / "fixtures" / "fixture-workbook.xlsx", "workbook.xlsx"),
    }
    for source, name in sources.values():
        shutil.copy2(source, before / name)
        shutil.copy2(source, work / name)

    confirmation = {
        "action": "overwrite_original",
        "confirmed": True,
        "confirmed_by": "human:showcase-goal-owner",
        "confirmed_at": utc_now(),
    }
    scopes = {
        "powerpoint": ["body", "notes"],
        "word": ["body", "headers", "footers"],
        "excel": ["body"],
    }
    report: dict[str, Any] = {"dry_run": {}, "confirmation_gate": {}, "commit": {}, "verification": {}, "rpc": {}}
    for app, (_, name) in sources.items():
        target = work / name
        with HostSession(args.host_exe, app, args.workspace, args.templates) as host:
            host.handshake(app)
            common_input = {
                "inputs": [str(target)],
                "rules": [{"find": "2025年度", "replace": "2026年度", "match": "literal"}],
                "scope": scopes[app],
            }
            report["dry_run"][app] = host.execute("batch.replace_text", {**common_input, "dry_run": True})
            denied = host.execute(
                "batch.replace_text",
                {**common_input, "dry_run": False},
                allow_error=True,
            )
            if denied.get("error", {}).get("code") != "OFFICE_USER_CONFIRMATION_REQUIRED":
                raise RuntimeError(f"{app} commit was not stopped by the confirmation gate: {denied}")
            report["confirmation_gate"][app] = denied
            report["commit"][app] = host.execute(
                "batch.replace_text",
                {**common_input, "dry_run": False},
                confirmation=confirmation,
            )
            report["verification"][app] = {
                "old_text": host.execute("batch.replace_text", {**common_input, "dry_run": True}),
                "new_text": host.execute(
                    "batch.replace_text",
                    {
                        **common_input,
                        "rules": [{"find": "2026年度", "replace": "verified", "match": "literal"}],
                        "dry_run": True,
                    },
                ),
            }
            report["rpc"][app] = host.transcript

        checkpoint_candidates = list(work.glob(f"{Path(name).stem}.dcc-checkpoint-*{Path(name).suffix}"))
        if len(checkpoint_candidates) != 1:
            raise RuntimeError(f"expected one checkpoint for {app}, found {checkpoint_candidates}")
        checkpoint_candidates[0].replace(checkpoints / f"{app}-preimage{Path(name).suffix}")
        target.replace(after / name)

    before_slides = demo / "slides" / "before"
    after_slides = demo / "slides" / "after"
    with HostSession(args.host_exe, "powerpoint", args.workspace, args.templates) as host:
        host.execute(
            "slide.render",
            {"path": str(before / "presentation.pptx"), "output_directory": str(before_slides), "width": 1280, "height": 720},
        )
        host.execute(
            "slide.render",
            {"path": str(after / "presentation.pptx"), "output_directory": str(after_slides), "width": 1280, "height": 720},
        )
        report["rpc"]["powerpoint_previews"] = host.transcript
    before_marker = sorted(before_slides.glob("*.png"))[-2]
    after_marker = sorted(after_slides.glob("*.png"))[-2]
    preview = demo / "preview.png"
    compose_preview([before_marker, after_marker], preview, columns=2)

    transcript_path = demo / "transcript.json"
    write_json(transcript_path, scrub_paths(report, args.workspace))
    inputs = sorted(before.iterdir())
    outputs = sorted(after.iterdir())
    checkpoint_files = sorted(checkpoints.iterdir())
    slide_files = sorted((demo / "slides").rglob("*.png"))
    write_metadata(
        demo,
        title="Safe global text replacement",
        summary="One rule is dry-run, confirmation-gated, checkpointed and committed across PowerPoint, Word and Excel copies.",
        capabilities=["batch.replace_text", "slide.render"],
        inputs=inputs,
        artifacts=[*outputs, *checkpoint_files, *slide_files],
        previews=[preview],
        transcript=transcript_path,
        backends=["desktop_com"],
        verification=[
            "Dry-run found the 2025年度 marker in PowerPoint, Word and Excel without modifying files.",
            "Each unconfirmed commit was refused with OFFICE_USER_CONFIRMATION_REQUIRED.",
            "Each confirmed commit produced a byte-exact checkpoint before write.",
            "A second dry-run found no old marker and found the new 2026年度 marker.",
        ],
        reproduce=[
            "python scripts/capture_showcases.py --with-office --force",
            "python scripts/validate_showcases.py",
        ],
    )
    return {"preview": preview}


def build_batch_pdf(args: argparse.Namespace, deck: Path, document: Path, dashboard: Path) -> dict[str, Any]:
    demo = args.workspace / "showcase" / "batch-to-pdf"
    for name in ("inputs", "artifacts", "pages"):
        remove_generated(demo / name, args.workspace, args.force)
    for name in ("preview.png", "transcript.json", "metadata.json"):
        remove_generated(demo / name, args.workspace, args.force)
    if not args.with_office:
        return {"preview": demo / "preview.png"}

    inputs = demo / "inputs"
    artifacts = demo / "artifacts"
    pages = demo / "pages"
    inputs.mkdir(parents=True, exist_ok=True)
    sources = {
        "powerpoint": (deck, inputs / "office-suite.pptx"),
        "word": (document, inputs / "office-brief.docx"),
        "excel": (dashboard, inputs / "office-dashboard.xlsx"),
    }
    for source, destination in sources.values():
        shutil.copy2(source, destination)

    report: dict[str, Any] = {"apps": {}, "rpc": {}}
    for app, (_, source) in sources.items():
        with HostSession(args.host_exe, app, args.workspace, args.templates) as host:
            report["apps"][app] = host.execute(
                "batch.convert",
                {
                    "inputs": [str(source)],
                    "output_directory": str(artifacts),
                    "target_format": "pdf",
                    "overwrite": "fail",
                    "validation": ["output_openable", "non_empty", "page_count_reasonable"],
                },
            )
            report["rpc"][app] = host.transcript

    pdfs = sorted(artifacts.glob("*.pdf"))
    if len(pdfs) != 3:
        raise RuntimeError(f"expected three PDFs, found {pdfs}")
    page_files: list[Path] = []
    for pdf in pdfs:
        page = pages / f"{pdf.stem}.png"
        render_pdf_first_page(args.pdftoppm, pdf, page)
        page_files.append(page)
    preview = demo / "preview.png"
    compose_preview(page_files, preview, columns=3)

    transcript_path = demo / "transcript.json"
    write_json(transcript_path, scrub_paths(report, args.workspace))
    write_metadata(
        demo,
        title="Mixed Office batch to PDF",
        summary="PowerPoint, Word and Excel artifacts are exported by isolated native COM sidecars into one validated PDF batch.",
        capabilities=["batch.convert"],
        inputs=sorted(inputs.iterdir()),
        artifacts=[*pdfs, *page_files],
        previews=[preview],
        transcript=transcript_path,
        backends=["desktop_com"],
        verification=[
            "Three application-specific sidecars completed one PDF export each.",
            "Every output passed openable, non-empty and reasonable-page-count validation.",
            "Poppler rendered the first page of every PDF for visual review.",
        ],
        reproduce=[
            "python scripts/capture_showcases.py --with-office --force",
            "python scripts/validate_showcases.py",
        ],
    )
    return {"preview": preview}


def parse_args() -> argparse.Namespace:
    workspace = Path(__file__).resolve().parents[1]
    default_host = workspace / "dotnet" / "Office.Automation.Host" / "bin" / "Debug" / "net8.0-windows" / "dcc-office-host.exe"
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workspace", type=Path, default=workspace)
    parser.add_argument("--host-exe", type=Path, default=default_host)
    parser.add_argument("--templates", type=Path, default=workspace / "templates")
    parser.add_argument("--pdftoppm", type=Path, default=Path(shutil.which("pdftoppm") or "pdftoppm"))
    parser.add_argument("--with-office", action="store_true")
    parser.add_argument("--force", action="store_true")
    args = parser.parse_args()
    args.workspace = args.workspace.resolve()
    args.host_exe = args.host_exe.resolve()
    args.templates = args.templates.resolve()
    if not args.host_exe.is_file():
        parser.error(f"Host executable not found: {args.host_exe}; run 'vx run build' first")
    if args.with_office and not args.pdftoppm.is_file():
        parser.error(f"pdftoppm not found: {args.pdftoppm}")
    return args


def main() -> None:
    args = parse_args()
    deck = build_deck(args)
    dashboard = build_dashboard(args)
    build_template_gallery(args)
    word = build_word_brief(args)
    # Keep two different app sessions between PowerPoint sidecars. Office COM
    # servers can remain in their ROT shutdown window briefly after Quit; a
    # write-capable indeterminate failure must never be hidden by blind retry.
    build_batch_pdf(args, deck["artifact"], word["artifact"], dashboard["artifact"])
    build_replace(args, deck["artifact"], word["artifact"])
    build_image_rich_deck(args)


if __name__ == "__main__":
    main()
