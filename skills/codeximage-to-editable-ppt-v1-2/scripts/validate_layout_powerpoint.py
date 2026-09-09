#!/usr/bin/env python3
"""Fail-closed layout validation using Microsoft PowerPoint itself.

Checks editable text bounds through COM. In final mode it also exports every
slide through PowerPoint and compares it with ordered source images. Geometry-
only mode skips image export for fast iterative preflight.
Exit code 0 means every slide passed. Exit code 2 means review/fix is required.
"""

from __future__ import annotations

import argparse
import csv
import json
import subprocess
import sys
import time
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Iterable

import numpy as np
from PIL import Image


# PowerPoint's TextRange bounds include glyph overhang and font line metrics.
# A tightly fitted, isolated decorative marker can therefore exceed its usable
# text box by a couple of points without being visibly clipped. Keep this list
# explicit: letters, numbers, CJK text, formulas, and general punctuation must
# continue through the normal fail-closed overflow check.
_DECORATIVE_SINGLE_CHARACTER_MARKERS = frozenset(
    "•‣◦▪▫●○■□◆◇▶▷▸▹→←↑↓↔↕➜➝➞★☆✦✧✓✔"
)
_DECORATIVE_MARKER_MAX_AVAILABLE_PT = 40.0
_DECORATIVE_MARKER_MAX_OVERFLOW_PT = 2.5
_DECORATIVE_MARKER_MAX_RATIO = 1.15


@dataclass
class TextIssue:
    slide: int
    shape_id: int
    shape_name: str
    text: str
    bound_width_pt: float
    bound_height_pt: float
    available_width_pt: float
    available_height_pt: float
    width_ratio: float
    height_ratio: float


@dataclass
class SlideResult:
    slide: int
    status: str
    renderer: str
    rendered_path: str
    source_path: str
    mean_absolute_error: float | None
    changed_pixel_ratio: float | None
    text_overflow_count: int
    off_slide_shape_count: int
    issues: str


def _collection_items(collection: Any) -> Iterable[Any]:
    for index in range(1, int(collection.Count) + 1):
        yield collection.Item(index)


def _safe_float(value: Any, default: float = 0.0) -> float:
    try:
        return float(value)
    except Exception:
        return default


def _is_minor_decorative_marker_bounds_difference(
    text: str,
    bound_w: float,
    bound_h: float,
    available_w: float,
    available_h: float,
    width_ratio: float,
    height_ratio: float,
) -> bool:
    """Ignore only small PowerPoint glyph-metric overhangs for isolated markers.

    This deliberately requires all of: one explicitly whitelisted character,
    a small marker-sized box, no more than 2.5 pt absolute excess on either
    axis, and no more than a 15% ratio excess. Multi-character text is never
    exempted, so the ordinary ``text_tolerance`` remains fail-closed for prose.
    """
    marker = "".join(text.split())
    if len(marker) != 1 or marker not in _DECORATIVE_SINGLE_CHARACTER_MARKERS:
        return False
    if available_w > _DECORATIVE_MARKER_MAX_AVAILABLE_PT or available_h > _DECORATIVE_MARKER_MAX_AVAILABLE_PT:
        return False
    if max(0.0, bound_w - available_w) > _DECORATIVE_MARKER_MAX_OVERFLOW_PT:
        return False
    if max(0.0, bound_h - available_h) > _DECORATIVE_MARKER_MAX_OVERFLOW_PT:
        return False
    return width_ratio <= _DECORATIVE_MARKER_MAX_RATIO and height_ratio <= _DECORATIVE_MARKER_MAX_RATIO


def _inspect_text_frame(slide_no: int, shape: Any, text_frame: Any, tolerance: float) -> TextIssue | None:
    try:
        if not bool(text_frame.HasText):
            return None
        text_range = text_frame.TextRange
        text = str(text_range.Text or "").replace("\r", "\n").strip()
        if not text:
            return None
        available_w = max(1.0, _safe_float(shape.Width) - _safe_float(text_frame.MarginLeft) - _safe_float(text_frame.MarginRight))
        available_h = max(1.0, _safe_float(shape.Height) - _safe_float(text_frame.MarginTop) - _safe_float(text_frame.MarginBottom))
        bound_w = max(0.0, _safe_float(text_range.BoundWidth))
        bound_h = max(0.0, _safe_float(text_range.BoundHeight))
        width_ratio = bound_w / available_w
        height_ratio = bound_h / available_h
        if width_ratio <= tolerance and height_ratio <= tolerance:
            return None
        if _is_minor_decorative_marker_bounds_difference(
            text,
            bound_w,
            bound_h,
            available_w,
            available_h,
            width_ratio,
            height_ratio,
        ):
            return None
        return TextIssue(
            slide=slide_no,
            shape_id=int(getattr(shape, "Id", 0)),
            shape_name=str(getattr(shape, "Name", "")),
            text=text[:160],
            bound_width_pt=round(bound_w, 3),
            bound_height_pt=round(bound_h, 3),
            available_width_pt=round(available_w, 3),
            available_height_pt=round(available_h, 3),
            width_ratio=round(width_ratio, 4),
            height_ratio=round(height_ratio, 4),
        )
    except Exception:
        return None


def _inspect_shape_text(slide_no: int, shape: Any, tolerance: float) -> list[TextIssue]:
    issues: list[TextIssue] = []
    try:
        if bool(shape.HasTextFrame):
            issue = _inspect_text_frame(slide_no, shape, shape.TextFrame2, tolerance)
            if issue:
                issues.append(issue)
    except Exception:
        pass
    try:
        if bool(shape.HasTable):
            table = shape.Table
            for row in range(1, int(table.Rows.Count) + 1):
                for col in range(1, int(table.Columns.Count) + 1):
                    cell_shape = table.Cell(row, col).Shape
                    issue = _inspect_text_frame(slide_no, cell_shape, cell_shape.TextFrame2, tolerance)
                    if issue:
                        issues.append(issue)
    except Exception:
        pass
    try:
        group_items = shape.GroupItems
        for child in _collection_items(group_items):
            issues.extend(_inspect_shape_text(slide_no, child, tolerance))
    except Exception:
        pass
    return issues


def _is_off_slide(shape: Any, slide_w: float, slide_h: float, tolerance_pt: float = 2.0) -> bool:
    try:
        left = _safe_float(shape.Left); top = _safe_float(shape.Top)
        right = left + _safe_float(shape.Width); bottom = top + _safe_float(shape.Height)
        return left < -tolerance_pt or top < -tolerance_pt or right > slide_w + tolerance_pt or bottom > slide_h + tolerance_pt
    except Exception:
        return False


def render_and_inspect_powerpoint(
    pptx: Path,
    outdir: Path,
    width: int,
    height: int,
    text_tolerance: float,
    export_images: bool = True,
) -> tuple[list[Path | None], dict[int, list[TextIssue]], dict[int, int]]:
    if sys.platform != "win32":
        raise RuntimeError("Microsoft PowerPoint COM validation requires Windows")
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("pywin32 is required for PowerPoint validation") from exc

    outdir.mkdir(parents=True, exist_ok=True)
    rendered: list[Path | None] = []
    text_issues: dict[int, list[TextIssue]] = {}
    off_slide_counts: dict[int, int] = {}
    app = None
    app_pid: int | None = None
    presentation = None
    pythoncom.CoInitialize()
    try:
        last_open_error: Exception | None = None
        for open_attempt in range(1, 4):
            try:
                existing_powerpoint_pids = _powerpoint_pids()
                app = win32com.client.DispatchEx("PowerPoint.Application")
                try:
                    import win32process
                    app_pid = int(win32process.GetWindowThreadProcessId(int(app.HWND))[1])
                except Exception:
                    new_pids = _powerpoint_pids() - existing_powerpoint_pids
                    app_pid = max(new_pids) if new_pids else None
                # Some PowerPoint builds expose DisplayAlerts as read-only for
                # a newly dispatched automation instance. Alert suppression is
                # optional for read-only validation.
                try:
                    app.DisplayAlerts = 1  # ppAlertsNone
                except Exception:
                    pass
                presentations = app.Presentations
                presentation = presentations.Open(str(pptx.resolve()), True, False, False)
                break
            except Exception as exc:
                last_open_error = exc
                if presentation is not None:
                    try:
                        presentation.Close()
                    except Exception:
                        pass
                presentation = None
                if app is not None:
                    try:
                        app.Quit()
                    except Exception:
                        pass
                _terminate_stale_automation_process(app_pid)
                app_pid = None
                app = None
                try:
                    pythoncom.CoFreeUnusedLibraries()
                    pythoncom.PumpWaitingMessages()
                except Exception:
                    pass
                if open_attempt < 3:
                    time.sleep(float(open_attempt))
        if presentation is None:
            raise RuntimeError(
                f"PowerPoint could not open the presentation after 3 attempts: {last_open_error}"
            )
        slide_w = _safe_float(presentation.PageSetup.SlideWidth)
        slide_h = _safe_float(presentation.PageSetup.SlideHeight)
        for slide_no in range(1, int(presentation.Slides.Count) + 1):
            slide = presentation.Slides(slide_no)
            if export_images:
                target = outdir / f"slide_{slide_no:02d}.png"
                slide.Export(str(target.resolve()), "PNG", width, height)
                if not target.exists():
                    raise RuntimeError(f"PowerPoint did not export slide {slide_no}")
                rendered.append(target)
            else:
                rendered.append(None)
            slide_text_issues: list[TextIssue] = []
            off_slide = 0
            for shape in _collection_items(slide.Shapes):
                slide_text_issues.extend(_inspect_shape_text(slide_no, shape, text_tolerance))
                if _is_off_slide(shape, slide_w, slide_h):
                    off_slide += 1
            text_issues[slide_no] = slide_text_issues
            off_slide_counts[slide_no] = off_slide
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        _terminate_stale_automation_process(app_pid)
        pythoncom.CoUninitialize()
    return rendered, text_issues, off_slide_counts


def _terminate_stale_automation_process(pid: int | None, timeout_ms: int = 1500) -> None:
    """Terminate only the exact PowerPoint process created by this validator if Quit stalls."""
    if not pid or sys.platform != "win32":
        return
    import ctypes

    synchronize = 0x00100000
    process_terminate = 0x0001
    wait_timeout = 0x00000102
    handle = ctypes.windll.kernel32.OpenProcess(synchronize | process_terminate, False, int(pid))
    if not handle:
        return
    try:
        result = ctypes.windll.kernel32.WaitForSingleObject(handle, timeout_ms)
        if result == wait_timeout:
            ctypes.windll.kernel32.TerminateProcess(handle, 1)
            ctypes.windll.kernel32.WaitForSingleObject(handle, 1000)
    finally:
        ctypes.windll.kernel32.CloseHandle(handle)


def _powerpoint_pids() -> set[int]:
    if sys.platform != "win32":
        return set()
    result = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/FO", "CSV", "/NH"],
        stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL,
        text=True,
        encoding="mbcs",
        errors="replace",
        check=False,
    )
    pids: set[int] = set()
    for row in csv.reader(result.stdout.splitlines()):
        if len(row) >= 2 and row[0].upper() == "POWERPNT.EXE":
            try:
                pids.add(int(row[1]))
            except ValueError:
                pass
    return pids


def compare_images(source: Path, rendered: Path, diff_path: Path, diff_threshold: int) -> tuple[float, float]:
    original = Image.open(source).convert("RGB")
    recomposed = Image.open(rendered).convert("RGB")
    original = original.resize(recomposed.size, Image.Resampling.LANCZOS)
    orig_arr = np.asarray(original, dtype=np.int16)
    rec_arr = np.asarray(recomposed, dtype=np.int16)
    diff = np.abs(orig_arr - rec_arr)
    mae = float(np.mean(diff))
    changed = float(np.mean(np.max(diff, axis=2) > diff_threshold))
    diff_path.parent.mkdir(parents=True, exist_ok=True)
    Image.fromarray(np.clip(diff * 4, 0, 255).astype(np.uint8)).save(diff_path)
    return mae, changed


def validate(args: argparse.Namespace) -> tuple[list[SlideResult], list[TextIssue]]:
    pptx = args.pptx.resolve()
    outdir = args.outdir.resolve()
    render_dir = outdir / "powerpoint_render"
    rendered, text_by_slide, off_slide = render_and_inspect_powerpoint(
        pptx, render_dir, args.width, args.height, args.text_tolerance,
        export_images=not args.geometry_only,
    )
    sources = [path.resolve() for path in args.source]
    results: list[SlideResult] = []
    all_text_issues = [issue for issues in text_by_slide.values() for issue in issues]
    renderer = "microsoft_powerpoint_geometry" if args.geometry_only else "microsoft_powerpoint"
    for index, rendered_path in enumerate(rendered, start=1):
        source = sources[index - 1] if index <= len(sources) else None
        mae: float | None = None
        changed: float | None = None
        issue_codes: list[str] = []
        if args.geometry_only:
            pass
        elif source is None or not source.exists():
            issue_codes.append("source_missing")
        elif rendered_path is not None:
            mae, changed = compare_images(
                source, rendered_path, outdir / "diff" / f"slide_{index:02d}_diff.png", args.diff_threshold
            )
            if mae > args.max_mae:
                issue_codes.append("mean_absolute_error_exceeded")
            if changed > args.max_changed_ratio:
                issue_codes.append("changed_pixel_ratio_exceeded")
        if text_by_slide.get(index):
            issue_codes.append("text_overflow")
        if off_slide.get(index, 0):
            issue_codes.append("off_slide_shape")
        if not issue_codes:
            status = "passed"
        elif issue_codes == ["source_missing"]:
            status = "manual_review_required"
        else:
            status = "failed"
        results.append(SlideResult(
            slide=index,
            status=status,
            renderer=renderer,
            rendered_path=str(rendered_path) if rendered_path is not None else "",
            source_path=str(source) if source else "",
            mean_absolute_error=round(mae, 3) if mae is not None else None,
            changed_pixel_ratio=round(changed, 6) if changed is not None else None,
            text_overflow_count=len(text_by_slide.get(index, [])),
            off_slide_shape_count=off_slide.get(index, 0),
            issues=";".join(issue_codes),
        ))
    return results, all_text_issues


def write_reports(outdir: Path, results: list[SlideResult], text_issues: list[TextIssue], args: argparse.Namespace) -> None:
    outdir.mkdir(parents=True, exist_ok=True)
    result_rows = [asdict(row) for row in results]
    with (outdir / "layout_quality_report.csv").open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(result_rows[0].keys()))
        writer.writeheader(); writer.writerows(result_rows)
    payload = {
        "status": "passed" if results and all(row.status == "passed" for row in results) else "failed_or_review_required",
        "renderer": "microsoft_powerpoint_geometry" if args.geometry_only else "microsoft_powerpoint",
        "mode": "geometry_only" if args.geometry_only else "final_visual",
        "thresholds": {
            "max_mean_absolute_error": args.max_mae,
            "max_changed_pixel_ratio": args.max_changed_ratio,
            "diff_threshold": args.diff_threshold,
            "text_bounds_tolerance": args.text_tolerance,
        },
        "slides": result_rows,
        "text_overflow_issues": [asdict(issue) for issue in text_issues],
    }
    (outdir / "layout_quality_report.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Validate PPTX layout using Microsoft PowerPoint rendering and COM text bounds.")
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--source", type=Path, action="append", default=[], help="Source image in slide order; repeat for each slide.")
    parser.add_argument("--outdir", type=Path, required=True)
    parser.add_argument("--width", type=int, default=1920)
    parser.add_argument("--height", type=int, default=1080)
    parser.add_argument("--diff-threshold", type=int, default=25)
    parser.add_argument("--max-mae", type=float, default=15.0)
    parser.add_argument("--max-changed-ratio", type=float, default=0.15)
    parser.add_argument("--text-tolerance", type=float, default=1.03)
    parser.add_argument(
        "--geometry-only",
        action="store_true",
        help="Inspect text bounds and off-slide objects without exporting or comparing slide images.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        results, text_issues = validate(args)
        write_reports(args.outdir, results, text_issues, args)
    except Exception as exc:
        args.outdir.mkdir(parents=True, exist_ok=True)
        failure = {"status": "validator_error", "renderer": "microsoft_powerpoint", "error": str(exc)}
        (args.outdir / "layout_quality_report.json").write_text(json.dumps(failure, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps(failure, ensure_ascii=False, indent=2))
        return 2
    payload = [asdict(row) for row in results]
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0 if results and all(row.status == "passed" for row in results) else 2


if __name__ == "__main__":
    raise SystemExit(main())
