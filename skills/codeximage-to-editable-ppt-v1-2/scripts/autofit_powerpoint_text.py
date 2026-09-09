#!/usr/bin/env python3
"""Apply bounded PowerPoint-native text overflow repair.

The repair is deliberately conservative: it never moves shapes or changes
slide geometry. It enables PowerPoint text-to-fit and, when a text range has a
uniform measurable font size, reduces that size only enough to fit. The script
writes a JSON audit and returns 0 when it completed successfully.
"""

from __future__ import annotations

import argparse
import json
import shutil
import sys
from pathlib import Path
from typing import Any, Iterable


def _items(collection: Any) -> Iterable[Any]:
    for index in range(1, int(collection.Count) + 1):
        yield collection.Item(index)


def _float(value: Any, default: float = 0.0) -> float:
    try:
        return float(value)
    except Exception:
        return default


def _repair_frame(slide_no: int, shape: Any, frame: Any, tolerance: float, min_font: float) -> dict | None:
    try:
        if not bool(frame.HasText):
            return None
        text_range = frame.TextRange
        text = str(text_range.Text or "").replace("\r", "\n").strip()
        if not text:
            return None
        available_w = max(1.0, _float(shape.Width) - _float(frame.MarginLeft) - _float(frame.MarginRight))
        available_h = max(1.0, _float(shape.Height) - _float(frame.MarginTop) - _float(frame.MarginBottom))
        bound_w = max(0.0, _float(text_range.BoundWidth))
        bound_h = max(0.0, _float(text_range.BoundHeight))
        ratio = max(bound_w / available_w, bound_h / available_h)
        if ratio <= tolerance:
            return None
        before_size = _float(text_range.Font.Size, -1.0)
        frame.AutoSize = 2  # msoAutoSizeTextToFitShape
        new_size: float | None = None
        if before_size > 0:
            new_size = max(min_font, before_size / (ratio * 1.02))
            if new_size < before_size:
                text_range.Font.Size = new_size
        return {
            "slide": slide_no,
            "shape_id": int(getattr(shape, "Id", 0)),
            "shape_name": str(getattr(shape, "Name", "")),
            "text": text[:160],
            "overflow_ratio_before": round(ratio, 4),
            "font_size_before": round(before_size, 3) if before_size > 0 else None,
            "font_size_requested": round(new_size, 3) if new_size is not None else None,
            "repair": "powerpoint_text_to_fit_shape",
        }
    except Exception:
        return None


def _repair_shape(slide_no: int, shape: Any, tolerance: float, min_font: float) -> list[dict]:
    changes: list[dict] = []
    try:
        if bool(shape.HasTextFrame):
            change = _repair_frame(slide_no, shape, shape.TextFrame2, tolerance, min_font)
            if change:
                changes.append(change)
    except Exception:
        pass
    try:
        if bool(shape.HasTable):
            table = shape.Table
            for row in range(1, int(table.Rows.Count) + 1):
                for col in range(1, int(table.Columns.Count) + 1):
                    cell = table.Cell(row, col).Shape
                    change = _repair_frame(slide_no, cell, cell.TextFrame2, tolerance, min_font)
                    if change:
                        changes.append(change)
    except Exception:
        pass
    try:
        for child in _items(shape.GroupItems):
            changes.extend(_repair_shape(slide_no, child, tolerance, min_font))
    except Exception:
        pass
    return changes


def repair(pptx: Path, output: Path | None, inplace: bool, tolerance: float, min_font: float) -> tuple[Path, list[dict]]:
    if sys.platform != "win32":
        raise RuntimeError("PowerPoint text repair requires Windows")
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("pywin32 is required") from exc
    source = pptx.resolve()
    if inplace:
        target = source
    else:
        if output is None:
            target = source.with_name(f"{source.stem}_autofit{source.suffix}")
        else:
            target = output.resolve()
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, target)
    app = None
    presentation = None
    changes: list[dict] = []
    pythoncom.CoInitialize()
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        app.DisplayAlerts = 1
        presentation = app.Presentations.Open(str(target), False, False, False)
        for slide_no in range(1, int(presentation.Slides.Count) + 1):
            slide = presentation.Slides(slide_no)
            for shape in _items(slide.Shapes):
                changes.extend(_repair_shape(slide_no, shape, tolerance, min_font))
        presentation.Save()
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
        pythoncom.CoUninitialize()
    return target, changes


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Repair PowerPoint text overflow without moving slide objects.")
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--inplace", action="store_true")
    parser.add_argument("--text-tolerance", type=float, default=1.03)
    parser.add_argument("--min-font-size", type=float, default=5.0)
    parser.add_argument("--report", type=Path)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.inplace and args.output:
        raise SystemExit("Use either --inplace or --output, not both")
    try:
        target, changes = repair(args.pptx, args.output, args.inplace, args.text_tolerance, args.min_font_size)
        payload = {"status": "completed", "pptx": str(target), "changed_count": len(changes), "changes": changes}
        report = args.report or target.with_name(f"{target.stem}_autofit_report.json")
        report.parent.mkdir(parents=True, exist_ok=True)
        report.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0
    except Exception as exc:
        payload = {"status": "error", "error": str(exc)}
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
