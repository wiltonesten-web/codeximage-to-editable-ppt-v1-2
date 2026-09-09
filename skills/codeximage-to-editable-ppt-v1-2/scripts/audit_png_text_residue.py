#!/usr/bin/env python3
"""Fail closed when non-text PNG assets contain likely residual text."""

from __future__ import annotations

import argparse
import csv
import json
import re
from pathlib import Path
from typing import Any

from PIL import Image


def load_json(path: Path, default: Any) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return default


def allowed_files(path: Path | None) -> set[str]:
    if path is None or not path.exists():
        return set()
    payload = load_json(path, [])
    if isinstance(payload, dict):
        payload = payload.get("allowed_files", [])
    return {str(item) for item in payload if str(item).strip()}


def manifest_non_text_files(root: Path) -> set[str]:
    names: set[str] = set()
    for manifest in root.rglob("visual_elements_manifest.json"):
        payload = load_json(manifest, [])
        if not isinstance(payload, list):
            continue
        for row in payload:
            if not isinstance(row, dict):
                continue
            name = str(row.get("file_name", "")).strip()
            if name and str(row.get("is_text_only", "no")).lower() != "yes":
                names.add(name)
    return names


def meaningful_token(token: str) -> bool:
    compact = re.sub(r"\s+", "", token)
    if re.search(r"[\u4e00-\u9fff]{2,}", compact):
        return True
    if re.search(r"[A-Za-z]{4,}", compact):
        return True
    if re.search(r"\d{2,}%?", compact):
        return True
    return False


def detect_tokens(path: Path, language: str, min_confidence: float) -> list[dict[str, Any]]:
    try:
        import pytesseract
        from pytesseract import Output
    except ImportError as exc:
        raise RuntimeError("pytesseract is required for PNG text-residue auditing") from exc

    image = Image.open(path).convert("RGBA")
    findings: dict[str, dict[str, Any]] = {}
    for background_name, background in (
        ("white", (255, 255, 255, 255)),
        ("black", (0, 0, 0, 255)),
    ):
        canvas = Image.new("RGBA", image.size, background)
        canvas.alpha_composite(image)
        data = pytesseract.image_to_data(
            canvas.convert("RGB"),
            lang=language,
            config="--psm 11",
            output_type=Output.DICT,
        )
        for index, raw_text in enumerate(data.get("text", [])):
            text = str(raw_text).strip()
            try:
                confidence = float(data["conf"][index])
            except Exception:
                confidence = -1.0
            if confidence < min_confidence or not meaningful_token(text):
                continue
            key = re.sub(r"\s+", "", text).lower()
            previous = findings.get(key)
            row = {
                "text": text,
                "confidence": round(confidence, 2),
                "background": background_name,
            }
            if previous is None or confidence > float(previous["confidence"]):
                findings[key] = row
    return sorted(findings.values(), key=lambda item: (-float(item["confidence"]), str(item["text"])))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, required=True)
    parser.add_argument("--report", type=Path)
    parser.add_argument("--allowlist", type=Path)
    parser.add_argument("--lang", default="chi_sim+eng")
    parser.add_argument("--min-confidence", type=float, default=60.0)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    root = args.root.resolve()
    report = (args.report or (root / "asset_text_residue_report.json")).resolve()
    allow = allowed_files(args.allowlist)
    manifest_names = manifest_non_text_files(root)
    assets = sorted(
        path for path in root.rglob("*.png")
        if any(part.startswith("split_png_elements") for part in path.parts)
        and (not manifest_names or path.name in manifest_names)
    )
    rows: list[dict[str, Any]] = []
    for asset in assets:
        if asset.name in allow:
            rows.append({
                "file_name": asset.name,
                "path": str(asset.relative_to(root)).replace("\\", "/"),
                "allowed_integral_text": True,
                "tokens": [],
            })
            continue
        tokens = detect_tokens(asset, args.lang, args.min_confidence)
        if not tokens:
            continue
        rows.append({
            "file_name": asset.name,
            "path": str(asset.relative_to(root)).replace("\\", "/"),
            "allowed_integral_text": asset.name in allow,
            "tokens": tokens,
        })
    violations = [row for row in rows if not row["allowed_integral_text"]]
    payload = {
        "status": "passed" if not violations else "failed_text_residue_detected",
        "root": ".",
        "root_note": "paths_relative_to_promoted_page_output",
        "language": args.lang,
        "min_confidence": args.min_confidence,
        "asset_count": len(assets),
        "allowed_files": sorted(allow),
        "detections": rows,
        "violations": violations,
    }
    report.parent.mkdir(parents=True, exist_ok=True)
    report.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    csv_path = report.with_suffix(".csv")
    with csv_path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=("file_name", "allowed_integral_text", "detected_text", "max_confidence"),
        )
        writer.writeheader()
        for row in rows:
            tokens = row["tokens"]
            writer.writerow({
                "file_name": row["file_name"],
                "allowed_integral_text": "yes" if row["allowed_integral_text"] else "no",
                "detected_text": " | ".join(str(item["text"]) for item in tokens),
                "max_confidence": max((float(item["confidence"]) for item in tokens), default=""),
            })
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0 if not violations else 2


if __name__ == "__main__":
    raise SystemExit(main())
