#!/usr/bin/env python3
"""Prepare, track, auto-finalize, and merge independent full-V1 page workers.

This script deliberately does not perform the refined rebuild. Codex page workers
must execute the original codeximage-to-editable-ppt-v1 workflow end to end in
isolated task directories. The coordinator only prepares page inputs, records
wall-clock timing, requires terminal PowerPoint validation inside each isolated
page task, and merges only passed one-page decks in source order. Working
artifacts are temporary by default and are removed after a successful merge.
"""

from __future__ import annotations

import argparse
import ctypes
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import time
from copy import deepcopy
from dataclasses import dataclass
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path
from typing import Any, Iterable

from PIL import Image, ImageChops, ImageStat
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


SKILL_VERSION = "1.2.7"
MAX_MULTI_ROOT_WORKERS = 6
SLIDE_SIZE_ROUNDING_TOLERANCE_EMU = 100
FULL_SLIDE_PICTURE_THRESHOLD = 0.90
RASTER_TEXT_OVERLAP_THRESHOLD = 0.50
RASTER_RESIDUE_MIN_PICTURE_COVERAGE = 0.10
FIGURE_TEXT_STATUS_CHOICES = ("panel_markers_extracted", "no_external_panel_markers")
CARD_DECOMPOSITION_STATUS_CHOICES = ("cards_decomposed", "no_structured_cards")
SCIENTIFIC_CHART_STATUS_CHOICES = ("charts_complete", "no_data_charts")
SCIENTIFIC_CHART_NAME_PREFIX = "SCI_CHART_COMPLETE::"
SCIENTIFIC_IMAGE_NAME_PREFIX = "SCI_IMAGE_COMPLETE::"
SCIENTIFIC_FRAGMENT_PREFIX = "SCI_CHART_FRAGMENT::"
SCIENTIFIC_PANEL_MARKER_PREFIX = "SCI_PANEL_MARKER::"
SCIENTIFIC_CHART_SOURCE_BOX_MARKER = "::SRCBOX="
SCIENTIFIC_CHART_EDGE_BAND_FRACTION = 0.015
SCIENTIFIC_CHART_MAX_EDGE_INK_RATIO = 0.20
SCIENTIFIC_CHART_SOURCE_MAE_LIMIT = 3.0
BRAND_MASTHEAD_STATUS_CHOICES = (
    "shared_faithful",
    "page_specific_faithful",
    "faithful",
    "not_applicable",
)
SHARED_MASTHEAD_CROP_FRACTION = 0.10
SHARED_MASTHEAD_SIGNATURE_BITS = 64 * 16
SHARED_MASTHEAD_SIMILARITY_THRESHOLD = 0.76
SHARED_MASTHEAD_PLACEMENT_TOLERANCE = 0.003
SUPPORTED_IMAGES = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"}
SUPPORTED_PRESENTATIONS = {".ppt", ".pptx"}
DEFAULT_V1_SKILL = Path(__file__).resolve().parents[2] / "codeximage-to-editable-ppt-v1"
DEFAULT_SKILL_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_EXTRACTOR = DEFAULT_SKILL_ROOT / "scripts" / "decompose_visual_elements.py"
DEFAULT_VALIDATOR = DEFAULT_SKILL_ROOT / "scripts" / "validate_layout_powerpoint.py"
DEFAULT_WORKER_CONTRACT = DEFAULT_SKILL_ROOT / "references" / "page_worker_contract.md"
V1_CONTRACT_SOURCE_HASHES = {
    "SKILL.md": "771d2184b7b442c032edb6841f302e447ef8ad5756182f491dc8ab8df66dacb4",
    "references/refined_rebuild_workflow.md": "eeaa19646fcaea8633de701c490fdc06ed8db9216a4a22a70fce975c823e20ab",
}


def now_iso() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="milliseconds")


def system_memory_bytes() -> tuple[int | None, int | None]:
    if os.name == "nt":
        class MemoryStatusEx(ctypes.Structure):
            _fields_ = [
                ("dwLength", ctypes.c_ulong),
                ("dwMemoryLoad", ctypes.c_ulong),
                ("ullTotalPhys", ctypes.c_ulonglong),
                ("ullAvailPhys", ctypes.c_ulonglong),
                ("ullTotalPageFile", ctypes.c_ulonglong),
                ("ullAvailPageFile", ctypes.c_ulonglong),
                ("ullTotalVirtual", ctypes.c_ulonglong),
                ("ullAvailVirtual", ctypes.c_ulonglong),
                ("ullAvailExtendedVirtual", ctypes.c_ulonglong),
            ]

        status = MemoryStatusEx()
        status.dwLength = ctypes.sizeof(MemoryStatusEx)
        if ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(status)):
            return int(status.ullTotalPhys), int(status.ullAvailPhys)
    return None, None


def active_powerpoint_process_count() -> int | None:
    if os.name != "nt":
        return None
    try:
        proc = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=5,
            check=False,
        )
    except (OSError, subprocess.SubprocessError):
        return None
    return sum(1 for line in proc.stdout.splitlines() if "POWERPNT.EXE" in line.upper())


def detect_system_capacity(
    cpu_threads_per_worker: int = 4,
    memory_gb_per_worker: float = 4.0,
    powerpoint_worker_cap: int = 4,
) -> dict[str, Any]:
    logical_cpus = max(1, int(os.cpu_count() or 1))
    total_memory, available_memory = system_memory_bytes()
    powerpoint_processes = active_powerpoint_process_count()
    cpu_capacity = max(1, logical_cpus // max(1, int(cpu_threads_per_worker)))
    memory_unit = max(0.5, float(memory_gb_per_worker)) * (1024**3)
    memory_capacity = (
        max(1, int(available_memory // memory_unit)) if available_memory is not None else powerpoint_worker_cap
    )
    occupied_powerpoint_slots = min(
        max(0, int(powerpoint_processes or 0)), max(0, powerpoint_worker_cap - 1)
    )
    powerpoint_capacity = max(1, int(powerpoint_worker_cap) - occupied_powerpoint_slots)
    hardware_safe_workers = max(1, min(cpu_capacity, memory_capacity, powerpoint_capacity))
    gib = float(1024**3)
    return {
        "assessment_kind": "conservative_safe_upper_bound_not_hardware_theoretical_maximum",
        "logical_cpu_count": logical_cpus,
        "total_memory_gb": round(total_memory / gib, 2) if total_memory is not None else None,
        "available_memory_gb": round(available_memory / gib, 2) if available_memory is not None else None,
        "active_powerpoint_processes": powerpoint_processes,
        "assumptions": {
            "logical_cpu_threads_per_worker": max(1, int(cpu_threads_per_worker)),
            "available_memory_gb_per_worker": max(0.5, float(memory_gb_per_worker)),
            "powerpoint_worker_cap": max(1, int(powerpoint_worker_cap)),
        },
        "capacity_by_resource": {
            "cpu": cpu_capacity,
            "available_memory": memory_capacity,
            "powerpoint": powerpoint_capacity,
        },
        "hardware_safe_worker_limit": hardware_safe_workers,
    }


def read_json(path: Path, default: Any = None) -> Any:
    if not path.exists():
        return default
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def append_task_event(task: dict[str, Any], stage: str, event: str, status: str) -> None:
    event_path = Path(task["task_dir"]) / "worker_events.jsonl"
    payload = {
        "timestamp": now_iso(),
        "monotonic_ns": time.monotonic_ns(),
        "pid": os.getpid(),
        "uid": task["uid"],
        "stage": stage,
        "event": event,
        "status": status,
    }
    with event_path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(payload, ensure_ascii=False) + "\n")


def safe_name(value: str) -> str:
    cleaned = "".join(c if c.isalnum() or c in "-_" else "_" for c in value).strip("_")
    return cleaned[:80] or "page"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def sha256_normalized_text_file(path: Path) -> str:
    data = path.read_bytes().replace(b"\r\n", b"\n").replace(b"\r", b"\n")
    return hashlib.sha256(data).hexdigest()


def verify_worker_contract_sources(v1_skill: Path, contract: Path) -> dict[str, str]:
    if not contract.exists():
        raise FileNotFoundError(f"Worker contract missing: {contract}")
    actual: dict[str, str] = {}
    mismatches: list[str] = []
    for relative, expected in V1_CONTRACT_SOURCE_HASHES.items():
        source = v1_skill / relative
        if not source.exists():
            raise FileNotFoundError(f"Canonical V1 source missing: {source}")
        value = sha256_normalized_text_file(source)
        actual[relative] = value
        if value.lower() != expected.lower():
            mismatches.append(relative)
    if mismatches:
        raise RuntimeError(
            "The canonical V1 files changed after the compact worker contract was authored: "
            + ", ".join(mismatches)
            + ". Update page_worker_contract.md and its hashes before dispatching workers."
        )
    return actual


def run_command(cmd: list[str], log_path: Path, timeout: int) -> int:
    started = now_iso()
    t0 = time.perf_counter()
    proc = subprocess.run(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
        env={**os.environ, "PYTHONUTF8": "1"},
    )
    log_path.parent.mkdir(parents=True, exist_ok=True)
    log_path.write_text(
        json.dumps(
            {
                "started_at": started,
                "finished_at": now_iso(),
                "duration_seconds": round(time.perf_counter() - t0, 3),
                "return_code": proc.returncode,
                "command": subprocess.list2cmdline(cmd),
                "output": proc.stdout,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    return proc.returncode


@dataclass
class SourcePage:
    source_index: int
    source_input: Path
    page_index: int
    page_id: str
    image_path: Path
    width: int
    height: int
    source_type: str


def collect_inputs(raw_inputs: Iterable[Path]) -> list[Path]:
    found: list[Path] = []
    for raw in raw_inputs:
        path = raw.expanduser().resolve()
        if path.is_dir():
            for candidate in sorted(path.iterdir()):
                if candidate.suffix.lower() in SUPPORTED_IMAGES | SUPPORTED_PRESENTATIONS:
                    found.append(candidate.resolve())
        elif path.is_file() and path.suffix.lower() in SUPPORTED_IMAGES | SUPPORTED_PRESENTATIONS:
            found.append(path)
        else:
            raise FileNotFoundError(f"Unsupported or missing input: {path}")
    if not found:
        raise RuntimeError("No supported inputs were found")
    return found


def intersect_area(
    left_a: int,
    top_a: int,
    width_a: int,
    height_a: int,
    left_b: int,
    top_b: int,
    width_b: int,
    height_b: int,
) -> int:
    right = min(left_a + width_a, left_b + width_b)
    bottom = min(top_a + height_a, top_b + height_b)
    left = max(left_a, left_b)
    top = max(top_a, top_b)
    return max(0, right - left) * max(0, bottom - top)


def shape_slide_coverage(shape: Any, slide_width: int, slide_height: int) -> float:
    slide_area = max(1, slide_width * slide_height)
    visible = intersect_area(
        int(shape.left),
        int(shape.top),
        int(shape.width),
        int(shape.height),
        0,
        0,
        slide_width,
        slide_height,
    )
    return float(visible) / float(slide_area)


def detect_masthead_crop_height(image: Image.Image) -> int:
    """Estimate a full-width masthead boundary without drifting into the page title."""
    width, height = image.size
    fallback = max(1, min(height, round(height * 0.10)))
    if height < 20 or width < 20:
        return fallback
    sample_width = min(192, width)
    sample = image.convert("RGB").resize((sample_width, height), Image.Resampling.BILINEAR)
    rows: list[tuple[float, float, float]] = []
    for y in range(height):
        values = list(sample.crop((0, y, sample_width, y + 1)).getdata())
        rows.append(
            tuple(sum(pixel[channel] for pixel in values) / sample_width for channel in range(3))
        )
    lower = max(3, round(height * 0.04))
    upper = min(height - 3, round(height * 0.18))
    best_y = fallback
    best_score = 0.0
    for y in range(lower, upper + 1):
        before = tuple(sum(rows[k][channel] for k in range(y - 3, y)) / 3.0 for channel in range(3))
        after = tuple(sum(rows[k][channel] for k in range(y, y + 3)) / 3.0 for channel in range(3))
        score = sum((before[channel] - after[channel]) ** 2 for channel in range(3)) ** 0.5
        if score > best_score:
            best_y = y
            best_score = score
    return best_y if best_score >= 25.0 else fallback


def masthead_probe(path: Path, input_index: int, page_index: int = 1) -> dict[str, Any] | None:
    """Return a compact perceptual signature for the top masthead band of a raster page."""
    if path.suffix.lower() not in SUPPORTED_IMAGES or not path.exists():
        return None
    with Image.open(path) as opened:
        image = opened.convert("RGB")
        width, height = image.size
        crop_height = detect_masthead_crop_height(image)
        band = image.crop((0, 0, width, crop_height))
        sample = band.resize((64, 16), Image.Resampling.LANCZOS)
        gray = sample.convert("L")
        pixels = list(gray.getdata())
        mean = sum(pixels) / max(1, len(pixels))
        signature = 0
        for value in pixels:
            signature = (signature << 1) | int(value >= mean)
        stats = ImageStat.Stat(sample.convert("HSV"))
        saturation_mean = float(stats.mean[1]) / 255.0
        luma_stddev = float(ImageStat.Stat(gray).stddev[0])
        mean_rgb = [float(value) for value in ImageStat.Stat(sample).mean]
    branded_candidate = saturation_mean >= 0.08 or luma_stddev >= 18.0
    return {
        "input_index": input_index,
        "page_index": page_index,
        "path": str(path),
        "width": width,
        "height": height,
        "crop_fraction": crop_height / float(height),
        "crop_height": crop_height,
        "signature_hex": f"{signature:0{SHARED_MASTHEAD_SIGNATURE_BITS // 4}x}",
        "mean_saturation": round(saturation_mean, 6),
        "mean_rgb": [round(value, 6) for value in mean_rgb],
        "luma_stddev": round(luma_stddev, 6),
        "branded_candidate": branded_candidate,
    }


def masthead_similarity(left: dict[str, Any], right: dict[str, Any]) -> float:
    if (left.get("width"), left.get("height")) != (right.get("width"), right.get("height")):
        return 0.0
    left_value = int(str(left["signature_hex"]), 16)
    right_value = int(str(right["signature_hex"]), 16)
    distance = (left_value ^ right_value).bit_count()
    spatial_similarity = 1.0 - (float(distance) / float(SHARED_MASTHEAD_SIGNATURE_BITS))
    left_rgb = [float(value) for value in left.get("mean_rgb") or (0.0, 0.0, 0.0)]
    right_rgb = [float(value) for value in right.get("mean_rgb") or (0.0, 0.0, 0.0)]
    color_distance = sum((a - b) ** 2 for a, b in zip(left_rgb, right_rgb)) ** 0.5
    color_similarity = max(0.0, 1.0 - color_distance / ((3 * (255**2)) ** 0.5))
    return 0.70 * spatial_similarity + 0.30 * color_similarity


def plan_shared_masthead_reuse(rows: list[dict[str, Any]]) -> dict[str, Any]:
    """Cluster visually matching raster mastheads and select one medoid source per group."""
    probes: list[dict[str, Any]] = []
    deferred: list[dict[str, Any]] = []
    for row in rows:
        input_index = int(row["input_index"])
        if row.get("input_kind") != "raster_slide_image":
            deferred.append(
                {
                    "input_index": input_index,
                    "reason": "masthead_grouping_deferred_until_pages_are_rendered",
                }
            )
            continue
        probe = masthead_probe(Path(row["path"]), input_index, 1)
        if probe:
            probes.append(probe)

    candidate_groups: list[list[dict[str, Any]]] = []
    unassigned: list[dict[str, Any]] = []
    for probe in probes:
        if not probe["branded_candidate"]:
            unassigned.append(
                {
                    "input_index": probe["input_index"],
                    "page_index": probe["page_index"],
                    "reason": "no_distinctive_top_masthead_detected",
                }
            )
            continue
        matching_groups = [
            group
            for group in candidate_groups
            if any(
                masthead_similarity(probe, member)
                >= SHARED_MASTHEAD_SIMILARITY_THRESHOLD
                for member in group
            )
        ]
        if not matching_groups:
            candidate_groups.append([probe])
        else:
            primary = matching_groups[0]
            primary.append(probe)
            for secondary in matching_groups[1:]:
                primary.extend(secondary)
                candidate_groups.remove(secondary)

    groups: list[dict[str, Any]] = []
    for raw_group in candidate_groups:
        if len(raw_group) < 2:
            probe = raw_group[0]
            unassigned.append(
                {
                    "input_index": probe["input_index"],
                    "page_index": probe["page_index"],
                    "reason": "no_matching_page_for_shared_masthead",
                }
            )
            continue
        medoid = max(
            raw_group,
            key=lambda candidate: sum(
                masthead_similarity(candidate, other) for other in raw_group
            ),
        )
        minimum_similarity = min(
            masthead_similarity(medoid, other) for other in raw_group
        )
        groups.append(
            {
                "group_id": f"masthead_group_{len(groups) + 1:03d}",
                "status": "reuse_recommended",
                "canonical_page": {
                    "input_index": medoid["input_index"],
                    "page_index": medoid["page_index"],
                },
                "members": [
                    {"input_index": item["input_index"], "page_index": item["page_index"]}
                    for item in raw_group
                ],
                "crop_fraction": float(medoid["crop_fraction"]),
                "normalized_placement": [0.0, 0.0, 1.0, float(medoid["crop_fraction"])],
                "minimum_similarity": round(minimum_similarity, 6),
            }
        )
    return {
        "status": "reuse_available" if groups else "no_shared_masthead_group_detected",
        "method": "top_band_spatial_and_brand_color_connected_grouping_with_medoid_canonical_page",
        "crop_fraction": SHARED_MASTHEAD_CROP_FRACTION,
        "similarity_threshold": SHARED_MASTHEAD_SIMILARITY_THRESHOLD,
        "groups": groups,
        "unassigned": unassigned,
        "deferred": deferred,
    }


def analyze_preflight_input(path: Path, input_index: int) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "input_index": input_index,
        "path": str(path),
        "sha256": sha256_file(path),
        "suffix": path.suffix.lower(),
        "pages": [],
    }
    if path.suffix.lower() in SUPPORTED_IMAGES:
        with Image.open(path) as image:
            width, height = image.size
        payload.update(
            {
                "input_kind": "raster_slide_image",
                "page_count": 1,
                "pages": [
                    {
                        "page_index": 1,
                        "classification": "full_page_raster_input",
                        "width": width,
                        "height": height,
                        "full_slide_picture_detected": True,
                        "maximum_picture_coverage": 1.0,
                        "source_use_policy": "reference_only_never_embed_as_full_slide_picture",
                        "text_residue_risk": "high_until_background_is_cleaned_or_regions_are_split",
                        "brand_masthead_strategy": "inspect_and_preserve_high_fidelity_if_present",
                        "figure_text_strategy": "extract_external_panel_markers_only_split_multi_panel_crops",
                        "scientific_chart_strategy": "preserve_one_complete_pixel_exact_source_crop_per_logical_chart_with_srcbox_title_or_editable_title_axes_ticks_values_units_legends_annotations_colorbars_and_safe_margins",
                        "structured_card_strategy": "decompose_card_chrome_text_icons_keep_only_irreducible_visual_raster",
                    }
                ],
            }
        )
        return payload

    if path.suffix.lower() == ".pptx":
        deck = Presentation(path)
        pages: list[dict[str, Any]] = []
        for page_index, slide in enumerate(deck.slides, start=1):
            picture_coverages = [
                shape_slide_coverage(shape, int(deck.slide_width), int(deck.slide_height))
                for shape in slide.shapes
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE
            ]
            maximum = max(picture_coverages, default=0.0)
            full_slide_count = sum(
                1 for value in picture_coverages if value >= FULL_SLIDE_PICTURE_THRESHOLD
            )
            non_picture_count = sum(
                1 for shape in slide.shapes if shape.shape_type != MSO_SHAPE_TYPE.PICTURE
            )
            if full_slide_count and non_picture_count:
                classification = "full_slide_raster_with_native_overlays"
            elif full_slide_count:
                classification = "full_slide_raster"
            elif picture_coverages and non_picture_count:
                classification = "mixed_raster_and_native_objects"
            elif picture_coverages:
                classification = "regional_raster_objects"
            else:
                classification = "native_or_vector_objects"
            pages.append(
                {
                    "page_index": page_index,
                    "classification": classification,
                    "shape_count": len(slide.shapes),
                    "picture_count": len(picture_coverages),
                    "full_slide_picture_count": full_slide_count,
                    "full_slide_picture_detected": bool(full_slide_count),
                    "maximum_picture_coverage": round(maximum, 6),
                    "source_use_policy": (
                        "extract_or_render_for_reference_then_rebuild"
                        if full_slide_count
                        else "preserve_native_structure_when_relevant"
                    ),
                    "text_residue_risk": "high" if full_slide_count else "normal",
                    "brand_masthead_strategy": "inspect_and_preserve_high_fidelity_if_present",
                    "figure_text_strategy": "extract_external_panel_markers_only_split_multi_panel_crops",
                    "scientific_chart_strategy": "preserve_one_complete_pixel_exact_source_crop_per_logical_chart_with_srcbox_title_or_editable_title_axes_ticks_values_units_legends_annotations_colorbars_and_safe_margins",
                    "structured_card_strategy": "decompose_card_chrome_text_icons_keep_only_irreducible_visual_raster",
                }
            )
        payload.update({"input_kind": "pptx", "page_count": len(pages), "pages": pages})
        return payload

    payload.update(
        {
            "input_kind": "legacy_ppt_requires_render_inspection",
            "page_count": None,
            "pages": [],
            "text_residue_risk": "unknown_until_rendered",
        }
    )
    return payload


def command_preflight(args: argparse.Namespace) -> int:
    inputs = collect_inputs(args.inputs)
    report_path = args.report.expanduser().resolve()
    intended_outdir = args.intended_outdir.expanduser().resolve() if args.intended_outdir else None
    rows = [analyze_preflight_input(path, index) for index, path in enumerate(inputs, start=1)]
    pages = [page for row in rows for page in row.get("pages", [])]
    full_slide_pages = [page for page in pages if page.get("full_slide_picture_detected")]
    shared_masthead_plan = plan_shared_masthead_reuse(rows)
    available_worker_slots = max(1, int(args.available_worker_slots))
    root_task_count = max(1, int(args.root_task_count))
    total_codex_worker_slots = available_worker_slots * root_task_count
    system_capacity = detect_system_capacity(
        cpu_threads_per_worker=int(args.cpu_threads_per_worker),
        memory_gb_per_worker=float(args.memory_gb_per_worker),
        powerpoint_worker_cap=int(args.powerpoint_worker_cap),
    )
    known_or_input_pages = len(pages) if pages else len(rows)
    recommended_workers = min(
        max(1, known_or_input_pages),
        total_codex_worker_slots,
        int(system_capacity["hardware_safe_worker_limit"]),
    )
    maximum_workers = min(
        max(1, known_or_input_pages), total_codex_worker_slots, MAX_MULTI_ROOT_WORKERS
    )
    payload = {
        "schema_version": 1,
        "skill_version": SKILL_VERSION,
        "created_at": now_iso(),
        "status": "awaiting_user_confirmation",
        "confirmation_required": True,
        "formal_output_untouched": True,
        "intended_outdir": str(intended_outdir) if intended_outdir else "",
        "inputs": rows,
        "summary": {
            "input_count": len(rows),
            "known_page_count": len(pages),
            "full_slide_raster_page_count": len(full_slide_pages),
            "legacy_inputs_requiring_render": sum(
                1 for row in rows if row.get("input_kind") == "legacy_ppt_requires_render_inspection"
            ),
            "shared_masthead_group_count": len(shared_masthead_plan["groups"]),
            "shared_masthead_reusable_page_count": sum(
                len(group["members"]) for group in shared_masthead_plan["groups"]
            ),
        },
        "shared_masthead_plan": shared_masthead_plan,
        "system_capacity": system_capacity,
        "parallelism_confirmation": {
            "selection_required": True,
            "root_task_count": root_task_count,
            "page_worker_slots_per_root_task": available_worker_slots,
            "available_codex_page_worker_slots": total_codex_worker_slots,
            "hardware_safe_worker_limit": system_capacity["hardware_safe_worker_limit"],
            "page_count_limit": max(1, known_or_input_pages),
            "maximum_allowed_workers": maximum_workers,
            "recommended_workers": recommended_workers,
            "allowed_workers": list(range(1, maximum_workers + 1)),
            "capacity_override_required_for": list(
                range(recommended_workers + 1, maximum_workers + 1)
            ),
            "multi_root_authorization_required": maximum_workers > available_worker_slots,
            "selected_workers": None,
        },
        "hard_rules_after_confirmation": [
            "never_embed_an_original_or_reencoded_full_slide_source_image",
            f"reject_any_picture_covering_at_least_{int(FULL_SLIDE_PICTURE_THRESHOLD * 100)}_percent_of_the_slide",
            "reject_detected_raster_text_residue_under_editable_text",
            "preserve_branded_mastheads_as_high_fidelity_components_or_mark_not_applicable",
            "reuse_each_confirmed_shared_masthead_group_as_one_immutable_asset_at_one_normalized_placement",
            "extract_external_panel_markers_or_mark_no_external_panel_markers",
            "require_exact_panel_marker_tags_and_text",
            "decompose_structured_information_cards_or_mark_no_structured_cards",
            "preserve_pixel_exact_source_grounded_data_charts_with_titles_axes_ticks_values_units_legends_annotations_colorbars_and_safe_margins",
            "reject_chart_padding_erasure_resampling_or_source_box_mismatch",
            "reject_shared_masthead_overlap_with_editable_page_text",
            "merge_only_terminally_and_raster_integrity_passed_pages",
            "preserve_audit_tags_and_run_fast_post_merge_content_invariants",
        ],
    }
    write_json(report_path, payload)
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0


def validate_preflight_confirmation(
    report_path: Path,
    inputs: list[Path],
    outdir: Path,
    confirmed: bool,
    workers: int,
    capacity_override_confirmed: bool,
) -> dict[str, Any]:
    if not confirmed:
        raise RuntimeError(
            "Preflight is complete but processing is not authorized. Ask the user to confirm, then rerun prepare with --confirmed."
        )
    report_path = report_path.expanduser().resolve()
    if not report_path.exists():
        raise FileNotFoundError(f"Preflight report missing: {report_path}")
    if report_path == outdir or outdir in report_path.parents:
        raise RuntimeError("Keep the preflight report outside the formal output directory")
    report = read_json(report_path, {})
    if report.get("status") != "awaiting_user_confirmation" or report.get("confirmation_required") is not True:
        raise RuntimeError("Invalid or already-consumed preflight report")
    parallelism = report.get("parallelism_confirmation") or {}
    if parallelism.get("selection_required") is not True:
        raise RuntimeError("Preflight report does not require an explicit page-worker selection")
    allowed_workers = [int(value) for value in parallelism.get("allowed_workers") or []]
    if workers not in allowed_workers:
        maximum = int(parallelism.get("maximum_allowed_workers") or 0)
        raise RuntimeError(
            f"Selected --workers {workers} was not confirmed by preflight; choose one of {allowed_workers} "
            f"(maximum {maximum})."
        )
    override_workers = [int(value) for value in parallelism.get("capacity_override_required_for") or []]
    if workers in override_workers and not capacity_override_confirmed:
        raise RuntimeError(
            "The selected worker count exceeds the conservative hardware recommendation. "
            "Ask the user to confirm the experimental capacity override, then rerun prepare "
            "with --capacity-override-confirmed."
        )
    intended = str(report.get("intended_outdir") or "").strip()
    if intended and os.path.normcase(str(Path(intended).resolve())) != os.path.normcase(str(outdir)):
        raise RuntimeError("Preflight intended output directory does not match prepare --outdir")
    reported_inputs = report.get("inputs") or []
    if len(reported_inputs) != len(inputs):
        raise RuntimeError("Preflight input count does not match prepare inputs")
    for expected_path, row in zip(inputs, reported_inputs):
        if os.path.normcase(str(expected_path)) != os.path.normcase(str(Path(row.get("path", "")).resolve())):
            raise RuntimeError("Preflight input order/path does not match prepare inputs")
        if sha256_file(expected_path) != str(row.get("sha256") or ""):
            raise RuntimeError(f"Input changed after preflight: {expected_path}")
    return report


def extract_source_pages(
    inputs: list[Path], staging: Path, python_exe: Path, extractor: Path, dpi: int, timeout: int
) -> list[SourcePage]:
    staging.mkdir(parents=True, exist_ok=True)
    pages: list[SourcePage] = []
    for source_index, source in enumerate(inputs, start=1):
        source_dir = staging / f"source_{source_index:03d}_{safe_name(source.stem)}"
        source_dir.mkdir(parents=True, exist_ok=True)
        if source.suffix.lower() in SUPPORTED_IMAGES:
            target = source_dir / f"page_001{source.suffix.lower()}"
            shutil.copy2(source, target)
            with Image.open(target) as image:
                width, height = image.size
            pages.append(SourcePage(source_index, source, 1, "image01", target, width, height, "input_image"))
            continue

        extraction = source_dir / "extraction"
        cmd = [
            str(python_exe),
            str(extractor),
            str(source),
            "--outdir",
            str(extraction),
            "--dpi",
            str(dpi),
            "--extract-pages-only",
            "--clean",
        ]
        rc = run_command(cmd, source_dir / "extract_pages.log.json", timeout)
        if rc != 0:
            raise RuntimeError(f"Page extraction failed for {source}")
        report = read_json(extraction / "image_source_report.json", [])
        if not report:
            raise RuntimeError(f"No pages extracted from {source}")
        for page_index, row in enumerate(report, start=1):
            original_page = Path(row["source_path"]).resolve()
            suffix = original_page.suffix.lower() or ".png"
            target = source_dir / f"page_{page_index:03d}{suffix}"
            shutil.copy2(original_page, target)
            pages.append(
                SourcePage(
                    source_index,
                    source,
                    page_index,
                    str(row.get("slide_id") or f"slide{page_index:02d}"),
                    target,
                    int(row.get("source_width") or 0),
                    int(row.get("source_height") or 0),
                    str(row.get("page_source_type") or "rendered_slide_image"),
                )
            )
    return pages


def materialize_shared_mastheads(
    preflight: dict[str, Any], pages: list[SourcePage], outdir: Path
) -> tuple[list[dict[str, Any]], dict[tuple[int, int], dict[str, Any]]]:
    """Create one immutable PNG per confirmed preflight masthead group."""
    page_lookup = {(page.source_index, page.page_index): page for page in pages}
    groups: list[dict[str, Any]] = []
    assignment: dict[tuple[int, int], dict[str, Any]] = {}
    for planned in (preflight.get("shared_masthead_plan") or {}).get("groups", []):
        canonical_ref = planned.get("canonical_page") or {}
        canonical_key = (
            int(canonical_ref.get("input_index") or 0),
            int(canonical_ref.get("page_index") or 0),
        )
        canonical_page = page_lookup.get(canonical_key)
        if canonical_page is None:
            continue
        fraction = float(planned.get("crop_fraction") or SHARED_MASTHEAD_CROP_FRACTION)
        with Image.open(canonical_page.image_path) as opened:
            image = opened.convert("RGBA")
            crop_height = max(1, min(image.height, round(image.height * fraction)))
            crop = image.crop((0, 0, image.width, crop_height))
        group_id = str(planned.get("group_id") or f"masthead_group_{len(groups) + 1:03d}")
        asset_dir = outdir / "shared_assets" / "mastheads"
        asset_dir.mkdir(parents=True, exist_ok=True)
        asset_path = asset_dir / f"{group_id}.png"
        crop.save(asset_path, format="PNG", optimize=True)
        normalized_placement = [0.0, 0.0, 1.0, crop_height / float(image.height)]
        group = {
            "group_id": group_id,
            "asset_path": str(asset_path.resolve()),
            "asset_sha256": sha256_file(asset_path),
            "canonical_source_index": canonical_page.source_index,
            "canonical_page_index": canonical_page.page_index,
            "crop_box_pixels": [0, 0, image.width, crop_height],
            "source_dimensions": [image.width, image.height],
            "normalized_placement": normalized_placement,
            "members": list(planned.get("members") or []),
            "immutable": True,
        }
        write_json(asset_dir / f"{group_id}.json", group)
        groups.append(group)
        for member in group["members"]:
            key = (int(member["input_index"]), int(member["page_index"]))
            if key in page_lookup:
                assignment[key] = group
    return groups, assignment


def worker_request(
    task: dict[str, Any], python_exe: Path, baseline_script: Path, validator: Path
) -> str:
    baseline_cmd = subprocess.list2cmdline(
        [
            str(python_exe),
            str(baseline_script),
            task["source_image"],
            "--outdir",
            task["baseline_dir"],
            "--dpi",
            "300",
            "--granularity",
            "fine",
            "--ocr",
            "--ocr-lang",
            "chi_sim+eng",
            "--ocr-confidence-threshold",
            "75",
            "--editable-text",
            "--review",
            "--quality-check",
            "--no-zip",
            "--clean",
        ]
    )
    terminal_validation_dir = Path(task["page_terminal_validation_report"]).parent
    terminal_validation_cmd = subprocess.list2cmdline(
        [
            str(python_exe),
            str(validator),
            task["refined_pptx"],
            "--source",
            task["source_image"],
            "--outdir",
            str(terminal_validation_dir),
            "--max-mae",
            "25",
            "--max-changed-ratio",
            "0.20",
            "--text-tolerance",
            "1.03",
        ]
    )
    shared_masthead = task.get("shared_masthead") or {}
    if shared_masthead:
        masthead_instruction = f"""- This page belongs to shared masthead group `{shared_masthead['group_id']}`.
- Insert the immutable shared asset `{shared_masthead['asset_path']}` unchanged. Its SHA-256 is `{shared_masthead['asset_sha256']}`.
- Place it at normalized slide coordinates `{shared_masthead['normalized_placement']}` with no additional crop, recolor, OCR, or reconstruction.
- Do not rebuild the masthead independently and do not extract the wordmark text from this shared brand crop."""
        masthead_step = "Use the supplied shared masthead asset exactly once at the supplied normalized placement. Skip independent masthead reconstruction and masthead-specific visual correction."
        brand_status = "shared_faithful"
        shared_flag = " --shared-masthead-used"
    else:
        masthead_instruction = "- No coordinator-approved shared masthead is assigned. Rebuild a page-specific branded masthead faithfully when present, or mark it not applicable."
        masthead_step = "Rebuild any page-specific branded masthead, institutional header, logo band, seal, wordmark, or distinctive header frame as a high-fidelity component."
        brand_status = "page_specific_faithful"
        shared_flag = ""
    return f"""# Compact full-V1 page worker request

Process exactly this one page. Read `{task['worker_contract']}` completely, then follow this request. The coordinator already read and verified the canonical V1 sources; do not reread either skill file.

## Isolation

- Task UID: `{task['uid']}`
- Source page: `{task['source_image']}`
- Write only inside: `{task['task_dir']}`
- Do not read another page worker's output.
- Do not read the full V1 or v1.2 skill files again.
- Do not generate ZIP archives or delivery bundles.
- Never place the source screenshot, or a re-encoded copy of it, as a full-slide picture.
- No picture may cover 90% or more of the slide. Build a clean background from native fills and tightly cropped independent decorations/figures.
- Preserve branded mastheads/institution headers as high-fidelity components; do not replace angled or logo-bearing header artwork with crude rectangles.
{masthead_instruction}
- For scientific figures, plots, heatmaps, tables, molecular diagrams, and paper-style multi-panel images, extract only cleanly separable external subfigure markers such as `(a)`, `(b)`, `(c)` as editable PowerPoint text. Split the major panels into separate tight image crops without those external markers when practical, then add only the marker back as editable text. Keep plot/image-body text such as `i`, `j`, `11 Å-NEP`, axis labels, legends, tick values, heatmap cell values, formulas, colorbar labels, and dense annotations inside the raster crop unless the user explicitly asks for full figure text extraction.
- A retained data chart must remain one complete raster per logical panel. Its crop must include the full plot, all axis titles, outermost tick labels and values, units, legends, annotations, and colorbar with at least 8 source pixels or 1% of panel width/height (whichever is larger) of clean safety margin on every side. Preserve the chart or subfigure title exactly once: preferably recreate a cleanly separable title as editable text at the source position; otherwise include it inside the complete raster crop. Never omit it. Never split one continuous chart into upper/lower or left/right fragments to satisfy another audit. Export the chart as an exact, unmodified rectangular crop from the source page: no transparent/white padding canvas, generative fill, erased labels, resampling, or secondary crop. Name it `SCI_CHART_COMPLETE::<stable_panel_id>::SRCBOX=<left>,<top>,<width>,<height>` using exact source-image pixel coordinates. Name a non-chart scientific illustration `SCI_IMAGE_COMPLETE::<stable_panel_id>`. Never use the forbidden `SCI_CHART_FRAGMENT::` prefix.
- Name every editable external panel marker `SCI_PANEL_MARKER::<stable_panel_id>::(<letter>)`, and make its actual text exactly equal the marker encoded in its name. A marker such as `(c)` may never become `(EUR)`, `(e)`, or another glyph during reconstruction.
- A structured information card is layout, not one image. For cards containing a title/header, central photo or complex illustration, icons, dividers, and explanatory rows, retain only the irreducible photo/illustration as raster. Rebuild the card frame, fills, separators, title, labels, and callouts as editable objects, and use one tight crop or native object per icon. A full-card crop is forbidden even when it occupies less than 90% of the slide.

## Mandatory V1 sequence

1. Record the start immediately:

```powershell
{python_exe} {Path(__file__).resolve()} mark-start --task-json {task['task_json']}
```

2. Run the original V1 baseline decomposition:

```powershell
{baseline_cmd}
```

3. Inspect the baseline JSON manifest, overlay, crops, and rough recomposed preview, then record that checkpoint:

```powershell
{python_exe} {Path(__file__).resolve()} mark-baseline --task-json {task['task_json']} --inspected
```

4. Perform the complete V1 refined editable rebuild using the source image only as ground truth. Do not use one full-slide screenshot as the reconstruction or hide source text with rectangles.
5. {masthead_step}
6. Extract only the external subfigure marker layer for scientific graphics. For each logical data-chart panel, create exactly one complete, pixel-exact source crop containing its title, entire plot, all axes/ticks/values/units, legends, annotations, and colorbar plus the required clean safety margin. Do not pad or alter that crop. Add only the external marker back as editable text. Do not split a continuous chart into strips or fragments. Name every chart picture `SCI_CHART_COMPLETE::<stable_panel_id>::SRCBOX=<left>,<top>,<width>,<height>`, every marker `SCI_PANEL_MARKER::<stable_panel_id>::(<letter>)`, and every retained non-chart scientific illustration `SCI_IMAGE_COMPLETE::<stable_panel_id>`.
7. Inspect structured information cards before final rendering. Split their title, explanatory rows, border/fills/dividers, and icons from the central photo or intricate illustration. Do not pass a comparison/benefit/risk/input-response-outcome card as one raster crop.
8. Use element-specific alpha masks for icons/logos/simple shapes and inspect every retained PNG for accidental text/background residue. Any raster underneath editable text must contain no old text, antialias fringe, or neighboring glyphs.
9. Render the refined page with Microsoft PowerPoint, visually compare it with the source, and correct severe overlap, clipping, crop errors, ugly wrapping, text overflow, and off-slide objects.
10. Run the terminal page validator after all corrections:

```powershell
{terminal_validation_cmd}
```

Open `{terminal_validation_dir / 'powerpoint_render' / 'slide_01.png'}` and compare it with the source. Repeat correction and terminal validation until the report passes with zero text overflow, zero off-slide shapes, no title/header collision, and no cropped chart title, axis title, outermost tick/value, unit, legend, annotation, or colorbar.
11. Produce only these acceptance paths:
   - refined PPTX: `{task['refined_pptx']}`
   - terminal validation report: `{task['page_terminal_validation_report']}`
12. Record completion only after opening and inspecting the passing terminal render. `record` automatically fails on a picture covering 90% or more of the slide, an embedded source screenshot, detected raster text underneath editable text, a high-confidence under-split portrait information card, or missing required statuses. Put the review evidence directly into the temporary worker result; replace `N`, all status values, and the correction summary with truthful values:

```powershell
{python_exe} {Path(__file__).resolve()} record --task-json {task['task_json']} --status completed --pptx {task['refined_pptx']} --terminal-validation-report {task['page_terminal_validation_report']} --iteration-count N --visual-inspection-confirmed --brand-masthead-status {brand_status}{shared_flag} --figure-text-status panel_markers_extracted --card-decomposition-status cards_decomposed --scientific-chart-status charts_complete --corrections-applied "summary or none"
```

Use `--brand-masthead-status shared_faithful --shared-masthead-used` for an assigned shared masthead, `page_specific_faithful` for a separately rebuilt masthead, and `not_applicable` only when the page has no branded masthead or institution header. Use `--figure-text-status no_external_panel_markers` only when there are no cleanly separable external `(a)/(b)/(c)`-style markers; otherwise every marker must use the required marker tag and exact text. Use `--card-decomposition-status no_structured_cards` only when the page has no structured information card. Use `--scientific-chart-status charts_complete` when at least one data chart is present and all such pictures carry complete-chart names with valid source boxes; otherwise use `no_data_charts`. If a chart or subfigure title is neither present in the complete crop nor recreated as editable text, or an axis, outer tick, value, unit, legend, annotation, colorbar, or required safety margin is missing, if the crop was padded/altered, or one chart was split into multiple fragments, record `--status failed`. Do not fail merely because complete plot-internal text remains rasterized.
"""


def shard_coordinator_request(shard: dict[str, Any]) -> str:
    requests = "\n".join(f"- `{path}`" for path in shard["worker_requests"])
    return f"""# Multi-root page shard coordinator request

Coordinate only shard `{shard['shard_id']}`. This is an auxiliary Codex task created with the user's explicit authorization for real multi-root parallelism.

## Assigned page worker requests

{requests}

## Rules

1. Do not rerun preflight or prepare and do not read the canonical V1 or V1.2 skill files.
2. Spawn one collaboration page worker per assigned request, back-to-back before waiting.
3. Give each page worker only its own request path. Each worker must follow the compact contract, complete PowerPoint correction and terminal validation, and write `worker_result.json`.
4. Do not inspect or edit another shard, merge pages, run the global watcher, or clean shared files.
5. Retry only a failed assigned page. Finish only after every assigned page records `completed`, or report the exact failed page.
"""


def command_prepare(args: argparse.Namespace) -> int:
    outdir = args.outdir.expanduser().resolve()
    inputs = collect_inputs(args.inputs)
    preflight_report_path = args.preflight_report.expanduser().resolve()
    preflight = validate_preflight_confirmation(
        preflight_report_path,
        inputs,
        outdir,
        bool(args.confirmed),
        int(args.workers),
        bool(args.capacity_override_confirmed),
    )
    if args.clean and outdir.exists():
        shutil.rmtree(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    contract_path = args.worker_contract.expanduser().resolve()
    contract_hashes = verify_worker_contract_sources(args.v1_skill.resolve(), contract_path)
    pages = extract_source_pages(
        inputs,
        outdir / "source_pages",
        args.python.resolve(),
        args.extractor.resolve(),
        args.dpi,
        args.timeout,
    )
    shared_masthead_groups, shared_masthead_assignment = materialize_shared_mastheads(
        preflight, pages, outdir
    )
    tasks: list[dict[str, Any]] = []
    for task_index, page in enumerate(pages, start=1):
        uid = f"page_{task_index:03d}_source_{page.source_index:03d}_p{page.page_index:03d}"
        task_dir = outdir / "page_tasks" / uid
        task_dir.mkdir(parents=True, exist_ok=True)
        source_copy = task_dir / f"source_page{page.image_path.suffix.lower()}"
        shutil.copy2(page.image_path, source_copy)
        contract_copy = task_dir / "page_worker_contract.md"
        shutil.copy2(contract_path, contract_copy)
        task = {
            "schema_version": 1,
            "skill_version": SKILL_VERSION,
            "uid": uid,
            "task_index": task_index,
            "source_index": page.source_index,
            "source_input": str(page.source_input),
            "page_index": page.page_index,
            "page_id": page.page_id,
            "source_type": page.source_type,
            "source_width": page.width,
            "source_height": page.height,
            "source_image": str(source_copy),
            "task_dir": str(task_dir),
            "task_json": str(task_dir / "task.json"),
            "baseline_dir": str(task_dir / "working" / "baseline_v1"),
            "refined_pptx": str(task_dir / f"{uid}_refined_editable.pptx"),
            "page_terminal_validation_report": str(
                task_dir / "working" / "terminal_validation" / "layout_quality_report.json"
            ),
            "worker_contract": str(contract_copy),
            "worker_request": str(task_dir / "worker_request.md"),
            "worker_result": str(task_dir / "worker_result.json"),
            "shared_masthead": shared_masthead_assignment.get(
                (page.source_index, page.page_index), {}
            ),
            "status": "pending",
        }
        write_json(Path(task["task_json"]), task)
        Path(task["worker_request"]).write_text(
            worker_request(task, args.python.resolve(), args.extractor.resolve(), args.validator.resolve()),
            encoding="utf-8",
        )
        tasks.append(task)

    root_task_count = int(preflight["parallelism_confirmation"].get("root_task_count") or 1)
    if int(args.workers) <= int(
        preflight["parallelism_confirmation"].get("page_worker_slots_per_root_task") or 1
    ):
        root_task_count = 1
    root_task_count = min(root_task_count, len(tasks), 2)
    chunk_size = max(1, (len(tasks) + root_task_count - 1) // root_task_count)
    shards: list[dict[str, Any]] = []
    for shard_index in range(root_task_count):
        assigned = tasks[shard_index * chunk_size : (shard_index + 1) * chunk_size]
        if not assigned:
            continue
        shard_id = f"shard_{shard_index + 1:03d}"
        shard_dir = outdir / "shards" / shard_id
        shard_dir.mkdir(parents=True, exist_ok=True)
        shard = {
            "schema_version": 1,
            "skill_version": SKILL_VERSION,
            "shard_id": shard_id,
            "role": "primary_root" if shard_index == 0 else "auxiliary_root",
            "task_uids": [task["uid"] for task in assigned],
            "worker_requests": [task["worker_request"] for task in assigned],
            "request_path": str(shard_dir / "shard_coordinator_request.md"),
        }
        Path(shard["request_path"]).write_text(shard_coordinator_request(shard), encoding="utf-8")
        write_json(shard_dir / "shard.json", shard)
        shards.append(shard)

    manifest = {
        "schema_version": 1,
        "skill_version": SKILL_VERSION,
        "created_at": now_iso(),
        "v1_skill": str(args.v1_skill.resolve()),
        "worker_contract": str(contract_path),
        "canonical_v1_hashes": contract_hashes,
        "artifact_mode": "temporary_working_artifacts",
        "preflight_report": str(preflight_report_path),
        "preflight_confirmed_at": now_iso(),
        "preflight_summary": preflight.get("summary", {}),
        "parallelism_confirmation": preflight.get("parallelism_confirmation", {}),
        "worker_limit": args.workers,
        "capacity_override_confirmed": bool(args.capacity_override_confirmed),
        "root_task_count": len(shards),
        "shared_masthead_groups": shared_masthead_groups,
        "shards": shards,
        "page_count": len(tasks),
        "tasks": tasks,
    }
    write_json(outdir / "parallel_v1_manifest.json", manifest)
    preflight["status"] = "confirmed_and_prepared"
    preflight["confirmed_at"] = now_iso()
    preflight["parallelism_confirmation"]["selected_workers"] = int(args.workers)
    preflight["parallelism_confirmation"]["capacity_override_confirmed"] = bool(
        args.capacity_override_confirmed
    )
    preflight["prepared_manifest"] = str(outdir / "parallel_v1_manifest.json")
    write_json(preflight_report_path, preflight)
    print(json.dumps({"status": "prepared", "page_count": len(tasks), "manifest": str(outdir / "parallel_v1_manifest.json")}, ensure_ascii=False, indent=2))
    return 0


def command_mark_start(args: argparse.Namespace) -> int:
    task_path = args.task_json.expanduser().resolve()
    task = read_json(task_path)
    result_path = Path(task["worker_result"])
    existing = read_json(result_path, {})
    if existing.get("started_at") and not args.restart:
        print(json.dumps(existing, ensure_ascii=False, indent=2))
        return 0
    result = {
        "uid": task["uid"],
        "status": "running",
        "started_at": now_iso(),
        "finished_at": "",
        "duration_seconds": None,
        "pptx": task["refined_pptx"],
        "notes": "",
    }
    write_json(result_path, result)
    append_task_event(task, "full_v1_page_workflow", "start", "running")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


def command_mark_baseline(args: argparse.Namespace) -> int:
    task_path = args.task_json.expanduser().resolve()
    task = read_json(task_path)
    result_path = Path(task["worker_result"])
    result = read_json(result_path, {})
    if result.get("status") != "running" or not result.get("started_at"):
        raise RuntimeError("Run mark-start before marking the baseline checkpoint")
    baseline = Path(task["baseline_dir"])
    required = (baseline / "visual_elements_manifest.json", baseline / "recomposed_from_elements.pptx")
    missing = [str(path) for path in required if not path.exists()]
    if missing:
        raise RuntimeError("Baseline checkpoint files are missing: " + ", ".join(missing))
    if not args.inspected:
        raise RuntimeError("The baseline checkpoint requires explicit --inspected confirmation")
    result["baseline_completed_at"] = now_iso()
    result["baseline_inspection_confirmed"] = True
    write_json(result_path, result)
    append_task_event(task, "baseline_inspection", "checkpoint", "passed")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


def parse_time(value: str) -> datetime:
    return datetime.fromisoformat(value)


def terminal_page_validation_issues(
    report_path: Path, expected_source: Path | None = None
) -> tuple[list[str], dict[str, Any]]:
    issues: list[str] = []
    if not report_path.exists():
        return ["terminal_validation_report_missing"], {}
    try:
        report = read_json(report_path, {})
    except Exception:
        return ["terminal_validation_report_invalid_json"], {}
    if report.get("status") != "passed":
        issues.append("terminal_validation_not_passed")
    if report.get("renderer") != "microsoft_powerpoint":
        issues.append("terminal_validation_not_powerpoint")
    if report.get("mode") != "final_visual":
        issues.append("terminal_validation_not_final_visual")
    slides = report.get("slides") or []
    if len(slides) != 1:
        issues.append(f"terminal_validation_expected_one_slide_found_{len(slides)}")
        return issues, report
    slide = slides[0]
    if slide.get("status") != "passed":
        issues.append("terminal_validation_slide_not_passed")
    if int(slide.get("text_overflow_count") or 0) != 0:
        issues.append("terminal_validation_text_overflow")
    if int(slide.get("off_slide_shape_count") or 0) != 0:
        issues.append("terminal_validation_off_slide_shape")
    rendered_path = Path(str(slide.get("rendered_path") or ""))
    if not rendered_path.exists():
        issues.append("terminal_validation_render_missing")
    if expected_source is not None:
        reported_source = Path(str(slide.get("source_path") or ""))
        try:
            source_matches = os.path.normcase(str(reported_source.resolve())) == os.path.normcase(
                str(expected_source.resolve())
            )
        except Exception:
            source_matches = False
        if not source_matches:
            issues.append("terminal_validation_source_mismatch")
    return issues, report


def normalized_text(value: str) -> str:
    return re.sub(r"[^0-9A-Za-z\u4e00-\u9fff]+", "", value).lower()


def meaningful_ocr_text(value: str) -> bool:
    compact = normalized_text(value)
    return bool(
        re.search(r"[\u4e00-\u9fff]{2,}", compact)
        or re.search(r"[a-z]{4,}", compact)
        or re.search(r"\d{2,}%?", compact)
    )


def detect_residual_tokens(image: Image.Image, min_confidence: float = 60.0) -> list[dict[str, Any]]:
    try:
        import pytesseract
        from pytesseract import Output
    except ImportError as exc:
        raise RuntimeError("pytesseract is required for suspicious raster/text overlap auditing") from exc
    available = set(pytesseract.get_languages(config=""))
    if {"chi_sim", "eng"}.issubset(available):
        language = "chi_sim+eng"
    elif "eng" in available:
        language = "eng"
    elif available:
        language = sorted(available)[0]
    else:
        raise RuntimeError("No Tesseract OCR language is available for raster residue auditing")
    data = pytesseract.image_to_data(
        image.convert("RGB"), lang=language, config="--psm 11", output_type=Output.DICT
    )
    tokens: list[dict[str, Any]] = []
    for index, raw in enumerate(data.get("text", [])):
        value = str(raw).strip()
        try:
            confidence = float(data["conf"][index])
        except Exception:
            confidence = -1.0
        if confidence >= min_confidence and meaningful_ocr_text(value):
            tokens.append(
                {
                    "text": value,
                    "confidence": round(confidence, 2),
                    "left": int(data.get("left", [0] * len(data.get("text", [])))[index]),
                    "top": int(data.get("top", [0] * len(data.get("text", [])))[index]),
                    "width": int(data.get("width", [0] * len(data.get("text", [])))[index]),
                    "height": int(data.get("height", [0] * len(data.get("text", [])))[index]),
                }
            )
    return tokens


def picture_overlap_crop(picture: Any, overlap: tuple[int, int, int, int]) -> Image.Image:
    image = Image.open(BytesIO(picture.image.blob)).convert("RGB")
    image_width, image_height = image.size
    crop_left = float(picture.crop_left or 0.0)
    crop_top = float(picture.crop_top or 0.0)
    crop_right = float(picture.crop_right or 0.0)
    crop_bottom = float(picture.crop_bottom or 0.0)
    visible_left = crop_left * image_width
    visible_top = crop_top * image_height
    visible_width = max(1.0, (1.0 - crop_left - crop_right) * image_width)
    visible_height = max(1.0, (1.0 - crop_top - crop_bottom) * image_height)
    overlap_left, overlap_top, overlap_width, overlap_height = overlap
    relative_left = (overlap_left - int(picture.left)) / max(1.0, float(picture.width))
    relative_top = (overlap_top - int(picture.top)) / max(1.0, float(picture.height))
    relative_right = (overlap_left + overlap_width - int(picture.left)) / max(
        1.0, float(picture.width)
    )
    relative_bottom = (overlap_top + overlap_height - int(picture.top)) / max(
        1.0, float(picture.height)
    )
    left = int(round(visible_left + relative_left * visible_width))
    top = int(round(visible_top + relative_top * visible_height))
    right = int(round(visible_left + relative_right * visible_width))
    bottom = int(round(visible_top + relative_bottom * visible_height))
    left = max(0, min(image_width - 1, left))
    top = max(0, min(image_height - 1, top))
    right = max(left + 1, min(image_width, right))
    bottom = max(top + 1, min(image_height, bottom))
    return image.crop((left, top, right, bottom))


def visible_picture_image(picture: Any) -> Image.Image:
    image = Image.open(BytesIO(picture.image.blob)).convert("RGB")
    image_width, image_height = image.size
    crop_left = float(picture.crop_left or 0.0)
    crop_top = float(picture.crop_top or 0.0)
    crop_right = float(picture.crop_right or 0.0)
    crop_bottom = float(picture.crop_bottom or 0.0)
    left = int(round(crop_left * image_width))
    top = int(round(crop_top * image_height))
    right = int(round((1.0 - crop_right) * image_width))
    bottom = int(round((1.0 - crop_bottom) * image_height))
    left = max(0, min(image_width - 1, left))
    top = max(0, min(image_height - 1, top))
    right = max(left + 1, min(image_width, right))
    bottom = max(top + 1, min(image_height, bottom))
    return image.crop((left, top, right, bottom))


def readable_figure_text_issues(pptx: Path) -> tuple[list[str], dict[str, Any]]:
    issues: list[str] = []
    report: dict[str, Any] = {
        "status": "passed",
        "audit_scope": "external_subfigure_markers_only",
        "hard_failure_policy": "do_not_fail_on_internal_plot_axis_legend_table_annotation_or_embedded_caption_text",
        "checks": [],
    }
    deck = Presentation(pptx)
    if len(deck.slides) != 1:
        return ["figure_text_audit_expected_one_slide"], report
    slide = deck.slides[0]
    slide_width = int(deck.slide_width)
    slide_height = int(deck.slide_height)
    for picture in [shape for shape in slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]:
        coverage = shape_slide_coverage(picture, slide_width, slide_height)
        top_ratio = float(int(picture.top)) / float(max(1, slide_height))
        report["checks"].append(
            {
                "picture_shape_id": int(picture.shape_id),
                "name": picture.name,
                "coverage_ratio": round(coverage, 6),
                "top_ratio": round(top_ratio, 6),
                "note": "internal figure text is allowed; only cleanly separable external subfigure markers such as (a)/(b)/(c) should be editable",
            }
        )
    issues = list(dict.fromkeys(issues))
    report["status"] = "passed" if not issues else "failed"
    report["issues"] = issues
    return issues, report


def _distinct_vertical_text_bands(tokens: list[dict[str, Any]], image_height: int) -> int:
    centers = sorted(
        (float(token.get("top", 0)) + float(token.get("height", 0)) / 2.0)
        / float(max(1, image_height))
        for token in tokens
    )
    bands: list[float] = []
    for center in centers:
        if not bands or abs(center - bands[-1]) > 0.055:
            bands.append(center)
        else:
            bands[-1] = (bands[-1] + center) / 2.0
    return len(bands)


def _edge_ink_ratios(image: Image.Image) -> dict[str, float]:
    image = image.convert("RGB")
    width, height = image.size
    corners = [
        image.getpixel((0, 0)),
        image.getpixel((max(0, width - 1), 0)),
        image.getpixel((0, max(0, height - 1))),
        image.getpixel((max(0, width - 1), max(0, height - 1))),
    ]
    background = tuple(sorted(pixel[channel] for pixel in corners)[len(corners) // 2] for channel in range(3))
    pixels = image.load()
    band = max(2, int(round(min(width, height) * SCIENTIFIC_CHART_EDGE_BAND_FRACTION)))

    def ink(x: int, y: int) -> bool:
        return max(abs(pixels[x, y][channel] - background[channel]) for channel in range(3)) > 28

    regions = {
        "top": ((x, y) for x in range(width) for y in range(min(band, height))),
        "bottom": ((x, y) for x in range(width) for y in range(max(0, height - band), height)),
        "left": ((x, y) for x in range(min(band, width)) for y in range(height)),
        "right": ((x, y) for x in range(max(0, width - band), width) for y in range(height)),
    }
    ratios: dict[str, float] = {}
    for side, coordinates in regions.items():
        values = list(coordinates)
        ratios[side] = sum(1 for x, y in values if ink(x, y)) / float(max(1, len(values)))
    return ratios


def _parse_scientific_chart_name(name: str) -> tuple[str, tuple[int, int, int, int] | None]:
    payload = name[len(SCIENTIFIC_CHART_NAME_PREFIX) :].strip()
    if SCIENTIFIC_CHART_SOURCE_BOX_MARKER not in payload:
        return payload, None
    panel_id, raw_box = payload.rsplit(SCIENTIFIC_CHART_SOURCE_BOX_MARKER, 1)
    values = [part.strip() for part in raw_box.split(",")]
    if len(values) != 4 or not all(re.fullmatch(r"\d+", value or "") for value in values):
        return panel_id.strip(), None
    return panel_id.strip(), tuple(int(value) for value in values)


def scientific_chart_completeness_issues(
    pptx: Path,
    source_image: Path | None = None,
    slide_index: int = 0,
    require_source_provenance: bool = False,
) -> tuple[list[str], dict[str, Any]]:
    """Check tagged charts against their exact source-page crop and safe margins."""
    issues: list[str] = []
    report: dict[str, Any] = {
        "status": "passed",
        "audit_scope": "worker_tagged_complete_data_charts",
        "required_name": (
            f"{SCIENTIFIC_CHART_NAME_PREFIX}<stable_panel_id>"
            f"{SCIENTIFIC_CHART_SOURCE_BOX_MARKER}<left>,<top>,<width>,<height>"
        ),
        "source_image": str(source_image) if source_image else "",
        "checks": [],
    }
    deck = Presentation(pptx)
    if slide_index < 0 or slide_index >= len(deck.slides):
        return [f"scientific_chart_slide_index_out_of_range:{slide_index}"], report
    slide = deck.slides[slide_index]
    source: Image.Image | None = None
    if source_image is not None:
        if not source_image.exists():
            issues.append("scientific_chart_source_image_missing")
        else:
            source = Image.open(source_image).convert("RGB")
    seen_ids: set[str] = set()
    for picture in [shape for shape in slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]:
        name = str(picture.name or "")
        if name.startswith(SCIENTIFIC_FRAGMENT_PREFIX):
            issues.append(f"scientific_chart_fragment_forbidden:picture_{picture.shape_id}")
            continue
        if not name.startswith(SCIENTIFIC_CHART_NAME_PREFIX):
            continue
        panel_id, source_box = _parse_scientific_chart_name(name)
        check: dict[str, Any] = {
            "picture_shape_id": int(picture.shape_id),
            "name": name,
            "panel_id": panel_id,
            "source_box": list(source_box) if source_box else None,
            "edge_ink_ratios": {},
        }
        if not panel_id:
            issues.append(f"scientific_chart_panel_id_missing:picture_{picture.shape_id}")
        elif panel_id in seen_ids:
            issues.append(f"scientific_chart_panel_fragmented_or_duplicated:{panel_id}")
        seen_ids.add(panel_id)
        if any(float(value or 0.0) > 0.0001 for value in (picture.crop_left, picture.crop_top, picture.crop_right, picture.crop_bottom)):
            issues.append(f"scientific_chart_powerpoint_crop_forbidden:picture_{picture.shape_id}")
        try:
            raw_image = Image.open(BytesIO(picture.image.blob))
            if "A" in raw_image.getbands():
                alpha_extrema = raw_image.getchannel("A").getextrema()
                check["alpha_extrema"] = list(alpha_extrema)
                if alpha_extrema[0] < 255:
                    issues.append(
                        f"scientific_chart_transparent_padding_or_erasure_forbidden:picture_{picture.shape_id}"
                    )
            image = visible_picture_image(picture)
            if require_source_provenance and source_box is None:
                issues.append(f"scientific_chart_source_box_missing_or_invalid:picture_{picture.shape_id}")
            if source is not None and source_box is not None:
                left, top, width, height = source_box
                source_width, source_height = source.size
                valid_box = (
                    width > 0
                    and height > 0
                    and left >= 0
                    and top >= 0
                    and left + width <= source_width
                    and top + height <= source_height
                )
                check["source_box_valid"] = valid_box
                if not valid_box:
                    issues.append(f"scientific_chart_source_box_out_of_bounds:picture_{picture.shape_id}")
                else:
                    expected = source.crop((left, top, left + width, top + height))
                    check["embedded_pixel_size"] = list(image.size)
                    check["expected_pixel_size"] = list(expected.size)
                    if image.size != expected.size:
                        issues.append(
                            f"scientific_chart_source_crop_dimensions_mismatch:picture_{picture.shape_id}"
                        )
                    else:
                        difference = ImageChops.difference(image.convert("RGB"), expected)
                        channel_means = ImageStat.Stat(difference).mean
                        source_mae = sum(channel_means) / float(max(1, len(channel_means)))
                        check["source_crop_mae"] = round(source_mae, 6)
                        if source_mae > SCIENTIFIC_CHART_SOURCE_MAE_LIMIT:
                            issues.append(
                                f"scientific_chart_not_exact_source_crop:picture_{picture.shape_id}:mae_{source_mae:.3f}"
                            )
            ratios = _edge_ink_ratios(image)
            check["edge_ink_ratios"] = {side: round(value, 6) for side, value in ratios.items()}
            failing = [side for side, value in ratios.items() if value > SCIENTIFIC_CHART_MAX_EDGE_INK_RATIO]
            check["unsafe_edges"] = failing
            if failing:
                issues.append(
                    f"scientific_chart_missing_safe_margin:picture_{picture.shape_id}:edges_{'_'.join(failing)}"
                )
        except Exception as exc:
            check["audit_error"] = str(exc)
            issues.append(f"scientific_chart_audit_unavailable:picture_{picture.shape_id}")
        report["checks"].append(check)
    issues = list(dict.fromkeys(issues))
    report["tagged_chart_count"] = len(report["checks"])
    report["status"] = "passed" if not issues else "failed"
    report["issues"] = issues
    return issues, report


def scientific_panel_marker_issues(
    pptx: Path, slide_index: int = 0
) -> tuple[list[str], dict[str, Any]]:
    issues: list[str] = []
    report: dict[str, Any] = {"status": "passed", "checks": []}
    deck = Presentation(pptx)
    if slide_index < 0 or slide_index >= len(deck.slides):
        return [f"scientific_panel_marker_slide_index_out_of_range:{slide_index}"], report
    for shape in deck.slides[slide_index].shapes:
        name = str(shape.name or "")
        if not name.startswith(SCIENTIFIC_PANEL_MARKER_PREFIX):
            continue
        match = re.fullmatch(r"SCI_PANEL_MARKER::([^:]+)::(\([A-Za-z]\))", name)
        actual = str(getattr(shape, "text", "") or "").strip()
        expected = match.group(2) if match else ""
        check = {
            "shape_id": int(shape.shape_id),
            "name": name,
            "actual_text": actual,
            "expected_text": expected,
        }
        if not match:
            issues.append(f"scientific_panel_marker_name_invalid:shape_{shape.shape_id}")
        elif actual != expected:
            issues.append(
                f"scientific_panel_marker_text_mismatch:shape_{shape.shape_id}:expected_{expected}:found_{actual}"
            )
        report["checks"].append(check)
    report["tagged_marker_count"] = len(report["checks"])
    report["status"] = "passed" if not issues else "failed"
    report["issues"] = issues
    return issues, report


def structured_card_decomposition_issues(pptx: Path) -> tuple[list[str], dict[str, Any]]:
    """Reject under-split information cards while sparing scientific plots."""
    issues: list[str] = []
    report: dict[str, Any] = {
        "status": "passed",
        "audit_scope": "portrait_information_card_title_plus_multiple_footer_rows",
        "checks": [],
    }
    deck = Presentation(pptx)
    if len(deck.slides) != 1:
        return ["structured_card_audit_expected_one_slide"], report
    slide = deck.slides[0]
    slide_width = int(deck.slide_width)
    slide_height = int(deck.slide_height)
    for picture in [shape for shape in slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]:
        picture_name = str(picture.name or "")
        if picture_name.startswith(SCIENTIFIC_CHART_NAME_PREFIX) or picture_name.startswith(
            SCIENTIFIC_IMAGE_NAME_PREFIX
        ):
            report["checks"].append(
                {
                    "picture_shape_id": int(picture.shape_id),
                    "name": picture_name,
                    "candidate": False,
                    "skipped_reason": "explicit_complete_scientific_visual_tag",
                }
            )
            continue
        coverage = shape_slide_coverage(picture, slide_width, slide_height)
        aspect_ratio = float(int(picture.width)) / float(max(1, int(picture.height)))
        check: dict[str, Any] = {
            "picture_shape_id": int(picture.shape_id),
            "name": picture.name,
            "coverage_ratio": round(coverage, 6),
            "aspect_ratio": round(aspect_ratio, 4),
            "candidate": False,
            "tokens": [],
        }
        # High-confidence geometry for a title + image + multi-row footer card.
        # Wide scientific plots and tiny logos are excluded intentionally.
        if not (0.03 <= coverage <= 0.35 and 0.45 <= aspect_ratio <= 1.25):
            report["checks"].append(check)
            continue
        try:
            image = visible_picture_image(picture)
            tokens = detect_residual_tokens(image)
            image_height = max(1, image.height)
            header_tokens = [
                token
                for token in tokens
                if (float(token.get("top", 0)) + float(token.get("height", 0)) / 2.0)
                / image_height
                <= 0.20
            ]
            footer_tokens = [
                token
                for token in tokens
                if (float(token.get("top", 0)) + float(token.get("height", 0)) / 2.0)
                / image_height
                >= 0.64
            ]
            footer_bands = _distinct_vertical_text_bands(footer_tokens, image_height)
            check.update(
                {
                    "tokens": tokens,
                    "header_token_count": len(header_tokens),
                    "footer_token_count": len(footer_tokens),
                    "footer_text_band_count": footer_bands,
                }
            )
            if header_tokens and len(footer_tokens) >= 2 and footer_bands >= 2:
                check["candidate"] = True
                issues.append(
                    f"structured_information_card_not_decomposed:picture_{picture.shape_id}"
                )
        except Exception as exc:
            check["audit_error"] = str(exc)
            issues.append(f"structured_card_audit_unavailable:picture_{picture.shape_id}")
        report["checks"].append(check)
    issues = list(dict.fromkeys(issues))
    report["status"] = "passed" if not issues else "failed"
    report["issues"] = issues
    return issues, report


def refined_raster_integrity_issues(
    pptx: Path, source_image: Path
) -> tuple[list[str], dict[str, Any]]:
    issues: list[str] = []
    report: dict[str, Any] = {
        "status": "passed",
        "full_slide_threshold": FULL_SLIDE_PICTURE_THRESHOLD,
        "pictures": [],
        "residue_checks": [],
    }
    deck = Presentation(pptx)
    if len(deck.slides) != 1:
        return ["raster_integrity_expected_one_slide"], report
    slide = deck.slides[0]
    source_hash = sha256_file(source_image)
    text_shapes = [
        shape
        for shape in slide.shapes
        if getattr(shape, "has_text_frame", False) and str(shape.text or "").strip()
    ]
    for picture in [shape for shape in slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]:
        coverage = shape_slide_coverage(picture, int(deck.slide_width), int(deck.slide_height))
        blob_hash = hashlib.sha256(picture.image.blob).hexdigest()
        details = {
            "shape_id": int(picture.shape_id),
            "name": picture.name,
            "coverage_ratio": round(coverage, 6),
            "source_hash_match": blob_hash == source_hash,
        }
        report["pictures"].append(details)
        structural_failure = False
        if coverage >= FULL_SLIDE_PICTURE_THRESHOLD:
            issues.append(f"full_slide_picture_forbidden:shape_{picture.shape_id}")
            structural_failure = True
        if blob_hash == source_hash:
            issues.append(f"source_image_embedded_as_picture:shape_{picture.shape_id}")
            structural_failure = True
        if structural_failure or coverage < RASTER_RESIDUE_MIN_PICTURE_COVERAGE:
            continue
        for text_shape in text_shapes:
            overlap_area = intersect_area(
                int(picture.left),
                int(picture.top),
                int(picture.width),
                int(picture.height),
                int(text_shape.left),
                int(text_shape.top),
                int(text_shape.width),
                int(text_shape.height),
            )
            text_area = max(1, int(text_shape.width) * int(text_shape.height))
            overlap_ratio = float(overlap_area) / float(text_area)
            if overlap_ratio < RASTER_TEXT_OVERLAP_THRESHOLD:
                continue
            overlap_left = max(int(picture.left), int(text_shape.left))
            overlap_top = max(int(picture.top), int(text_shape.top))
            overlap_right = min(
                int(picture.left) + int(picture.width),
                int(text_shape.left) + int(text_shape.width),
            )
            overlap_bottom = min(
                int(picture.top) + int(picture.height),
                int(text_shape.top) + int(text_shape.height),
            )
            check = {
                "picture_shape_id": int(picture.shape_id),
                "text_shape_id": int(text_shape.shape_id),
                "text": str(text_shape.text or "")[:160],
                "overlap_ratio": round(overlap_ratio, 6),
                "tokens": [],
            }
            try:
                crop = picture_overlap_crop(
                    picture,
                    (
                        overlap_left,
                        overlap_top,
                        overlap_right - overlap_left,
                        overlap_bottom - overlap_top,
                    ),
                )
                tokens = detect_residual_tokens(crop)
                check["tokens"] = tokens
                if tokens:
                    issues.append(
                        f"raster_text_residue_under_editable_text:picture_{picture.shape_id}:text_{text_shape.shape_id}"
                    )
            except Exception as exc:
                check["audit_error"] = str(exc)
                issues.append(
                    f"raster_text_residue_audit_unavailable:picture_{picture.shape_id}:text_{text_shape.shape_id}"
                )
            report["residue_checks"].append(check)
    issues = list(dict.fromkeys(issues))
    report["status"] = "passed" if not issues else "failed"
    report["issues"] = issues
    return issues, report


def shared_masthead_integrity_issues(
    pptx: Path, shared_masthead: dict[str, Any], slide_index: int = 0
) -> tuple[list[str], dict[str, Any]]:
    issues: list[str] = []
    asset_path = Path(str(shared_masthead.get("asset_path") or ""))
    expected_hash = str(shared_masthead.get("asset_sha256") or "")
    expected_box = [float(value) for value in shared_masthead.get("normalized_placement") or []]
    report: dict[str, Any] = {
        "group_id": shared_masthead.get("group_id", ""),
        "asset_path": str(asset_path),
        "expected_sha256": expected_hash,
        "expected_normalized_placement": expected_box,
        "matches": [],
    }
    if not asset_path.exists() or not expected_hash:
        issues.append("shared_masthead_asset_missing_or_unhashed")
    elif sha256_file(asset_path) != expected_hash:
        issues.append("shared_masthead_asset_changed_after_prepare")
    if len(expected_box) != 4:
        issues.append("shared_masthead_expected_placement_invalid")
    if not pptx.exists() or issues:
        report["status"] = "failed"
        report["issues"] = issues
        return issues, report

    deck = Presentation(pptx)
    if slide_index < 0 or slide_index >= len(deck.slides):
        issues.append(f"shared_masthead_slide_index_out_of_range:{slide_index}")
    else:
        slide = deck.slides[slide_index]
        slide_width = max(1, int(deck.slide_width))
        slide_height = max(1, int(deck.slide_height))
        for shape in slide.shapes:
            if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
                continue
            try:
                blob_hash = hashlib.sha256(shape.image.blob).hexdigest()
            except Exception:
                continue
            if blob_hash != expected_hash:
                continue
            actual_box = [
                float(shape.left) / slide_width,
                float(shape.top) / slide_height,
                float(shape.width) / slide_width,
                float(shape.height) / slide_height,
            ]
            crop_values = [
                float(getattr(shape, name, 0.0) or 0.0)
                for name in ("crop_left", "crop_top", "crop_right", "crop_bottom")
            ]
            placement_passed = all(
                abs(actual - expected) <= SHARED_MASTHEAD_PLACEMENT_TOLERANCE
                for actual, expected in zip(actual_box, expected_box)
            )
            crop_passed = all(abs(value) <= 1e-6 for value in crop_values)
            report["matches"].append(
                {
                    "shape_id": shape.shape_id,
                    "actual_normalized_placement": actual_box,
                    "crop_values": crop_values,
                    "placement_passed": placement_passed,
                    "crop_passed": crop_passed,
                }
            )
        if not report["matches"]:
            issues.append("shared_masthead_asset_not_embedded")
        elif not any(
            match["placement_passed"] and match["crop_passed"]
            for match in report["matches"]
        ):
            issues.append("shared_masthead_placement_or_crop_mismatch")
        valid_matches = [
            match
            for match in report["matches"]
            if match["placement_passed"] and match["crop_passed"]
        ]
        if valid_matches:
            masthead_bottom = int(
                max(expected_box[1] + expected_box[3], 0.0) * slide_height
            )
            collision_checks: list[dict[str, Any]] = []
            for shape in slide.shapes:
                if not getattr(shape, "has_text_frame", False):
                    continue
                text_value = str(getattr(shape, "text", "") or "").strip()
                if not text_value:
                    continue
                shape_top = int(shape.top)
                shape_bottom = shape_top + int(shape.height)
                intersects = shape_top < masthead_bottom and shape_bottom > 0
                collision_checks.append(
                    {
                        "shape_id": int(shape.shape_id),
                        "name": str(shape.name or ""),
                        "text": text_value[:120],
                        "shape_top": shape_top,
                        "shape_bottom": shape_bottom,
                        "masthead_bottom": masthead_bottom,
                        "intersects": intersects,
                    }
                )
                if intersects:
                    issues.append(
                        f"shared_masthead_overlaps_editable_text:shape_{shape.shape_id}"
                    )
            report["text_clearance_checks"] = collision_checks
    report["status"] = "passed" if not issues else "failed"
    report["issues"] = issues
    return issues, report


def command_record(args: argparse.Namespace) -> int:
    task_path = args.task_json.expanduser().resolve()
    task = read_json(task_path)
    result_path = Path(task["worker_result"])
    result = read_json(result_path, {})
    started_at = result.get("started_at") or args.started_at
    if not started_at:
        raise RuntimeError("Worker start was not recorded; run mark-start first or pass --started-at")
    finished_at = args.finished_at or now_iso()
    pptx = (args.pptx or Path(task["refined_pptx"])).expanduser().resolve()
    terminal_validation_report = (
        args.terminal_validation_report
        or Path(task.get("page_terminal_validation_report") or "")
    ).expanduser().resolve()
    issues: list[str] = []
    terminal_report: dict[str, Any] = {}
    raster_integrity: dict[str, Any] = {}
    figure_text_audit: dict[str, Any] = {}
    structured_card_audit: dict[str, Any] = {}
    scientific_chart_audit: dict[str, Any] = {}
    scientific_panel_marker_audit: dict[str, Any] = {}
    shared_masthead_audit: dict[str, Any] = {}
    if args.status == "completed":
        brand_masthead_status = str(getattr(args, "brand_masthead_status", "") or "")
        figure_text_status = str(getattr(args, "figure_text_status", "") or "")
        card_decomposition_status = str(
            getattr(args, "card_decomposition_status", "") or ""
        )
        scientific_chart_status = str(getattr(args, "scientific_chart_status", "") or "")
        shared_masthead = task.get("shared_masthead") or {}
        shared_masthead_used = bool(getattr(args, "shared_masthead_used", False))
        if brand_masthead_status not in set(BRAND_MASTHEAD_STATUS_CHOICES):
            issues.append("brand_masthead_status_missing_or_invalid")
        if shared_masthead:
            if brand_masthead_status != "shared_faithful":
                issues.append("shared_masthead_requires_shared_faithful_status")
            if not shared_masthead_used:
                issues.append("shared_masthead_usage_not_confirmed")
        elif brand_masthead_status == "shared_faithful" or shared_masthead_used:
            issues.append("shared_masthead_claimed_without_assignment")
        if figure_text_status not in set(FIGURE_TEXT_STATUS_CHOICES):
            issues.append("figure_text_status_missing_or_invalid")
        if card_decomposition_status not in set(CARD_DECOMPOSITION_STATUS_CHOICES):
            issues.append("card_decomposition_status_missing_or_invalid")
        if scientific_chart_status not in set(SCIENTIFIC_CHART_STATUS_CHOICES):
            issues.append("scientific_chart_status_missing_or_invalid")
        if not pptx.exists():
            issues.append("refined_pptx_missing")
        else:
            deck = Presentation(pptx)
            if len(deck.slides) != 1:
                issues.append(f"expected_one_slide_found_{len(deck.slides)}")
            else:
                raster_issues, raster_integrity = refined_raster_integrity_issues(
                    pptx, Path(task["source_image"])
                )
                issues.extend(raster_issues)
                if figure_text_status in set(FIGURE_TEXT_STATUS_CHOICES):
                    figure_issues, figure_text_audit = readable_figure_text_issues(pptx)
                    issues.extend(figure_issues)
                if card_decomposition_status in set(CARD_DECOMPOSITION_STATUS_CHOICES):
                    card_issues, structured_card_audit = structured_card_decomposition_issues(
                        pptx
                    )
                    issues.extend(card_issues)
                if scientific_chart_status in set(SCIENTIFIC_CHART_STATUS_CHOICES):
                    chart_issues, scientific_chart_audit = scientific_chart_completeness_issues(
                        pptx,
                        source_image=Path(task["source_image"]),
                        require_source_provenance=scientific_chart_status == "charts_complete",
                    )
                    issues.extend(chart_issues)
                    tagged_count = int(scientific_chart_audit.get("tagged_chart_count") or 0)
                    if scientific_chart_status == "charts_complete" and tagged_count < 1:
                        issues.append("scientific_chart_status_requires_tagged_complete_chart")
                    if scientific_chart_status == "no_data_charts" and tagged_count:
                        issues.append("tagged_scientific_chart_conflicts_with_no_data_charts_status")
                marker_issues, scientific_panel_marker_audit = scientific_panel_marker_issues(
                    pptx
                )
                issues.extend(marker_issues)
                tagged_marker_count = int(
                    scientific_panel_marker_audit.get("tagged_marker_count") or 0
                )
                if figure_text_status == "panel_markers_extracted" and tagged_marker_count < 1:
                    issues.append("figure_text_status_requires_tagged_panel_marker")
                if figure_text_status == "no_external_panel_markers" and tagged_marker_count:
                    issues.append("tagged_panel_marker_conflicts_with_no_external_markers_status")
                if shared_masthead:
                    masthead_issues, shared_masthead_audit = shared_masthead_integrity_issues(
                        pptx, shared_masthead
                    )
                    issues.extend(masthead_issues)
        if result.get("baseline_inspection_confirmed") is not True:
            issues.append("baseline_inspection_not_confirmed")
        if int(args.iteration_count or 0) < 1:
            issues.append("review_iteration_missing")
        if not args.visual_inspection_confirmed:
            issues.append("terminal_render_visual_inspection_not_confirmed")
        terminal_issues, terminal_report = terminal_page_validation_issues(
            terminal_validation_report,
            expected_source=Path(task["source_image"]),
        )
        issues.extend(terminal_issues)
    status = args.status if not issues else "failed"
    slide_metrics = {}
    if terminal_report.get("slides"):
        slide = terminal_report["slides"][0]
        slide_metrics = {
            "mean_absolute_error": slide.get("mean_absolute_error"),
            "changed_pixel_ratio": slide.get("changed_pixel_ratio"),
            "text_overflow_count": slide.get("text_overflow_count"),
            "off_slide_shape_count": slide.get("off_slide_shape_count"),
        }
    payload = {
        "uid": task["uid"],
        "status": status,
        "started_at": started_at,
        "baseline_completed_at": result.get("baseline_completed_at", ""),
        "baseline_inspection_confirmed": result.get("baseline_inspection_confirmed") is True,
        "finished_at": finished_at,
        "duration_seconds": round((parse_time(finished_at) - parse_time(started_at)).total_seconds(), 3),
        "pptx": str(pptx),
        "pptx_sha256": sha256_file(pptx) if pptx.exists() else "",
        "terminal_validation_report": str(terminal_validation_report),
        "terminal_validation_status": terminal_report.get("status", "") if terminal_report else "",
        "terminal_metrics": slide_metrics,
        "raster_integrity": raster_integrity,
        "raster_integrity_status": raster_integrity.get("status", "") if raster_integrity else "",
        "brand_masthead_status": str(getattr(args, "brand_masthead_status", "") or ""),
        "shared_masthead_used": bool(getattr(args, "shared_masthead_used", False)),
        "shared_masthead_group": str((task.get("shared_masthead") or {}).get("group_id") or ""),
        "shared_masthead_audit": shared_masthead_audit,
        "shared_masthead_audit_status": shared_masthead_audit.get("status", "") if shared_masthead_audit else "",
        "figure_text_status": str(getattr(args, "figure_text_status", "") or ""),
        "figure_text_audit": figure_text_audit,
        "figure_text_audit_status": figure_text_audit.get("status", "") if figure_text_audit else "",
        "card_decomposition_status": str(
            getattr(args, "card_decomposition_status", "") or ""
        ),
        "structured_card_audit": structured_card_audit,
        "structured_card_audit_status": structured_card_audit.get("status", "") if structured_card_audit else "",
        "scientific_chart_status": str(getattr(args, "scientific_chart_status", "") or ""),
        "scientific_chart_audit": scientific_chart_audit,
        "scientific_chart_audit_status": scientific_chart_audit.get("status", "") if scientific_chart_audit else "",
        "scientific_panel_marker_audit": scientific_panel_marker_audit,
        "scientific_panel_marker_audit_status": scientific_panel_marker_audit.get("status", "") if scientific_panel_marker_audit else "",
        "iteration_count": int(args.iteration_count or 0),
        "visual_inspection_confirmed": bool(args.visual_inspection_confirmed),
        "corrections_applied": list(args.corrections_applied or []),
        "notes": args.notes or "",
        "issues": issues,
    }
    write_json(result_path, payload)
    append_task_event(task, "full_v1_page_workflow", "end", status)
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0 if status == "completed" else 2


def merge_refined_pages_openxml(page_files: list[Path], output: Path) -> None:
    if not page_files:
        raise RuntimeError("No page decks to merge")
    output.parent.mkdir(parents=True, exist_ok=True)
    if len(page_files) == 1:
        shutil.copy2(page_files[0], output)
        return
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    page_sizes: list[tuple[int, int]] = []
    for page_file in page_files:
        deck = Presentation(page_file)
        if len(deck.slides) != 1:
            raise RuntimeError(f"Expected one slide: {page_file}")
        page_sizes.append((int(deck.slide_width), int(deck.slide_height)))
    # Full-V1 workers may differ by a few EMUs because of independent integer
    # rounding. Select the modal canvas, accept only tiny rounding differences,
    # and preserve every passed page object's absolute geometry without scaling.
    target_width, target_height = max(
        set(page_sizes), key=lambda size: (page_sizes.count(size), -page_sizes.index(size))
    )
    merged = Presentation()
    merged.slide_width = target_width
    merged.slide_height = target_height
    blank = merged.slide_layouts[6]
    for page_file in page_files:
        source_deck = Presentation(page_file)
        width_delta = abs(target_width - int(source_deck.slide_width))
        height_delta = abs(target_height - int(source_deck.slide_height))
        if (
            width_delta > SLIDE_SIZE_ROUNDING_TOLERANCE_EMU
            or height_delta > SLIDE_SIZE_ROUNDING_TOLERANCE_EMU
        ):
            raise RuntimeError(f"Material slide-size mismatch: {page_file}")
        source_slide = source_deck.slides[0]
        target_slide = merged.slides.add_slide(blank)
        source_bg = source_slide._element.cSld.bg
        if source_bg is not None:
            target_bg = target_slide._element.cSld.bg
            if target_bg is not None:
                target_slide._element.cSld.remove(target_bg)
            target_slide._element.cSld.insert(0, deepcopy(source_bg))
        for shape in source_slide.shapes:
            left = int(shape.left)
            top = int(shape.top)
            width = int(shape.width)
            height = int(shape.height)
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                picture = target_slide.shapes.add_picture(
                    BytesIO(shape.image.blob), left, top, width, height
                )
                picture.crop_left = shape.crop_left
                picture.crop_right = shape.crop_right
                picture.crop_top = shape.crop_top
                picture.crop_bottom = shape.crop_bottom
                picture.rotation = shape.rotation
                picture.name = shape.name
            else:
                cloned = deepcopy(shape.element)
                target_slide.shapes._spTree.insert_element_before(cloned, "p:extLst")
                copied = target_slide.shapes[-1]
                copied.left = left
                copied.top = top
                copied.width = width
                copied.height = height
    merged.save(output)


def concurrency_metrics(results: list[dict[str, Any]]) -> dict[str, Any]:
    intervals = [(parse_time(r["started_at"]), parse_time(r["finished_at"])) for r in results]
    events: list[tuple[datetime, int]] = []
    for start, finish in intervals:
        events.append((start, 1))
        events.append((finish, -1))
    active = 0
    peak = 0
    for _, delta in sorted(events, key=lambda item: (item[0], item[1])):
        active += delta
        peak = max(peak, active)
    first = min(start for start, _ in intervals)
    last = max(finish for _, finish in intervals)
    wall = (last - first).total_seconds()
    sequential = sum(float(r["duration_seconds"]) for r in results)
    return {
        "worker_first_start": first.isoformat(timespec="milliseconds"),
        "worker_last_finish": last.isoformat(timespec="milliseconds"),
        "worker_wall_seconds": round(wall, 3),
        "sequential_worker_seconds": round(sequential, 3),
        "measured_speedup": round(sequential / wall, 3) if wall > 0 else None,
        "peak_concurrent_workers": peak,
    }


def timing_acceptance(
    metrics: dict[str, Any], page_count: int, worker_limit: int, min_speedup: float
) -> dict[str, Any]:
    parallel_expected = page_count >= 2 and worker_limit >= 2
    required_peak = min(page_count, worker_limit) if parallel_expected else 1
    overlap_passed = int(metrics["peak_concurrent_workers"]) >= required_peak
    speedup = metrics.get("measured_speedup")
    speedup_passed = (not parallel_expected) or (speedup is not None and float(speedup) >= min_speedup)
    return {
        "parallel_expected": parallel_expected,
        "worker_limit": worker_limit,
        "required_peak_concurrent_workers": required_peak,
        "observed_peak_concurrent_workers": int(metrics["peak_concurrent_workers"]),
        "minimum_speedup": min_speedup if parallel_expected else None,
        "overlap_passed": overlap_passed,
        "speedup_passed": speedup_passed,
        "status": "passed" if overlap_passed and speedup_passed else "failed",
    }


def verify_merge_integrity(output: Path, expected_page_count: int) -> dict[str, Any]:
    issues: list[str] = []
    actual_page_count = 0
    if not output.exists() or output.stat().st_size <= 0:
        issues.append("merged_pptx_missing_or_empty")
    else:
        try:
            merged = Presentation(output)
            actual_page_count = len(merged.slides)
        except Exception as exc:
            issues.append(f"merged_pptx_cannot_open:{exc}")
    if actual_page_count != expected_page_count:
        issues.append(
            f"merged_page_count_mismatch_expected_{expected_page_count}_found_{actual_page_count}"
        )
    return {
        "status": "passed" if not issues else "failed",
        "pptx_saved": output.exists() and output.stat().st_size > 0,
        "expected_page_count": expected_page_count,
        "actual_page_count": actual_page_count,
        "order_source": "deterministic_manifest_iteration",
        "issues": issues,
    }


def post_merge_content_validation(
    output: Path, tasks: list[dict[str, Any]], results: list[dict[str, Any]]
) -> dict[str, Any]:
    """Fast invariant audit after merge; deliberately avoids a second render loop."""
    issues: list[str] = []
    pages: list[dict[str, Any]] = []
    if not output.exists():
        return {"status": "failed", "issues": ["merged_pptx_missing"], "pages": []}
    deck = Presentation(output)
    if len(deck.slides) != len(tasks):
        return {
            "status": "failed",
            "issues": ["merged_page_count_prevents_content_validation"],
            "pages": [],
        }
    for slide_index, (task, result) in enumerate(zip(tasks, results)):
        page_issues: list[str] = []
        chart_issues, chart_audit = scientific_chart_completeness_issues(
            output,
            source_image=Path(task["source_image"]),
            slide_index=slide_index,
            require_source_provenance=result.get("scientific_chart_status") == "charts_complete",
        )
        page_issues.extend(chart_issues)
        actual_chart_count = int(chart_audit.get("tagged_chart_count") or 0)
        expected_chart_count = int(
            (result.get("scientific_chart_audit") or {}).get("tagged_chart_count") or 0
        )
        if actual_chart_count != expected_chart_count:
            page_issues.append(
                f"scientific_chart_tag_count_changed_during_merge:expected_{expected_chart_count}:found_{actual_chart_count}"
            )
        marker_issues, marker_audit = scientific_panel_marker_issues(
            output, slide_index=slide_index
        )
        page_issues.extend(marker_issues)
        actual_marker_count = int(marker_audit.get("tagged_marker_count") or 0)
        expected_marker_count = int(
            (result.get("scientific_panel_marker_audit") or {}).get("tagged_marker_count")
            or 0
        )
        if actual_marker_count != expected_marker_count:
            page_issues.append(
                f"scientific_panel_marker_count_changed_during_merge:expected_{expected_marker_count}:found_{actual_marker_count}"
            )
        masthead_audit: dict[str, Any] = {}
        shared_masthead = task.get("shared_masthead") or {}
        if shared_masthead:
            masthead_issues, masthead_audit = shared_masthead_integrity_issues(
                output, shared_masthead, slide_index=slide_index
            )
            page_issues.extend(masthead_issues)
        prefixed = [f"page_{slide_index + 1}:{issue}" for issue in page_issues]
        issues.extend(prefixed)
        pages.append(
            {
                "uid": task["uid"],
                "slide_index": slide_index,
                "status": "passed" if not page_issues else "failed",
                "chart_audit": chart_audit,
                "panel_marker_audit": marker_audit,
                "shared_masthead_audit": masthead_audit,
                "issues": page_issues,
            }
        )
    return {
        "status": "passed" if not issues else "failed",
        "mode": "fast_post_merge_invariants_no_second_render",
        "checks": [
            "chart_shape_names_preserved",
            "exact_source_crop_provenance_preserved",
            "panel_marker_names_and_text_preserved",
            "shared_masthead_placement_and_title_clearance_preserved",
        ],
        "pages": pages,
        "issues": issues,
    }


def compact_worker_result(result: dict[str, Any]) -> dict[str, Any]:
    return {
        key: result.get(key)
        for key in (
            "uid",
            "status",
            "started_at",
            "baseline_completed_at",
            "finished_at",
            "duration_seconds",
            "iteration_count",
            "visual_inspection_confirmed",
            "brand_masthead_status",
            "shared_masthead_used",
            "shared_masthead_group",
            "shared_masthead_audit_status",
            "shared_masthead_audit",
            "figure_text_status",
            "figure_text_audit_status",
            "figure_text_audit",
            "card_decomposition_status",
            "structured_card_audit_status",
            "structured_card_audit",
            "scientific_chart_status",
            "scientific_chart_audit_status",
            "scientific_chart_audit",
            "scientific_panel_marker_audit_status",
            "scientific_panel_marker_audit",
            "corrections_applied",
            "terminal_validation_status",
            "terminal_metrics",
            "raster_integrity_status",
            "raster_integrity",
            "notes",
        )
    }


def cleanup_working_artifacts(outdir: Path) -> dict[str, Any]:
    removed: list[str] = []
    issues: list[str] = []
    for name in ("source_pages", "page_tasks", "shards", "shared_assets"):
        target = (outdir / name).resolve()
        if target.parent != outdir:
            issues.append(f"unsafe_cleanup_target:{target}")
            continue
        if target.exists():
            try:
                shutil.rmtree(target)
                removed.append(name)
            except Exception as exc:
                issues.append(f"cleanup_failed:{name}:{exc}")
    for name in ("parallel_v1_manifest.json", "parallel_v1_manifest.csv"):
        target = (outdir / name).resolve()
        if target.parent != outdir:
            issues.append(f"unsafe_cleanup_target:{target}")
            continue
        if target.exists():
            try:
                target.unlink()
                removed.append(name)
            except Exception as exc:
                issues.append(f"cleanup_failed:{name}:{exc}")
    return {
        "mode": "temporary_working_artifacts_removed",
        "status": "passed" if not issues else "warning",
        "removed": removed,
        "issues": issues,
    }


def command_finalize(args: argparse.Namespace) -> int:
    outdir = args.outdir.expanduser().resolve()
    manifest = read_json(outdir / "parallel_v1_manifest.json")
    tasks = manifest["tasks"]
    results: list[dict[str, Any]] = []
    page_files: list[Path] = []
    failures: list[dict[str, Any]] = []
    for task in tasks:
        result = read_json(Path(task["worker_result"]), {})
        if result.get("status") != "completed":
            failures.append({"uid": task["uid"], "result": result})
            continue
        page_file = Path(result["pptx"]).resolve()
        if not page_file.exists() or len(Presentation(page_file).slides) != 1:
            failures.append({"uid": task["uid"], "result": result, "issue": "invalid_page_pptx"})
            continue
        if result.get("raster_integrity_status") != "passed":
            failures.append(
                {"uid": task["uid"], "result": result, "issue": "page_raster_integrity_failed"}
            )
            continue
        if result.get("pptx_sha256") != sha256_file(page_file):
            failures.append(
                {"uid": task["uid"], "result": result, "issue": "page_changed_after_acceptance"}
            )
            continue
        terminal_validation_report = Path(
            result.get("terminal_validation_report")
            or task.get("page_terminal_validation_report")
            or ""
        ).resolve()
        terminal_issues, _ = terminal_page_validation_issues(
            terminal_validation_report,
            expected_source=Path(task["source_image"]),
        )
        if terminal_issues:
            failures.append(
                {
                    "uid": task["uid"],
                    "result": result,
                    "issue": "page_terminal_validation_failed",
                    "terminal_validation_issues": terminal_issues,
                }
            )
            continue
        results.append(result)
        page_files.append(page_file)
    if failures:
        write_json(outdir / "parallel_v1_finalization.json", {"status": "failed", "failures": failures})
        raise RuntimeError(f"Cannot finalize: {len(failures)} page worker(s) are incomplete or invalid")

    finalization_started = now_iso()
    merge_started = now_iso()
    t0 = time.perf_counter()
    merged = outdir / args.merged_output
    merge_refined_pages_openxml(page_files, merged)
    merge_seconds = round(time.perf_counter() - t0, 3)
    merge_finished = now_iso()
    merge_integrity = verify_merge_integrity(merged, len(tasks))
    post_merge_validation = post_merge_content_validation(merged, tasks, results)
    finalization_finished = now_iso()
    timing = concurrency_metrics(results)
    shared_groups = list(manifest.get("shared_masthead_groups") or [])
    expected_shared_pages = sum(len(group.get("members") or []) for group in shared_groups)
    passed_shared_pages = sum(
        1
        for result in results
        if result.get("shared_masthead_used") is True
        and result.get("shared_masthead_audit_status") == "passed"
    )
    shared_masthead_consistency = {
        "status": "passed" if passed_shared_pages == expected_shared_pages else "failed",
        "group_count": len(shared_groups),
        "expected_shared_page_count": expected_shared_pages,
        "passed_shared_page_count": passed_shared_pages,
        "validation": "identical_asset_sha256_and_normalized_placement_checked_per_page_before_merge",
    }
    acceptance = timing_acceptance(
        timing,
        page_count=len(tasks),
        worker_limit=int(manifest.get("worker_limit") or 1),
        min_speedup=args.min_speedup,
    )
    first_worker_start = parse_time(timing["worker_first_start"])
    true_end_to_end = (parse_time(finalization_finished) - first_worker_start).total_seconds()
    timing.update(
        {
            "skill_version": SKILL_VERSION,
            "page_count": len(tasks),
            "merge_started_at": merge_started,
            "merge_finished_at": merge_finished,
            "merge_seconds": merge_seconds,
            "finalization_started_at": finalization_started,
            "finalization_finished_at": finalization_finished,
            "end_to_end_seconds_from_first_worker": round(true_end_to_end, 3),
            "post_merge_content_validation": post_merge_validation,
            "shared_masthead_consistency": shared_masthead_consistency,
            "timing_acceptance": acceptance,
            "workers": [compact_worker_result(result) for result in results],
        }
    )
    write_json(outdir / "parallel_timing_report.json", timing)
    merge_passed = (
        merge_integrity["status"] == "passed"
        and post_merge_validation["status"] == "passed"
    )
    # Timing is reported independently from content/merge acceptance. A timing
    # miss must stay visible, but it must not strand already-passed page work or
    # retain temporary artifacts after a mechanically successful merge.
    timing_passed = acceptance["status"] == "passed"
    final_status = "completed" if merge_passed and timing_passed else "review_required"
    page_order = [
        {
            "uid": task["uid"],
            "task_index": task["task_index"],
            "source_index": task["source_index"],
            "source_input": task["source_input"],
            "page_index": task["page_index"],
            "source_sha256": sha256_file(Path(task["source_image"])),
        }
        for task in tasks
    ]
    final = {
        "status": final_status,
        "merged_pptx": str(merged),
        "page_count": len(tasks),
        "page_order": page_order,
        "page_terminal_validation_status": "passed",
        "shared_masthead_consistency": shared_masthead_consistency,
        "merge_integrity_status": merge_integrity["status"],
        "merge_integrity": merge_integrity,
        "post_merge_content_validation": post_merge_validation,
        "timing_status": acceptance["status"],
        "timing_report": str(outdir / "parallel_timing_report.json"),
        "cleanup": {"mode": "preserved_for_debugging", "status": "skipped"},
    }
    write_json(outdir / "parallel_v1_finalization.json", final)
    if final_status == "completed" and not getattr(args, "keep_working_artifacts", False):
        final["cleanup"] = cleanup_working_artifacts(outdir)
        write_json(outdir / "parallel_v1_finalization.json", final)
    print(json.dumps(final, ensure_ascii=False, indent=2))
    return 0 if final_status == "completed" else 2


def command_watch_finalize(args: argparse.Namespace) -> int:
    outdir = args.outdir.expanduser().resolve()
    lock_path = outdir / ".parallel_finalize.lock"
    lock_fd: int | None = None
    try:
        lock_fd = os.open(str(lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        os.write(lock_fd, f"pid={os.getpid()} started_at={now_iso()}".encode("utf-8"))
        deadline = time.monotonic() + float(args.timeout)
        while True:
            final_path = outdir / "parallel_v1_finalization.json"
            existing_final = read_json(final_path, {})
            if existing_final.get("status") == "completed":
                print(json.dumps(existing_final, ensure_ascii=False, indent=2))
                return 0
            manifest = read_json(outdir / "parallel_v1_manifest.json", {})
            tasks = manifest.get("tasks") or []
            if not tasks:
                raise RuntimeError("No prepared page tasks were found")
            results = [read_json(Path(task["worker_result"]), {}) for task in tasks]
            failed = [result for result in results if result.get("status") == "failed"]
            if failed:
                raise RuntimeError(f"Cannot auto-finalize: {len(failed)} page worker(s) failed")
            if all(result.get("status") == "completed" for result in results):
                finalize_args = argparse.Namespace(
                    outdir=outdir,
                    merged_output=args.merged_output,
                    min_speedup=args.min_speedup,
                    allow_timing_miss=args.allow_timing_miss,
                    keep_working_artifacts=args.keep_working_artifacts,
                )
                return command_finalize(finalize_args)
            if time.monotonic() >= deadline:
                raise TimeoutError("Timed out waiting for page workers")
            time.sleep(max(0.1, float(args.poll_seconds)))
    finally:
        if lock_fd is not None:
            os.close(lock_fd)
            try:
                if lock_path.exists():
                    lock_path.unlink()
            except OSError:
                pass


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Coordinate independent full-V1 page workers")
    sub = parser.add_subparsers(dest="command", required=True)

    preflight = sub.add_parser(
        "preflight",
        help="Classify raster/full-slide inputs without touching the formal output directory",
    )
    preflight.add_argument("inputs", nargs="+", type=Path)
    preflight.add_argument("--report", required=True, type=Path)
    preflight.add_argument("--intended-outdir", type=Path)
    preflight.add_argument("--available-worker-slots", type=int, required=True, choices=range(1, 17))
    preflight.add_argument("--root-task-count", type=int, default=1, choices=(1, 2))
    preflight.add_argument("--cpu-threads-per-worker", type=int, default=4)
    preflight.add_argument("--memory-gb-per-worker", type=float, default=4.0)
    preflight.add_argument("--powerpoint-worker-cap", type=int, default=4, choices=range(1, 17))
    preflight.set_defaults(func=command_preflight)

    prepare = sub.add_parser("prepare", help="Split inputs into isolated full-V1 page tasks")
    prepare.add_argument("inputs", nargs="+", type=Path)
    prepare.add_argument("--outdir", required=True, type=Path)
    prepare.add_argument("--workers", type=int, required=True, choices=range(1, 17))
    prepare.add_argument("--python", type=Path, default=Path(sys.executable))
    prepare.add_argument("--v1-skill", type=Path, default=DEFAULT_V1_SKILL)
    prepare.add_argument("--worker-contract", type=Path, default=DEFAULT_WORKER_CONTRACT)
    prepare.add_argument("--extractor", type=Path, default=DEFAULT_EXTRACTOR)
    prepare.add_argument("--validator", type=Path, default=DEFAULT_VALIDATOR)
    prepare.add_argument("--dpi", type=int, default=300)
    prepare.add_argument("--timeout", type=int, default=1800)
    prepare.add_argument("--preflight-report", required=True, type=Path)
    prepare.add_argument("--confirmed", action="store_true")
    prepare.add_argument("--capacity-override-confirmed", action="store_true")
    prepare.add_argument("--clean", action="store_true")
    prepare.set_defaults(func=command_prepare)

    start = sub.add_parser("mark-start", help="Record a page worker start time")
    start.add_argument("--task-json", required=True, type=Path)
    start.add_argument("--restart", action="store_true")
    start.set_defaults(func=command_mark_start)

    baseline = sub.add_parser("mark-baseline", help="Record an inspected V1 baseline checkpoint")
    baseline.add_argument("--task-json", required=True, type=Path)
    baseline.add_argument("--inspected", action="store_true")
    baseline.set_defaults(func=command_mark_baseline)

    record = sub.add_parser("record", help="Validate and record a page worker result")
    record.add_argument("--task-json", required=True, type=Path)
    record.add_argument("--status", required=True, choices=("completed", "failed"))
    record.add_argument("--pptx", type=Path)
    record.add_argument("--terminal-validation-report", type=Path)
    record.add_argument("--iteration-count", type=int, default=0)
    record.add_argument("--visual-inspection-confirmed", action="store_true")
    record.add_argument("--brand-masthead-status", choices=BRAND_MASTHEAD_STATUS_CHOICES)
    record.add_argument("--shared-masthead-used", action="store_true")
    record.add_argument("--figure-text-status", choices=FIGURE_TEXT_STATUS_CHOICES)
    record.add_argument(
        "--card-decomposition-status", choices=CARD_DECOMPOSITION_STATUS_CHOICES
    )
    record.add_argument("--scientific-chart-status", choices=SCIENTIFIC_CHART_STATUS_CHOICES)
    record.add_argument("--corrections-applied", action="append", default=[])
    record.add_argument("--started-at")
    record.add_argument("--finished-at")
    record.add_argument("--notes", default="")
    record.set_defaults(func=command_record)

    finalize = sub.add_parser(
        "finalize",
        help="Merge terminally validated V1 pages and perform mechanical integrity assertions",
    )
    finalize.add_argument("--outdir", required=True, type=Path)
    finalize.add_argument("--merged-output", default="merged_v1_refined_editable.pptx")
    finalize.add_argument("--min-speedup", type=float, default=1.4)
    finalize.add_argument("--allow-timing-miss", action="store_true")
    finalize.add_argument("--keep-working-artifacts", action="store_true")
    finalize.set_defaults(func=command_finalize)

    watch = sub.add_parser(
        "watch-finalize",
        help="Wait for all page workers and merge immediately after the last terminal pass",
    )
    watch.add_argument("--outdir", required=True, type=Path)
    watch.add_argument("--poll-seconds", type=float, default=1.0)
    watch.add_argument("--timeout", type=float, default=21600.0)
    watch.add_argument("--merged-output", default="merged_v1_refined_editable.pptx")
    watch.add_argument("--min-speedup", type=float, default=1.4)
    watch.add_argument("--allow-timing-miss", action="store_true")
    watch.add_argument("--keep-working-artifacts", action="store_true")
    watch.set_defaults(func=command_watch_finalize)
    return parser


def main() -> int:
    args = build_parser().parse_args()
    try:
        return int(args.func(args))
    except Exception as exc:
        print(json.dumps({"status": "error", "error": str(exc)}, ensure_ascii=False, indent=2), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
