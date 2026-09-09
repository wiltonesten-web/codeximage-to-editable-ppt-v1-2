# codeximage-to-editable-ppt-v1-2

[简体中文](README.md) | **English**

An upgraded orchestration layer for
[`codeximage-to-editable-ppt-v1`](https://github.com/wiltonesten-web/codeximage-to-editable-ppt-v1),
designed for multi-page image-based PowerPoint files. It preserves the complete
single-page V1 quality contract while using measured parallel execution for
editable reconstruction. It can process 3, 6, 9, or other numbers of images in
parallel when the corresponding task-tree and hardware capacity is available.

> **Efficiency improvement:** Independent page workers reconstruct multiple
> slides at the same time, reducing the total waiting time for multi-page jobs
> without lowering the V1 single-page quality standard. Actual speedup is
> reported from measured worker overlap and runtime results.

The packaged skill version is **1.2.7**. It adds hardware-aware preflight,
multi-root page orchestration, reusable branded mastheads, stricter scientific
chart and panel-marker checks, structured-card decomposition rules, ordered
OpenXML merge, and post-merge content invariants.

## Highlights

- Preflight multi-page slide images or image-based PPT/PPTX pages before writing
  to the formal output directory.
- Estimate a conservative worker count from CPU, available memory, active
  PowerPoint processes, and available Codex worker slots.
- Run up to three page workers in each Codex task tree, scaling to 3, 6, 9, or
  other page counts through explicitly authorized task trees when capacity permits.
- Detect a repeated branded masthead and reuse one immutable, hash-verified crop
  at a consistent normalized placement.
- Reject full-slide screenshot reuse, raster text residue under editable text,
  under-split information cards, and masthead/title collisions.
- Preserve each logical scientific chart as one pixel-exact source crop with its
  title, axes, ticks, values, units, legends, annotations, colorbar, and safety
  margins.
- Preserve external scientific panel markers such as `(a)`, `(b)`, and `(c)`
  with exact tag-to-text validation.
- Merge passed one-page decks in source order while preserving audit tags, then
  verify fast post-merge content invariants.
- Report observed worker overlap and refuse to claim a parallelism level that
  was not actually measured.

## Relationship to V1

This repository is an upgrade layer, not a replacement copy of V1. The installed
v1.2 skill reads the canonical V1 single-page reconstruction contract from the
sibling folder:

```text
%USERPROFILE%\.codex\skills\codeximage-to-editable-ppt-v1
```

Install V1 first, then install this repository's v1.2 skill. Both folders must
remain installed side by side.

## Install as a Codex Skill

Install the required V1 base skill:

```powershell
python "$env:USERPROFILE\.codex\skills\.system\skill-installer\scripts\install-skill-from-github.py" `
  --repo wiltonesten-web/codeximage-to-editable-ppt-v1 `
  --path skills/codeximage-to-editable-ppt-v1
```

Install v1.2:

```powershell
python "$env:USERPROFILE\.codex\skills\.system\skill-installer\scripts\install-skill-from-github.py" `
  --repo wiltonesten-web/codeximage-to-editable-ppt-v1-2 `
  --path skills/codeximage-to-editable-ppt-v1-2
```

Restart Codex after installation so the new skill is discovered.

For manual installation, copy only
`skills/codeximage-to-editable-ppt-v1-2` into `%USERPROFILE%\.codex\skills\`.
Do not copy the repository root as the skill folder.

## Python Dependencies

Install the Python packages from the repository root:

```powershell
python -m pip install -r skills/codeximage-to-editable-ppt-v1-2/requirements.txt
```

Microsoft PowerPoint is required for the strict final rendering and layout QA
workflow on Windows. Tesseract OCR, LibreOffice, and Poppler may be needed for
particular input or fallback paths.

## Typical Workflow

Use the skill in Codex with a request such as:

```text
Use $codeximage-to-editable-ppt-v1-2 to preflight these slide images and rebuild
them as an editable deck. Show me the page strategy and safe worker choices before
starting page workers.
```

The skill first produces a read-only preflight report and asks the user to
confirm the page strategy and exact worker count. Execution that requires
additional task trees also requires explicit user authorization. It does not
silently fall back to sequential waves while claiming full concurrency.

## Command-Line Preflight

Example for six raster slide pages:

```powershell
python skills/codeximage-to-editable-ppt-v1-2/scripts/parallel_v1_pages.py preflight `
  slide01.png slide02.png slide03.png slide04.png slide05.png slide06.png `
  --report output.preflight.json `
  --intended-outdir output `
  --available-worker-slots 3 `
  --root-task-count 2 `
  --powerpoint-worker-cap 6
```

Preflight does not perform the full editable reconstruction. Continue through
Codex so that page workers can execute the V1 reconstruction and PowerPoint QA
contract.

## Default Successful Output

```text
output/
|-- merged_v1_refined_editable.pptx
|-- parallel_timing_report.json
`-- parallel_v1_finalization.json
```

When validation fails, diagnostic task and shard artifacts are preserved for
review instead of presenting the run as a completed delivery.

## Repository Layout

```text
skills/
`-- codeximage-to-editable-ppt-v1-2/
    |-- SKILL.md
    |-- VERSION
    |-- agents/
    |-- references/
    |-- scripts/
    |-- config.example.yaml
    `-- requirements.txt
```

## Validation

Validate the skill metadata:

```powershell
python "$env:USERPROFILE\.codex\skills\.system\skill-creator\scripts\quick_validate.py" `
  skills/codeximage-to-editable-ppt-v1-2
```

Run the included tests in an environment where the V1 and v1.2 skill folders are
installed side by side:

```powershell
python -m unittest discover `
  -s skills/codeximage-to-editable-ppt-v1-2/scripts `
  -p "test_*.py" -v
```

## Limitations

- The skill coordinates multi-page inputs; use the base V1 skill directly for a
  single page.
- Higher parallel counts such as 6 or 9 depend on enough available Codex task
  trees, sufficient hardware capacity, and explicit user authorization for
  auxiliary tasks.
- A selected worker count above the hardware-safe recommendation requires an
  explicit capacity override.
- Raster sources cannot recover hidden text, original vectors, chart data,
  animations, or font files that are absent from the input.
- Strict delivery depends on a faithful Microsoft PowerPoint render and terminal
  validation; a script-only baseline is not a finished editable deck.

## Project Status

This repository packages v1.2.7 as an independent community project. It is not
affiliated with or endorsed by OpenAI or Microsoft.

## License

MIT License. See [LICENSE](LICENSE).
