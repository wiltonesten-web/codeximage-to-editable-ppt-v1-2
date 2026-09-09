---
name: codeximage-to-editable-ppt-v1-2
description: Preflight two to six image-based PPT/PPTX pages or slide screenshots, estimate safe hardware capacity, detect reusable branded mastheads, ask the user to confirm the page strategy and worker count, and preserve the complete original codeximage-to-editable-ppt-v1 reconstruction and PowerPoint QA workflow. Use one Codex task tree for up to three simultaneous page workers and, with explicit user authorization, two independent Codex task trees for four to six simultaneous full-V1 page workers. Reuse each confirmed common masthead as one immutable shared asset; reject full-slide screenshot reuse, raster text residue, masthead/title collisions, crude branded masthead replacement, incorrect external subfigure markers, under-split information cards, padded or altered chart crops, and data charts missing titles, axes, ticks, values, units, legends, annotations, or colorbars. Require exact source-crop provenance, measured worker overlap, audit-tag-preserving ordered merge, and fast post-merge content invariants.
---

# Multi-root Full-V1 Page Orchestrator v1.2.7

## Core contract

Preserve the original V1 single-page reconstruction and review quality. Increase real page concurrency without replacing AI page workers with ordinary subprocess builders.

- For one page, use `../codeximage-to-editable-ppt-v1/SKILL.md` directly.
- For two or three selected workers, use one Codex task tree and one collaboration subagent per page.
- For four to six selected workers, require two Codex task trees. The primary task handles shard 1; create one auxiliary Codex task for shard 2 only after the user explicitly authorizes that task creation.
- Never claim six-page parallelism unless the final timing report measures six overlapping page-worker intervals.
- If task creation or cross-task concurrency is unavailable, stop and report that the selected parallel target cannot be met. Do not silently run two sequential waves.
- Detect common mastheads during preflight. After confirmation, crop each accepted masthead group once from its medoid page, save it as an immutable shared asset, and give every assigned worker the same asset hash and normalized placement.

Before preparing tasks, the primary coordinator reads these canonical files once:

- `../codeximage-to-editable-ppt-v1/SKILL.md`
- `../codeximage-to-editable-ppt-v1/references/refined_rebuild_workflow.md`

Page workers read only their copied `page_worker_contract.md`, `worker_request.md`, source page, and task files. The auxiliary root reads only its generated `shard_coordinator_request.md`. Neither rereads the canonical skills.

## Full-V1 page contract

Every page worker independently performs:

1. baseline decomposition and visual inspection;
2. element-specific masks, crops, editable text, and native shapes;
3. refined one-page editable reconstruction;
4. Microsoft PowerPoint render and source comparison;
5. correction of residue, crop, overlap, clipping, wrapping, overflow, off-slide, font, layer, and alignment defects;
6. terminal Microsoft PowerPoint validation and final-render inspection;
7. repeated correction and validation until the page passes.

Hard rules:

- Never place the source screenshot or a re-encoded copy as a full-slide picture.
- Reject any picture covering 90% or more of the slide.
- Never hide source text with rectangles and overlay editable text.
- Reject raster text, partial glyphs, or antialias residue underneath editable text.
- Preserve branded mastheads, institutional headers, logos, and distinctive header frames as high-fidelity components. If a header has angled geometry, gradient bands, seals, or combined logo/wordmark artwork, crop the real header frame or logo as independent tight elements and combine them with editable text when practical; do not replace them with crude rectangles merely because the pixel diff passes.
- When the coordinator assigns a shared masthead, insert that exact asset unchanged at the supplied normalized coordinates. Do not independently crop, redraw, OCR, recolor, or reposition it. Shared brand wordmarks remain inside the shared raster asset.
- Treat a structured information card as layout, not as one picture. A card with a title/header, central photo or complex illustration, icons, dividers, and explanatory rows must be decomposed: keep only the irreducible photo/illustration as raster; rebuild the card border, fills, and dividers as native shapes; make the title and explanatory labels editable; and keep each icon as its own tight crop or native icon. A tight full-card screenshot is forbidden even when it covers far less than 90% of the slide. This applies to comparison, benefit/risk, input-response-outcome, and similar infographic cards.
- Preserve each logical data chart as one complete raster panel. Include its full plotting area, all axis titles, outermost ticks and values, units, legends, annotations, and colorbar, plus a clean safety margin of at least 8 source pixels or 1% of panel width/height, whichever is larger, on all four sides. Preserve the chart or subfigure title exactly once: preferably rebuild a cleanly separable title as editable text at the original position; otherwise keep it inside the complete raster crop. Never omit it. Never split one continuous chart into upper/lower or left/right fragments to evade card QA.
- A data-chart raster must be a pixel-exact rectangular crop of the source page. Transparent or white padding canvases, label erasure, generative fill, resampling, and PowerPoint secondary crop are forbidden. Tag it `SCI_CHART_COMPLETE::<stable_panel_id>::SRCBOX=<left>,<top>,<width>,<height>` using exact source pixels. The record gate compares the embedded bitmap with that source rectangle, so a cropped axis/title cannot be hidden by adding blank margin.
- Tag every editable external marker as `SCI_PANEL_MARKER::<stable_panel_id>::(<letter>)`; its text must equal the encoded marker exactly. This prevents `(c)` from silently becoming another glyph.
- For scientific figures, plots, heatmaps, tables, molecular diagrams, and paper-style multi-panel images, extract only external subfigure markers such as `(a)`, `(b)`, `(c)` as editable PowerPoint text when they sit outside the actual plot/image body or can be cleanly separated. Crop each major panel as an independent tight picture without the external marker when practical, then add the marker back as editable text. Keep all plot/image-body text inside the raster crop, including axis letters such as `i`/`j`, in-panel labels, legends, tick values, heatmap cell values, molecule labels, colorbar labels, formulas, and embedded panel captions such as `11 Å-NEP`, unless the user explicitly asks for full figure text extraction.

Workers write only the refined PPTX, terminal validation evidence, and temporary `worker_result.json`. Do not generate ZIP archives, duplicate previews, CSV manifests, or delivery bundles. The completion record must include `brand_masthead_status`, `figure_text_status`, `card_decomposition_status`, and `scientific_chart_status`; an assigned shared masthead additionally requires `shared_faithful`, explicit use confirmation, matching asset SHA-256, matching normalized placement, and zero secondary crop.

The `record` gate checks raster integrity, text residue, structured cards, tagged scientific charts, panel-marker identity, and masthead/title clearance. A tagged data chart fails when its source box is absent or invalid, its bitmap differs from the exact source crop, it uses transparency/padding or PowerPoint secondary crop, it has excessive foreground ink touching an outer edge, duplicates a panel ID, or declares a fragment. Card QA skips tagged complete scientific visuals, preventing chart titles and axes from being mistaken for card text. Plot-internal labels remain rasterized.

## Coordinator workflow

### 1. Preflight capacity and request confirmation

Use the configured project Python environment with the packages in
`requirements.txt`. Prefer an isolated Python 3 virtual environment.

For a possible six-worker run, run read-only preflight with two task roots and three child slots per root:

```powershell
python scripts/parallel_v1_pages.py preflight input1.png input2.png input3.png input4.png input5.png input6.png `
  --report output.preflight.json `
  --intended-outdir output `
  --available-worker-slots 3 `
  --root-task-count 2 `
  --powerpoint-worker-cap 6
```

Preflight must not touch the formal output directory. Report:

- page classifications and full-slide raster prohibitions;
- detected shared masthead groups, their canonical medoid pages, reusable page count, and pages that require page-specific handling;
- editable-text, clean-background, and independent-raster strategies;
- CPU, available memory, active PowerPoint processes, and conservative safe recommendation;
- page count, page-worker slots per root, total two-root capacity, allowed choices, and whether the selected count exceeds the safe recommendation;
- quality risks and estimated time.

Stop and ask the user to confirm:

1. the page strategy;
2. the exact page-worker count;
3. when selecting four to six, authorization to create one auxiliary Codex task;
4. when selecting above the hardware-safe recommendation, the experimental capacity override.

Do not prepare, create another task, or dispatch workers before all applicable confirmations are explicit.

### 2. Prepare ordered tasks and shards

For six selected workers:

```powershell
python scripts/parallel_v1_pages.py prepare input1.png input2.png input3.png input4.png input5.png input6.png `
  --outdir output `
  --workers 6 `
  --preflight-report output.preflight.json `
  --confirmed `
  --capacity-override-confirmed `
  --clean
```

`prepare` verifies source hashes, order, confirmation state, worker count, capacity override, and line-ending-normalized compact-contract source hashes. It creates contiguous ordered shards. For six pages and two roots:

```text
shard_001: pages 1, 2, 3
shard_002: pages 4, 5, 6
```

Before creating page requests, `prepare` creates one immutable PNG per confirmed shared masthead group under temporary `shared_assets/mastheads/`. Each task receives only its assigned asset path, hash, and normalized placement. Unassigned pages keep the page-specific masthead workflow.

### 3. Dispatch one or two task trees

For one root, spawn the assigned page workers back-to-back and start the watcher.

For two roots:

1. Read `shards/shard_001/shard_coordinator_request.md` and spawn its page workers in the current task.
2. Use the Codex task-creation tool to create one auxiliary local task with the exact contents/path of `shards/shard_002/shard_coordinator_request.md`. User confirmation of four-to-six parallelism is the required task-creation authorization.
3. Do not override the user's configured model or reasoning level.
4. Start the global `watch-finalize` process in the primary task immediately after both roots are dispatched.
5. Monitor the auxiliary task with the task waiting tool. Leave approval or input requests visible to the user.
6. The auxiliary root may dispatch only its assigned pages; it must not merge, clean, or inspect shard 1.

Retry only failed pages. Do not recreate completed pages.

### 4. Finalize and prove concurrency

The global watcher merges only after every page has passed raster integrity, structured-card decomposition, source-grounded tagged scientific-chart completeness, exact panel-marker identity, shared-masthead consistency and title clearance when assigned, and terminal PowerPoint validation:

```powershell
python scripts/parallel_v1_pages.py watch-finalize --outdir output
```

The final timing gate requires:

```text
observed_peak_concurrent_workers >= min(selected_workers, page_count)
```

For six selected workers and six pages, the measured peak must be six. A peak of three means the two roots were serialized or throttled; preserve diagnostics and return `review_required`. Never label it six-page parallel success.

Passed one-page decks are merged once with deterministic OpenXML copying. The merger must preserve picture names and all audit tags. After merge, run one fast invariant audit—without a second render/correction loop—to verify per-page chart-tag counts, exact source-crop provenance, panel-marker names/text, shared masthead placement, and masthead/title clearance. A failed invariant keeps diagnostics and returns `review_required`; simple page-count validation alone is insufficient.

## Final outputs

Successful default delivery contains only:

```text
output/
|-- merged_v1_refined_editable.pptx
|-- parallel_timing_report.json
`-- parallel_v1_finalization.json
```

Timing or page failures preserve task and shard artifacts for diagnosis. Report total user-visible task time, worker wall time, sequential-equivalent time, required and observed peak concurrency, measured speedup, per-page terminal metrics, brand masthead status, figure text status, structured-card status, scientific-chart status, merge integrity, and cleanup status.

Read `references/parallel_v1_orchestration.md` for the multi-root execution graph and failure rules.
