# Multi-root full-V1 orchestration

## Execution graph

```text
primary coordinator reads canonical V1 once
  -> read-only preflight classifies pages and estimates capacity
  -> preflight clusters matching top mastheads and selects one medoid source per reusable group
  -> user confirms strategy, worker count, auxiliary task, and any capacity override
  -> prepare verifies hashes, crops each confirmed shared masthead once, and creates ordered page tasks plus contiguous shards
  -> primary root dispatches shard 1 page workers
  -> auxiliary Codex task dispatches shard 2 page workers
  -> all page workers run the compact full-V1 contract independently
  -> primary watcher sees all passing terminal results and shared-masthead hash/placement gates
  -> require measured peak concurrency equal to the selected page-worker target
  -> merge in manifest order while preserving picture names and audit tags
  -> assert saved/openable/page-count/order integrity
  -> run fast post-merge chart-source, marker, masthead, and title-clearance invariants without a second render loop
  -> clean temporary artifacts only after page, merge, and timing gates pass
```

## Root and shard rules

- Use one root for one to three selected page workers.
- Use two independent Codex task roots for four to six selected workers.
- Require explicit user authorization before creating the auxiliary task.
- Assign contiguous pages to shards so source order is obvious. Six pages split as 1–3 and 4–6.
- Each root spawns only its assigned compact page workers.
- Only the primary root runs the global watcher and final merge.
- Absolute task paths are the synchronization boundary; roots do not exchange page content through chat.

## Shared masthead rules

- Detect each full-width masthead boundary from the upper color transition (with a conservative 10% fallback), normalize those bands, compare both spatial structure and brand color, join transitively matching same-brand bands, and group them before choosing the shared standard.
- Select the medoid page as the canonical source so one outlier does not define the shared header.
- Materialize one immutable PNG per confirmed group after user confirmation, never during read-only preflight.
- Give assigned page workers the same asset path, SHA-256, and normalized placement; workers must not recreate or correct it independently.
- Validate the embedded image hash, normalized position and size, and zero secondary crop before accepting each page.
- Keep pages without a reliable match on the page-specific masthead path.

## Capacity and confirmation

Preflight records two limits:

- `recommended_workers`: conservative hardware-safe recommendation;
- `maximum_allowed_workers`: page-count and two-root Codex capacity, capped at six.

Selecting above the recommendation requires `--capacity-override-confirmed`. Selecting four to six also authorizes one auxiliary Codex task. If the task-creation tool is unavailable, fail before page work begins.

## Truthful timing

For worker intervals `d_i`, earliest start `S`, latest finish `F`, and merge finish `M`:

```text
sequential_equivalent = sum(d_i)
worker_wall_time = F - S
measured_speedup = sequential_equivalent / worker_wall_time
true_end_to_end = M - S
required_peak = min(selected_workers, page_count)
```

The run passes parallelism acceptance only when observed peak concurrency reaches `required_peak`. For six selected workers and six pages, require a measured peak of six. Cross-task throttling is a timing failure, not a quality failure, but the run must not be reported as successful six-page parallelism.

## Completion and cleanup

- Retry only missing or failed pages.
- Preserve passed page results when another page or root fails.
- Merge only terminally validated, raster-integrity-passed one-page decks.
- Verify accepted PPTX hashes before merge.
- Preserve all `SCI_CHART_COMPLETE`, `SCI_IMAGE_COMPLETE`, and `SCI_PANEL_MARKER` names during every batch or cross-batch OpenXML merge.
- After every final merge, verify exact chart source boxes, expected chart/marker tag counts, exact marker text, shared masthead placement, and masthead/title clearance. Do not accept page-count-only validation.
- Keep shards and page tasks when concurrency, page validation, or merge integrity fails.
- On complete success, remove sources, page tasks, shard requests, manifests, previews, renders, and worker synchronization files.
