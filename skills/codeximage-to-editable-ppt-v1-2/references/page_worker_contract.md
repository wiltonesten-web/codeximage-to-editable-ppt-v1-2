# Compact full-V1 page worker contract

This is the complete operational contract for one isolated page worker. The coordinator has already read the canonical V1 instructions and verifies their line-ending-normalized SHA-256 hashes before creating this copy. Do not reread the full V1 or v1.2 skill files.

## Required page workflow

1. Start timing before any decomposition work.
2. Run the supplied baseline command. Inspect the JSON manifest, element crops, overlay, and rough recomposed preview; then record the baseline checkpoint.
3. Rebuild the page as an editable one-slide PPTX using the source image as ground truth:
   - editable text stays editable;
   - simple lines, boxes, dividers, and fills use native PowerPoint shapes where practical;
   - complex scientific plots, molecular diagrams, photographs, logos, and intricate illustrations may remain cropped raster elements;
   - branded mastheads, institution headers, logo bands, seals, wordmarks, and distinctive page-header frames must be treated as high-fidelity brand components. Crop angled/gradient/combined header artwork as tight independent images when native shapes would be visibly crude, and add separate editable text or logo crops as appropriate;
   - if the coordinator assigns a shared masthead, use the supplied immutable asset exactly once at the supplied normalized placement. Do not independently crop, redraw, OCR, recolor, resize, or reposition it; keep the shared brand wordmark inside that shared raster asset;
   - distinguish scientific figure bodies from structured information cards. A card with a title/header, central photo or intricate illustration, icons, separators, and explanatory rows must not remain one raster picture. Keep only the irreducible photo/illustration rasterized; rebuild the card border, fills, separators, title, labels, and callouts as editable PowerPoint objects; use one tight crop or native object per icon. This includes comparison, benefit/risk, input-response-outcome, and similar infographic cards;
   - preserve each logical data chart as one complete picture. Its source crop must include the full plot, every axis title, outermost tick/number, unit, legend, annotation, and colorbar plus a clean margin of at least 8 source pixels or 1% per side. Preserve the chart or subfigure title exactly once: preferably recreate a cleanly separable title as editable text in the original position; otherwise keep it inside the complete crop. Never omit the title and never split a continuous chart into strips or fragments. Export an exact unmodified rectangular source crop—no padding canvas, erased text, generative fill, resampling, transparency, or secondary crop—and name it `SCI_CHART_COMPLETE::<stable_panel_id>::SRCBOX=<left>,<top>,<width>,<height>`. Name retained non-chart scientific visuals `SCI_IMAGE_COMPLETE::<stable_panel_id>`;
   - name every editable external panel marker `SCI_PANEL_MARKER::<stable_panel_id>::(<letter>)` and keep its actual text exactly equal to the marker in the name;
   - for scientific figures, plots, heatmaps, tables, molecular diagrams, and paper-style multi-panel images, extract only external subfigure markers such as `(a)`, `(b)`, `(c)` as editable PowerPoint text when they sit outside the actual plot/image body or can be cleanly separated. For a multi-panel figure, crop each major panel as its own tight picture without the external marker when practical, then add the marker back as editable text. Keep all plot/image-body text inside the raster crop, including axis letters such as `i`/`j`, in-panel labels, legends, tick values, heatmap cell values, molecule labels, colorbar labels, formulas, and embedded panel captions such as `11 Å-NEP`, unless the user explicitly asks for full figure text extraction;
   - never use the source screenshot, or a re-encoded copy, as a full-slide picture;
   - no picture may cover 90% or more of the slide; construct backgrounds from native fills plus tightly cropped independent decorations;
   - never hide original screenshot text with rectangles and then place editable text on top;
   - icons, logos, and simple raster elements require element-specific alpha masks and tight crops;
   - inspect every retained raster for accidental text, background, halo, or neighboring-object residue;
   - any raster underneath editable text must contain no old text, partial glyph, or antialias fringe.
4. Preserve the source canvas, layout, alignment, layer order, typography, colors, and visual hierarchy. Do not reuse another page's builder or result.
5. Render with Microsoft PowerPoint, compare the render with the source, and correct crop errors, residue, overlap, clipping, wrapping, font substitution, text overflow, off-slide objects, and conspicuous displacement.
6. Run the supplied terminal validator only after corrections. Open its final PowerPoint render and repeat correction plus validation until the report passes.
7. Record completion with a truthful review iteration count, explicit visual-inspection confirmation, and a concise correction summary.

## Acceptance gate

Completion is accepted only when all of the following are true:

- the refined PPTX exists and contains exactly one slide;
- no picture covers 90% or more of the slide;
- no embedded picture is byte-identical to the source screenshot;
- the raster-residue audit finds no old text underneath editable text;
- branded mastheads or institution headers are either faithfully preserved as independent high-fidelity components or explicitly marked `not_applicable`;
- an assigned shared masthead is embedded with the coordinator-provided SHA-256, exact normalized placement, and no secondary crop, and is recorded as `shared_faithful` with explicit shared-use confirmation;
- cleanly separable external subfigure markers such as `(a)`, `(b)`, `(c)` are either extracted as editable text or explicitly marked `no_external_panel_markers`;
- structured information cards are decomposed into editable card chrome/text and irreducible image content, or the page is explicitly marked `no_structured_cards`;
- no retained portrait card raster contains both a readable title band and two or more readable explanatory rows in its footer;
- every data chart is recorded as `charts_complete` and represented by exactly one tagged complete picture per logical panel, or the page is recorded as `no_data_charts`;
- every chart or subfigure title is present exactly once, either as editable text at the original position or inside its complete raster panel;
- no data chart differs from its declared exact source rectangle, uses transparent/white padding, erasure, resampling, a secondary PowerPoint crop, touches a crop edge without the required clean safety margin, duplicates a panel ID, or uses `SCI_CHART_FRAGMENT::`;
- every extracted panel marker has a valid marker tag and exact matching text;
- no editable title or other page text overlaps the assigned shared masthead;
- the baseline inspection checkpoint was recorded;
- at least one PowerPoint render/comparison iteration was performed;
- visual inspection of the terminal render is explicitly confirmed;
- the terminal report uses `microsoft_powerpoint` in `final_visual` mode;
- the terminal report contains exactly one passed slide, references this task's source image, and points to an existing render;
- text overflow and off-slide shape counts are both zero.

The terminal report's MAE and changed-pixel ratio, review count, brand masthead status, figure text status, structured-card status, scientific-chart status, correction summary, and timing are embedded into the temporary `worker_result.json`; a second page-review JSON is not required.

The completion record checks raster integrity, old-text residue underneath editable text, and the common under-split portrait information-card pattern. It does not fail merely because retained figure crops contain internal scientific labels; the required editable extraction scope is only the cleanly separable external subfigure marker layer.

## Artifact policy

- Work only inside the assigned task directory.
- Do not inspect another worker's files, edit the shared manifest, merge decks, or spawn page workers.
- Do not create ZIP archives, delivery bundles, duplicate manifests, or standalone diff packages.
- Baseline crops, overlay, rough preview, terminal render, task packet, and worker result are temporary working files. The coordinator removes them automatically after the passed one-page decks are merged.
- On failure, leave the task directory intact so only that page can be retried or diagnosed.
