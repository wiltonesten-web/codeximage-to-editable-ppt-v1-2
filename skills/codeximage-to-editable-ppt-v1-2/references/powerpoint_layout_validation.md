# Terminal page validation and simple merge

Page workers own all V1 content quality. Complete the strict PowerPoint acceptance inside each isolated page task before the page enters the merge queue.

1. Render the corrected one-page deck with Microsoft PowerPoint.
2. Compare that render with the page's source image.
3. For a fully editable V1 reconstruction, use MAE at most `35` and changed-pixel ratio at most `0.30` as guardrails, with zero text overflow and zero off-slide shapes. These looser image thresholds recognize intentional native-text/native-shape reconstruction; they do not excuse an obvious visual defect.
4. Open the terminal render. Automated thresholds never replace visual inspection.
5. Correct and repeat terminal validation until the page passes.
6. Before recording completion, run the automatic raster-integrity gate: reject pictures covering at least 90% of the slide, source-screenshot picture reuse, and detected raster text underneath editable text.
7. Record page completion only after raster integrity, the refined deck, and the terminal report/render pass and the worker embeds its baseline checkpoint, review count, correction summary, and visual-inspection confirmation in `worker_result.json`. Do not require a duplicate preview or separate review JSON.
8. Merge only raster-integrity- and terminally-passed one-page decks, in manifest order, using deterministic OpenXML copying.
9. After merge, do not render or compare the merged deck. Assert only that it saved, opens, contains the expected page count, and was assembled in deterministic manifest order.

If Microsoft PowerPoint is unavailable, the page cannot pass terminal validation and must not enter the merge queue.
