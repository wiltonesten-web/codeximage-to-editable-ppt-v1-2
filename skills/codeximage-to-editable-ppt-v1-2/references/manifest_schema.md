# Manifest schema

The manifest is the source of truth for recomposition. It is written as both CSV and JSON.

Core fields:

| Field | Meaning |
|---|---|
| slide_id | Page identifier such as `slide01` or `image01` |
| element_id | Unique element identifier |
| file_name | PNG filename |
| relative_path | Path relative to output root |
| element_type | Element category such as `title_txt`, `figure`, `icon`, `line` |
| is_text_only | `yes` or `no` |
| recognized_text | OCR text for text-only elements |
| OCR_confidence | Mean OCR confidence |
| x, y, width, height | Pixel bbox in source page coordinates |
| z_order | Layer order from background to foreground |
| opacity | Estimated opacity; usually 1.0 |
| rotation | Estimated rotation; currently usually 0 |
| crop_mask_type | `rectangle`, `circle`, `ellipse`, `irregular_mask`, `full_background` |
| transparent_background | `yes` when the PNG uses alpha transparency outside the visible element contour; otherwise `no` |
| alpha_note | How alpha was applied, or why transparent background was not reliable |
| shadow_effect | Detected or inferred shadow note |
| converted_to_editable_text | `yes`, `no`, or `n/a` |
| conversion_note | Reason for conversion or fallback |
| page_source_type | How the source page image was obtained |
| source_page_path | Path to extracted/rendered source page image |
| source_width, source_height | Source page dimensions in pixels |
| dominant_color | Median visible color of element |
| font_family_guess | Estimated/default font family |
| font_size_guess | Estimated font size in points |
| font_color_guess | Estimated text color |
| parent_group_id | Optional inferred group ID |
| needs_manual_review | `yes` or `no` |

For `icon`, `logo`, `symbol`, `button`, `badge`, `arrow`, `line`, `divider`, `shape`, `simple_shape`, and text-only elements, `transparent_background` should normally be `yes`. If it is `no`, `alpha_note` must explain the reason, such as `complex_background`, `low_mask_confidence`, `edge_color_blending`, `shadow_requires_background`, `rectangular_crop_fallback`, or `transparent_required_but_unreliable`.
