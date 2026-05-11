"""Agent-loop slide generator — generate → evaluate → iterate until style match.

For each slide:
  1. Match slide content to the correct spec page type and layout variant
  2. Generate SVG with style guidance from spec
  3. Evaluate style similarity against spec reference (5 dimensions)
  4. If score < 0.90, add fix instructions and regenerate (max 5 iterations)
  5. Return best-scoring SVG
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from ppt_skill.style_evaluator import evaluate_style, StyleReport
from ppt_skill.spec.spec_model import PageType, LayoutSubType


MAX_ITERATIONS = 5
PASS_THRESHOLD = 0.90


@dataclass
class SlideResult:
    """Result of generating one slide with agent loop."""
    slide_index: int = 0
    svg_text: str = ""
    page_type: str = ""
    layout_sub_type: str = ""
    iterations: int = 0
    best_score: float = 0.0
    reports: list[StyleReport] = field(default_factory=list)
    passed: bool = False


def _match_spec_page(
    target_page_type: str,
    target_layout: str | None,
    spec_pages: list[dict],
) -> dict | None:
    """Find the best matching spec page for the target content type.

    Priority:
      1. Exact page_type + exact layout_sub_type match
      2. Exact page_type + any layout (for content pages)
      3. Exact page_type match
      4. Any page_type match (fallback)
    """
    if not spec_pages:
        return None

    # Priority 1: exact match
    for page in spec_pages:
        pt = page.get("page_type", "")
        lt = page.get("layout_sub_type", "")
        if pt == target_page_type and lt == target_layout:
            return page

    # Priority 2: page_type match (for content)
    if target_page_type == "content":
        for page in spec_pages:
            if page.get("page_type") == "content":
                return page

    # Priority 3: page_type match
    for page in spec_pages:
        if page.get("page_type") == target_page_type:
            return page

    # Priority 4: fallback to first page
    return spec_pages[0]


def _build_generation_prompt(
    slide_entry: dict[str, Any],
    spec_page: dict[str, Any],
    previous_issues: list[str] | None = None,
) -> str:
    """Build the SVG generation prompt with style guidance from spec."""
    title = slide_entry.get("title", "")
    body = slide_entry.get("body", [])
    page_type = slide_entry.get("page_type", "content")
    layout_type = slide_entry.get("layout_type", "full_width")

    # Spec style guidance
    bg_color = spec_page.get("background_color", "#FFFFFF")
    bg_description = spec_page.get("background_description", "Plain white background")
    layout_description = spec_page.get("layout_description", "")
    colors = spec_page.get("colors", {})
    typography = spec_page.get("typography", {})
    assets_dir = spec_page.get("assets_dir", "")
    design_patterns = spec_page.get("design_patterns", [])
    element_groups = spec_page.get("element_groups", [])
    elements = spec_page.get("elements", [])

    color_guide = "\n".join(
        f"    {k}: {v}" for k, v in colors.items() if v
    ) if colors else ""

    font_family = typography.get("heading_family", typography.get("body_family", "Arial"))

    # ── Build element layout guide from spec ──
    element_lines = []
    for el in elements[:15]:  # Limit to first 15 for prompt length
        eid = el.get("id", 0)
        etype = el.get("element_type", "?")
        erole = el.get("semantic_role", "")
        pos = el.get("position", {})
        text_preview = el.get("text", "")[:60]

        line = f"  Element #{eid}: type={etype}, role={erole}"
        if pos:
            line += f", pos=({pos.get('x',0):.2f},{pos.get('y',0):.2f}) size=({pos.get('w',0):.2f}x{pos.get('h',0):.2f})"
        if text_preview:
            line += f", text='{text_preview}'"
        if el.get("fill_color"):
            line += f", fill={el['fill_color']}"
        if el.get("text_style"):
            ts = el.get("text_style", {})
            line += f", font={ts.get('font_family','')} {ts.get('font_size_pt','')}pt"
        if el.get("saved_asset"):
            line += f", img={el['saved_asset']}"
        if el.get("children"):
            line += f", children={len(el['children'])}"
        element_lines.append(line)

    # Table info
    for el in elements:
        table_data = el.get("table")
        if table_data:
            element_lines.append(f"  TABLE: {table_data.get('rows',0)} rows x {table_data.get('cols',0)} columns")
            element_lines.append(f"    Header row: {', '.join(c.get('text','') for c in table_data.get('cells',[])[:table_data.get('cols',1)])}")
            break

    element_guide = "\n".join(element_lines) if element_lines else "  (No elements)"

    # ── Design patterns ──
    dp_guide = ""
    if design_patterns:
        dp_guide = "Design patterns:\n" + "\n".join(f"  - {dp}" for dp in design_patterns)

    # ── Element groups ──
    group_guide = ""
    if element_groups:
        group_guide = "Element groups:\n"
        for g in element_groups[:5]:
            group_guide += f"  - [{','.join(str(i) for i in g.get('element_ids',[]))}] role={g.get('group_role','')}: {g.get('description','')}\n"

    # ── Asset info ──
    asset_guide = ""
    if assets_dir:
        asset_guide = f"Reusable assets directory: {assets_dir}/"

    prompt = f"""Generate an SVG slide that faithfully reproduces the reference layout described below.

## Content to Display
- Title: {title}
- Body: {chr(10).join(f'  - {b}' for b in body) if body else 'N/A'}
- Page Type: {page_type}
- Layout: {layout_type}

## Reference Layout Structure
The source PPT slide has these elements arranged in this layout:

{element_guide}

{dp_guide}
{group_guide}
{asset_guide}

## Visual Style (MUST MATCH)
1. **Background**: {bg_description}. Background color: {bg_color}. Use a full-size <rect> for background.
2. **Color Palette** (use ONLY these hex colors):
{color_guide}
3. **Typography**: Use ONLY font-family="{font_family}". Title size 28-36pt, body 14-18pt.
4. **Layout Pattern**: {layout_description or f"Standard {layout_type} layout"}
   Follow the positions and sizes from the Reference Layout Structure above.
5. **NO banned features**: No masks, no rgba(), no @font-face, no <style> tags, no HTML entities.
   Use ONLY hex colors. Use <tspan> for inline text styling. Use raw Unicode.

## Technical Constraints
- SVG viewBox="0 0 1280 720"
- All shapes must be standard SVG elements (rect, circle, path, text, tspan)
- Top-level <g> groups with descriptive IDs matching their role
- Do NOT use <foreignObject>, <textPath>, <script>
- All colors as #HEX, opacity via fill-opacity attribute
- If images are referenced (from assets/), use <image> tags with xlink:href
"""

    # Previous iteration feedback
    if previous_issues:
        issues_text = "\n".join(f"  - {i}" for i in previous_issues)
        prompt += f"""
## CRITICAL Fixes from Previous Attempt
The previous attempt had these style issues:
{issues_text}

FIX ALL of the above style issues. Content can be adjusted slightly if needed.
"""

    return prompt


def generate_slide_with_loop(
    slide_entry: dict[str, Any],
    spec_pages: list[dict],
    spec_metadata: dict[str, Any],
    generate_callback,
) -> SlideResult:
    """Generate one slide using agent-loop evaluation.

    Args:
        slide_entry: Slide content dict (title, body, page_type, layout_type).
        spec_pages: All available spec pages from the loaded spec.
        spec_metadata: Spec metadata (colors, typography, etc.).
        generate_callback: Async function (prompt_str) -> SVG text.
                           Called by the orchestrator (AI runtime).

    Returns:
        SlideResult with best SVG and evaluation report.
    """
    # Determine target page type
    slide_page_type = slide_entry.get("page_type", "content")
    target_layout = slide_entry.get("layout_type", "full_width")

    if slide_entry.get("slide_number", 0) == 1 or slide_page_type == "title":
        slide_page_type = "cover"
    elif slide_entry.get("is_end", False):
        slide_page_type = "end_page"

    # Find matching spec page
    spec_page = _match_spec_page(slide_page_type, target_layout, spec_pages)
    if spec_page is None:
        # Create minimal spec page from metadata
        spec_page = {
            "page_type": slide_page_type,
            "layout_sub_type": target_layout,
            "background_color": spec_metadata.get("colors", {}).get("background1", "#FFFFFF"),
            "colors": spec_metadata.get("colors", {}),
            "typography": spec_metadata.get("typography", {}),
            "regions": [],
            "background_description": "Default background",
        }

    result = SlideResult(
        slide_index=slide_entry.get("slide_number", 0),
        page_type=slide_page_type,
        layout_sub_type=target_layout,
    )

    previous_issues: list[str] | None = None
    best_svg = ""
    best_score = 0.0

    for iteration in range(1, MAX_ITERATIONS + 1):
        # Build prompt
        prompt = _build_generation_prompt(
            slide_entry, spec_page, previous_issues
        )

        # Generate SVG (via callback — AI runtime handles this)
        try:
            svg_text = generate_callback(prompt)
        except Exception as e:
            result.reports.append(StyleReport(issues=[f"Generation failed: {e}"]))
            continue

        if not svg_text or "<svg" not in svg_text.lower():
            result.reports.append(StyleReport(issues=["Generated text is not valid SVG"]))
            continue

        # Evaluate against spec
        report = evaluate_style(svg_text, spec_page)
        result.reports.append(report)

        # Track best
        if report.overall > best_score:
            best_score = report.overall
            best_svg = svg_text

        if report.passed:
            result.svg_text = svg_text
            result.iterations = iteration
            result.best_score = report.overall
            result.passed = True
            return result

        # Prepare fix instructions for next iteration
        previous_issues = report.issues
        if iteration < MAX_ITERATIONS:
            previous_issues.append(
                f"Overall style score: {report.overall:.0%}, "
                f"target: {PASS_THRESHOLD:.0%}. "
                f"Color={report.color_score:.0%}, Font={report.typography_score:.0%}, "
                f"Layout={report.layout_score:.0%}, BG={report.background_score:.0%}"
            )

    # Return best attempt
    result.svg_text = best_svg
    result.iterations = MAX_ITERATIONS
    result.best_score = best_score
    result.passed = best_score >= PASS_THRESHOLD
    return result
