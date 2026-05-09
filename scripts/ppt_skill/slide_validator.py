"""Pre-flight slide validation — catch layout bugs before VL evaluation.

Checks every text element on a slide for:
  1. Boundary overflow (text extends beyond slide edges)
  2. Minimum margins (text too close to edges)  
  3. Text box width issues (zero-width causing vertical text)
  4. Text box overlap detection
  5. Readability (font size > 0, font name set)

Returns list of Violation objects. Empty list = clean slide.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class Violation:
    severity: str        # "BLOCKER", "WARNING", "INFO"
    code: str            # e.g. "OVERFLOW_RIGHT", "ZERO_WIDTH", "OVERLAP"
    element_name: str    # shape name or text content snippet
    description: str
    fix_hint: str


# ── Thresholds ───────────────────────────────────────────────────────

MIN_MARGIN_INCHES = 0.1     # text should be at least 0.1in from edge
MIN_TEXT_WIDTH_INCHES = 0.5  # text box should be at least 0.5in wide
MAX_OVERLAP_INCHES = 0.02   # overlap tolerance


def validate_slide(slide, slide_width_emu: int, slide_height_emu: int) -> list[Violation]:
    """Check a PPTX slide for common layout issues.

    Args:
        slide: python-pptx Slide object
        slide_width_emu: Slide width in EMU
        slide_height_emu: Slide height in EMU

    Returns:
        List of Violation objects. Empty if slide passes.
    """
    violations: list[Violation] = []
    sw_in = slide_width_emu / 914400
    sh_in = slide_height_emu / 914400

    text_boxes: list[dict[str, Any]] = []

    for shape in slide.shapes:
        if not hasattr(shape, 'text_frame') or not shape.text_frame.text.strip():
            continue

        left = (shape.left or 0) / 914400
        top = (shape.top or 0) / 914400
        width = (shape.width or 0) / 914400
        height = (shape.height or 0) / 914400
        right = left + width
        bottom = top + height
        name = shape.name
        text = shape.text_frame.text.strip()[:40]

        text_boxes.append({"left": left, "top": top, "right": right, "bottom": bottom,
                           "name": name, "text": text, "width": width})

        # ── Check 1: Boundary overflow ──
        if left < -0.01:
            violations.append(Violation("BLOCKER", "OVERFLOW_LEFT", name,
                f"Text box extends {abs(left):.2f}in beyond left edge",
                f"Move right by at least {abs(left):.2f}in"))
        if right > sw_in + 0.01:
            violations.append(Violation("BLOCKER", "OVERFLOW_RIGHT", name,
                f"\"{text}\" extends {right - sw_in:.2f}in beyond right edge",
                f"Reduce width or move left by {right - sw_in:.2f}in"))
        if top < -0.01:
            violations.append(Violation("BLOCKER", "OVERFLOW_TOP", name,
                f"Text box extends {abs(top):.2f}in above top edge",
                f"Move down by at least {abs(top):.2f}in"))
        if bottom > sh_in + 0.05:
            violations.append(Violation("WARNING", "OVERFLOW_BOTTOM", name,
                f"\"{text}\" extends {bottom - sh_in:.2f}in below slide",
                f"Move up or reduce height"))

        # ── Check 2: Margin violations ──
        if top < MIN_MARGIN_INCHES and top >= 0:
            violations.append(Violation("INFO", "TIGHT_TOP", name,
                f"\"{text}\" is only {top:.2f}in from top edge",
                f"Consider moving down to {MIN_MARGIN_INCHES}in minimum"))
        if left < MIN_MARGIN_INCHES and left >= 0:
            violations.append(Violation("INFO", "TIGHT_LEFT", name,
                f"\"{text}\" is only {left:.2f}in from left edge",
                f"Consider moving right to {MIN_MARGIN_INCHES}in minimum"))

        # ── Check 3: Width issues (vertical text cause) ──
        if width < MIN_TEXT_WIDTH_INCHES:
            violations.append(Violation("BLOCKER", "ZERO_WIDTH", name,
                f"\"{text}\" has width {width:.2f}in — will cause vertical/stacked text",
                f"Set text box width to at least {MIN_TEXT_WIDTH_INCHES}in"))

        # ── Check 4: Font presence ──
        para = shape.text_frame.paragraphs[0]
        if para.runs:
            run = para.runs[0]
            if not run.font.name:
                violations.append(Violation("WARNING", "NO_FONT", name,
                    f"\"{text}\" has no font family set",
                    "Set run.font.name (use run.add_run() and set font on the run)"))
            if not run.font.size or run.font.size == 0:
                violations.append(Violation("WARNING", "NO_FONT_SIZE", name,
                    f"\"{text}\" has zero font size",
                    "Set run.font.size = Pt(sz)"))

    # ── Check 5: Text box overlaps ──
    for i in range(len(text_boxes)):
        for j in range(i + 1, len(text_boxes)):
            a, b = text_boxes[i], text_boxes[j]
            # Check if boxes overlap vertically AND horizontally
            v_overlap = a["bottom"] > b["top"] + MAX_OVERLAP_INCHES and a["top"] < b["bottom"] - MAX_OVERLAP_INCHES
            h_overlap = a["right"] > b["left"] and a["left"] < b["right"]
            if v_overlap and h_overlap:
                # Don't flag table cells (same row)
                if abs(a["top"] - b["top"]) < 0.05 and a["left"] < b["left"] + 2 and a["left"] > b["left"] - 2:
                    continue
                violations.append(Violation("WARNING", "OVERLAP",
                    f"{a['name']} vs {b['name']}",
                    f"\"{a['text']}\" overlaps \"{b['text']}\" (v={a['top']:.1f}-{a['bottom']:.1f} vs {b['top']:.1f}-{b['bottom']:.1f})",
                    f"Increase spacing between these elements"))

    return violations


def validate_presentation(prs) -> dict[int, list[Violation]]:
    """Validate all slides in a presentation.

    Returns dict mapping slide_index → list of Violations.
    """
    results: dict[int, list[Violation]] = {}
    slides = list(prs.slides)
    sw = prs.slide_width
    sh = prs.slide_height
    for i, slide in enumerate(slides):
        v = validate_slide(slide, sw, sh)
        if v:
            results[i] = v
    return results


def report(prs) -> str:
    """Generate a human-readable validation report."""
    results = validate_presentation(prs)
    blockers = 0; warnings = 0; infos = 0

    lines = ["=" * 60, "SLIDE VALIDATION REPORT", "=" * 60]
    for si, violations in sorted(results.items()):
        lines.append(f"\nSlide {si}: {len(violations)} issue(s)")
        for v in violations:
            lines.append(f"  [{v.severity:7s}] {v.code}: {v.description}")
            if v.fix_hint:
                lines.append(f"           Fix: {v.fix_hint}")
            if v.severity == "BLOCKER": blockers += 1
            elif v.severity == "WARNING": warnings += 1
            else: infos += 1

    lines.extend([
        "\n" + "=" * 60,
        f"SUMMARY: {blockers} blockers, {warnings} warnings, {infos} info",
        "READY" if blockers == 0 else f"BLOCKED ({blockers} issues must be fixed)",
        "=" * 60,
    ])
    return "\n".join(lines)
