"""Style similarity evaluator — compare generated slide against spec reference.

Measures style fidelity (NOT content) across 5 dimensions:
  1. Color palette  — hex color usage match against spec palette
  2. Typography     — font family + size hierarchy match
  3. Layout         — region position overlap (IoU-based)
  4. Background     — background color/presence match
  5. Density        — content density alignment (breathing/dense/anchor)

Composite score: weighted average, 0.0–1.0. Threshold: >= 0.90 passes.
"""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field
from typing import Any


@dataclass
class StyleReport:
    """Detailed style evaluation report for one slide."""
    overall: float = 0.0
    color_score: float = 0.0
    typography_score: float = 0.0
    layout_score: float = 0.0
    background_score: float = 0.0
    density_score: float = 0.0
    issues: list[str] = field(default_factory=list)
    passed: bool = False

    # Weights per dimension
    WEIGHTS = {
        "color": 0.30,
        "typography": 0.20,
        "layout": 0.30,
        "background": 0.15,
        "density": 0.05,
    }


# ── Color extraction from SVG ────────────────────────────────────────

def _extract_svg_colors(svg_text: str) -> Counter:
    """Extract all hex color values from SVG text."""
    colors = re.findall(r'#([0-9a-fA-F]{6})', svg_text)
    return Counter(c.lower() for c in colors)


def _hex_distance(c1: str, c2: str) -> float:
    """Perceptual color distance (0.0–1.0, lower is more similar)."""
    r1, g1, b1 = int(c1[0:2], 16), int(c1[2:4], 16), int(c1[4:6], 16)
    r2, g2, b2 = int(c2[0:2], 16), int(c2[2:4], 16), int(c2[4:6], 16)
    # Weighted Euclidean distance in RGB (simplified perceptual)
    dr = (r1 - r2) / 255.0
    dg = (g1 - g2) / 255.0
    db = (b1 - b2) / 255.0
    return (dr * dr * 0.299 + dg * dg * 0.587 + db * db * 0.114) ** 0.5


def evaluate_colors(svg_text: str, spec_palette: dict[str, str]) -> tuple[float, list[str]]:
    """Compare SVG color usage against spec color palette."""
    svg_colors = _extract_svg_colors(svg_text)
    if not svg_colors:
        return 0.0, ["No hex colors found in SVG"]

    spec_colors = {v.strip("#").lower() for v in spec_palette.values() if v}

    if not spec_colors:
        return 1.0, []  # No spec palette to compare against

    # Count how many of the dominant SVG colors are close to spec colors
    total_usage = sum(svg_colors.values())
    matched_usage = 0
    issues: list[str] = []

    for color_hex, count in svg_colors.most_common(10):
        # Check if this color matches any spec color (within tolerance)
        best_dist = min((_hex_distance(color_hex, sc) for sc in spec_colors), default=1.0)
        if best_dist < 0.15:  # 15% perceptual distance threshold
            matched_usage += count
        elif count / total_usage > 0.05:  # Significant color not in palette
            issues.append(f"Color #{color_hex} ({count/total_usage:.0%}) not in spec palette")

    score = matched_usage / total_usage if total_usage else 0.0
    return score, issues


# ── Typography extraction from SVG ────────────────────────────────────

def _extract_svg_fonts(svg_text: str) -> dict[str, Any]:
    """Extract font families and sizes from SVG text elements."""
    font_families = re.findall(r'font-family="([^"]*)"', svg_text)
    font_sizes = re.findall(r'font-size="([^"]*)"', svg_text)

    sizes = []
    for s in font_sizes:
        try:
            sizes.append(float(s.replace("px", "")))
        except ValueError:
            pass

    return {
        "families": Counter(f.split(",")[0].strip().strip("'\"") for f in font_families),
        "sizes": sorted(sizes),
        "size_count": len(sizes),
    }


def evaluate_typography(svg_text: str, spec_typography: dict[str, Any]) -> tuple[float, list[str]]:
    """Compare SVG font usage against spec typography."""
    svg_fonts = _extract_svg_fonts(svg_text)
    issues: list[str] = []

    if not svg_fonts["families"]:
        return 0.0, ["No fonts found in SVG"]

    spec_family = spec_typography.get("heading_family", "") or spec_typography.get("body_family", "")
    if not spec_family:
        return 1.0, []

    # Font family match
    family_score = 0.0
    for family, count in svg_fonts["families"].most_common(3):
        family_lower = family.lower()
        spec_lower = spec_family.lower()
        if family_lower == spec_lower or spec_lower in family_lower or family_lower in spec_lower:
            family_score = 1.0
            break
        elif any(w in family_lower for w in spec_lower.split()):
            family_score = 0.5

    if family_score < 0.5:
        issues.append(f"Expected font '{spec_family}', got '{', '.join(svg_fonts['families'])}'")

    # Font size check: at least 2 distinct sizes (hierarchy)
    size_score = 0.5
    sizes = svg_fonts["sizes"]
    if len(set(sizes)) >= 2:
        size_score = 1.0
    elif len(sizes) >= 1:
        size_score = 0.7

    score = family_score * 0.7 + size_score * 0.3
    return score, issues


# ── Layout comparison ────────────────────────────────────────────────

def _extract_svg_rects(svg_text: str) -> list[dict]:
    """Extract rectangle regions from SVG (shapes, text bounding boxes)."""
    # Match rect, text x/y, and foreignObject elements
    rects = re.findall(
        r'<(?:rect|text|image|foreignObject)[^>]*?'
        r'(?:x|y|width|height)="([^"]*)"[^>]*?'
        r'(?:x|y|width|height)="([^"]*)"[^>]*?'
        r'(?:x|y|width|height)="([^"]*)"[^>]*?'
        r'(?:x|y|width|height)="([^"]*)"',
        svg_text,
    )
    return rects


def evaluate_layout(svg_text: str, spec_regions: list[dict],
                    svg_width: int = 1280, svg_height: int = 720) -> tuple[float, list[str]]:
    """Compare SVG layout against spec region positions (IoU-based)."""
    if not spec_regions:
        return 1.0, []

    issues: list[str] = []
    matched_regions = 0
    total_score = 0.0

    for spec_region in spec_regions:
        # Spec region in normalized coords (0–1)
        sx = spec_region.get("x", 0)
        sy = spec_region.get("y", 0)
        sw = spec_region.get("width", 0)
        sh = spec_region.get("height", 0)
        role = spec_region.get("role", "")

        if sw <= 0 or sh <= 0:
            continue

        # Search SVG text for matching role/region
        # Simple heuristic: check if there's a rect/text near the expected position
        region_score = 0.0

        # Look for elements near the expected position
        expected_x = sx * svg_width
        expected_y = sy * svg_height

        # Extract positions from SVG (simplified: use re match on x/y attrs)
        svg_x_matches = re.findall(r'x="([^"]*)"', svg_text)
        svg_y_matches = re.findall(r'y="([^"]*)"', svg_text)

        position_matches = []
        for xm in svg_x_matches:
            for ym in svg_y_matches:
                try:
                    px, py = float(xm), float(ym)
                    if abs(px - expected_x) < sw * svg_width * 0.3 and abs(py - expected_y) < sh * svg_height * 0.3:
                        position_matches.append(abs(px - expected_x) + abs(py - expected_y))
                except ValueError:
                    pass

        if position_matches:
            matched_regions += 1
            # Closer = higher score
            best = min(position_matches)
            max_dist = (sw * svg_width + sh * svg_height) * 0.3
            region_score = 1.0 - (best / max_dist if max_dist else 1.0)
            region_score = max(0.0, region_score)
        else:
            issues.append(f"Region '{role}' not found near expected position")

        total_score += region_score

    # Overall layout score
    region_count = len(spec_regions)
    pos_score = total_score / region_count if region_count else 1.0
    coverage_score = matched_regions / region_count if region_count else 1.0

    score = pos_score * 0.6 + coverage_score * 0.4
    return max(0.0, min(1.0, score)), issues


# ── Background comparison ────────────────────────────────────────────

def evaluate_background(svg_text: str, spec_bg_color: str,
                        spec_has_image: bool = False) -> tuple[float, list[str]]:
    """Compare SVG background against spec background."""
    issues: list[str] = []

    if not spec_bg_color:
        return 1.0, []

    # Find background rect (first large rect usually)
    bg_rects = re.findall(
        r'<rect[^>]*?fill="([^"]*)"[^>]*?(?:width="(\d+)[^"]*?"[^>]*?height="(\d+)[^"]*?"|height="(\d+)[^"]*?"[^>]*?width="(\d+)[^"]*?")',
        svg_text,
    )
    if not bg_rects:
        return 0.0, ["No background rect found"]

    bg_fill = bg_rects[0][0] if isinstance(bg_rects[0], tuple) else bg_rects[0]
    if bg_fill.startswith("#"):
        bg_hex = bg_fill.lstrip("#").lower()
        spec_hex = spec_bg_color.lstrip("#").lower()
        distance = _hex_distance(bg_hex, spec_hex)
        score = max(0.0, 1.0 - distance * 3)  # More lenient on background
    elif bg_fill.startswith("url("):
        score = 1.0 if spec_has_image else 0.5
    else:
        score = 0.5  # Named color, hard to compare

    if score < 0.5:
        issues.append(f"Background color mismatch: expected #{spec_bg_color}")

    return score, issues


# ── Density check ────────────────────────────────────────────────────

def evaluate_density(svg_text: str,
                     target_density: str = "dense") -> tuple[float, list[str]]:
    """Check if SVG content density matches target."""
    char_count = len(re.sub(r'<[^>]+>', '', svg_text).strip())
    issues: list[str] = []

    if target_density == "breathing":
        expected = (0, 300)
    elif target_density == "anchor":
        expected = (500, float("inf"))
    else:  # dense
        expected = (100, 800)

    if expected[0] <= char_count <= expected[1]:
        score = 1.0
    elif char_count < expected[0]:
        score = char_count / expected[0] if expected[0] else 0.0
        issues.append(f"Content too sparse ({char_count} chars, expected {expected[0]}+)")
    else:
        ratio = expected[1] / char_count if char_count else 0.0
        score = max(0.3, ratio)
        issues.append(f"Content too dense ({char_count} chars, expected <{expected[1]})")

    return score, issues


# ── Composite evaluator ──────────────────────────────────────────────


def evaluate_style(
    svg_text: str,
    spec_page: dict[str, Any],
) -> StyleReport:
    """Evaluate generated SVG style against a spec page reference.

    Args:
        svg_text: Generated SVG content as string.
        spec_page: Spec page dict with keys:
            - page_type, layout_sub_type
            - background_color, background_description
            - regions (list of region dicts)
            - colors: spec color palette dict
            - typography: font dict

    Returns:
        StyleReport with per-dimension scores and overall pass/fail.
    """
    report = StyleReport()
    colors_palette = spec_page.get("colors", {})

    # 1. Color evaluation
    report.color_score, ci = evaluate_colors(svg_text, colors_palette)
    report.issues.extend(f"[color] {i}" for i in ci)

    # 2. Typography evaluation
    report.typography_score, ti = evaluate_typography(
        svg_text, spec_page.get("typography", {})
    )
    report.issues.extend(f"[typography] {i}" for i in ti)

    # 3. Layout evaluation
    regions = spec_page.get("regions", [])
    if regions:
        report.layout_score, li = evaluate_layout(svg_text, regions)
        report.issues.extend(f"[layout] {i}" for i in li)
    else:
        report.layout_score = 1.0  # No layout spec to compare

    # 4. Background evaluation
    bg_color = spec_page.get("background_color", "")
    bg_has_image = spec_page.get("has_background_image", False)
    report.background_score, bi = evaluate_background(svg_text, bg_color, bg_has_image)
    report.issues.extend(f"[background] {i}" for i in bi)

    # 5. Density evaluation
    target_density = spec_page.get("target_density", "dense")
    report.density_score, di = evaluate_density(svg_text, target_density)
    report.issues.extend(f"[density] {i}" for i in di)

    # Composite score
    w = StyleReport.WEIGHTS
    report.overall = (
        report.color_score * w["color"]
        + report.typography_score * w["typography"]
        + report.layout_score * w["layout"]
        + report.background_score * w["background"]
        + report.density_score * w["density"]
    )
    report.passed = report.overall >= 0.90

    return report
