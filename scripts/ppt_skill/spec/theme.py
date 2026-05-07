"""Theme-level extraction from PPTX files — colors, fonts, and backgrounds.

All extraction uses direct lxml/XML parsing to bypass python-pptx limitations:
- python-pptx's font/color APIs return None for inherited (theme-defaulted) values.
- python-pptx's Slide.background._element returns <p:cSld> instead of <p:bg>
  (known bug #1126). We find <p:bg> via lxml on the parent cSld element.

OOXML namespace URIs (stable, from ISO/IEC 29500):
  a = http://schemas.openxmlformats.org/drawingml/2006/main
  p = http://schemas.openxmlformats.org/presentationml/2006/main
"""

from __future__ import annotations

import zipfile
from typing import Any

from lxml import etree

# OOXML namespace URIs (stable per ISO/IEC 29500)
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Namespace map for lxml xpath/find queries
NSMAP = {"a": A_NS, "p": P_NS, "r": R_NS}


# ---------------------------------------------------------------------------
# Color resolution helpers
# ---------------------------------------------------------------------------


def _resolve_color_element(color_elem: etree._Element) -> str | None:
    """Resolve a single OOXML color element to a HEX string (with '#' prefix).

    Handles srgbClr, sysClr (with lastClr fallback), and schemeClr.
    Returns None if the element has no resolvable color.

    Args:
        color_elem: The inner color element (e.g., <a:srgbClr val="4472C4"/>).

    Returns:
        HEX string like "#4472C4" or None.
    """
    if color_elem is None:
        return None

    tag = etree.QName(color_elem.tag).localname

    if tag == "srgbClr":
        val = color_elem.get("val")
        return f"#{val.upper()}" if val else None

    if tag == "sysClr":
        # Prefer lastClr (explicit fallback) over val (system semantic name)
        val = color_elem.get("lastClr") or color_elem.get("val")
        return f"#{val.upper()}" if val else None

    if tag == "schemeClr":
        # schemeClr references another theme slot — caller must resolve
        ref_name = color_elem.get("val")
        return f"@scheme:{ref_name}" if ref_name else None

    return None


def _resolve_color_with_map(
    color_elem: etree._Element, scheme_map: dict[str, str]
) -> str | None:
    """Resolve a color element, using scheme_map for schemeClr references.

    If the element is a schemeClr, looks up the reference name in scheme_map.
    schemeClr references not found in the map return None.
    """
    if color_elem is None:
        return None

    tag = etree.QName(color_elem.tag).localname

    if tag == "schemeClr":
        ref_name = color_elem.get("val")
        if ref_name and ref_name in scheme_map:
            return scheme_map[ref_name]
        # Could not resolve — return None rather than a placeholder
        return None

    return _resolve_color_element(color_elem)


# ---------------------------------------------------------------------------
# Color extraction
# ---------------------------------------------------------------------------


def extract_theme_colors(pptx_path: str) -> dict[str, str]:
    """Extract resolved 12-color palette from the PPTX theme's clrScheme.

    Opens the PPTX as a zipfile, reads ppt/theme/theme1.xml, and parses the
    <a:clrScheme> element. Each child (dk1, dk2, lt1, lt2, accent1–6,
    hlink, folHlink) is resolved to a concrete HEX color.

    Resolution chain (per the OOXML spec):
    1. srgbClr → direct HEX value (prefixed with '#')
    2. sysClr → lastClr attribute (system color explicit fallback)
    3. schemeClr → resolved against clrScheme's own entries (two-pass)

    Args:
        pptx_path: Path to a .pptx file.

    Returns:
        Dict mapping OOXML scheme names to HEX strings, e.g.:
        {"dk1": "#000000", "accent1": "#4472C4", ...}
        Missing colors are omitted from the dict.
    """
    with zipfile.ZipFile(pptx_path, "r") as zf:
        theme_xml = zf.read("ppt/theme/theme1.xml")

    root = etree.fromstring(theme_xml)
    clr_scheme = root.find(f"{{{A_NS}}}themeElements/{{{A_NS}}}clrScheme")

    if clr_scheme is None:
        return {}

    # --- Pass 1: extract direct colors (srgbClr, sysClr) ---
    colors: dict[str, str] = {}
    deferred: dict[str, etree._Element] = {}

    for child in clr_scheme:
        name = etree.QName(child.tag).localname  # dk1, lt1, accent1, etc.
        inner = list(child)[0] if list(child) else None

        if inner is None:
            continue

        inner_tag = etree.QName(inner.tag).localname

        if inner_tag == "schemeClr":
            # Defer — needs resolution against other entries
            deferred[name] = inner
        else:
            resolved = _resolve_color_element(inner)
            if resolved:
                colors[name] = resolved

    # --- Pass 2: resolve schemeClr references against Pass 1 results ---
    for name, inner in deferred.items():
        resolved = _resolve_color_with_map(inner, colors)
        if resolved:
            colors[name] = resolved

    return colors


# ---------------------------------------------------------------------------
# Font extraction
# ---------------------------------------------------------------------------


def extract_theme_fonts(pptx_path: str) -> dict[str, str]:
    """Extract heading and body font families from the theme's fontScheme.

    Reads <a:fontScheme> from theme1.xml and extracts the latin typeface
    from <a:majorFont> and <a:minorFont>.

    Args:
        pptx_path: Path to a .pptx file.

    Returns:
        Dict with keys "majorFont" and "minorFont" mapping to typeface
        names (e.g., {"majorFont": "Calibri", "minorFont": "Calibri"}).
        Missing fonts are omitted.
    """
    with zipfile.ZipFile(pptx_path, "r") as zf:
        theme_xml = zf.read("ppt/theme/theme1.xml")

    root = etree.fromstring(theme_xml)
    font_scheme = root.find(f"{{{A_NS}}}themeElements/{{{A_NS}}}fontScheme")

    if font_scheme is None:
        return {}

    fonts: dict[str, str] = {}

    for font_class in ("majorFont", "minorFont"):
        elem = font_scheme.find(f"{{{A_NS}}}{font_class}")
        if elem is None:
            continue
        latin = elem.find(f"{{{A_NS}}}latin")
        if latin is not None:
            typeface = latin.get("typeface")
            if typeface:
                fonts[font_class] = typeface

    return fonts


# ---------------------------------------------------------------------------
# Background extraction (with python-pptx bug #1126 workaround)
# ---------------------------------------------------------------------------


def _find_bg_from_element(parent_elem: etree._Element) -> etree._Element | None:
    """Find the <p:bg> element within a parent (cSld, sldLayout, sldMaster)."""
    return parent_elem.find(f"{{{P_NS}}}bg")


def _parse_bg_fill(bg_elem: etree._Element) -> dict[str, Any] | None:
    """Extract fill information from a <p:bg> element.

    Handles solid fills (srgbClr, theme refs via bgRef/schemeClr)
    and gradient fills (gradFill with stop list).

    Returns:
        Dict with type and color data, or None if no fill found.
    """
    bg_pr = bg_elem.find(f"{{{P_NS}}}bgPr")
    if bg_pr is None:
        return None

    # --- Check for gradient fill ---
    grad_fill = bg_pr.find(f"{{{A_NS}}}gradFill")
    if grad_fill is not None:
        stops: list[dict[str, Any]] = []
        gs_lst = grad_fill.find(f"{{{A_NS}}}gsLst")
        if gs_lst is not None:
            for gs in gs_lst.findall(f"{{{A_NS}}}gs"):
                pos_val = gs.get("pos")
                pos = int(pos_val) / 100000.0 if pos_val else 0.0
                # Get the color element inside the stop
                color_child = list(gs)[0] if list(gs) else None
                color = _resolve_color_element(color_child) if color_child is not None else None
                stops.append({"position": pos, "color": color})

        path = grad_fill.get("path", "circle")
        return {
            "type": "gradient",
            "path_type": path,
            "stops": stops,
        }

    # --- Check for solid fill ---
    solid = bg_pr.find(f"{{{A_NS}}}solidFill")
    if solid is not None:
        color_elem = list(solid)[0] if list(solid) else None
        if color_elem is not None:
            converted = _resolve_color_element(color_elem)
            if converted:
                return {"type": "solid", "color": converted}

    # --- Check for theme reference (bgRef) ---
    bg_ref = bg_pr.find(f"{{{A_NS}}}bgRef")
    if bg_ref is not None:
        scheme = bg_ref.find(f"{{{A_NS}}}schemeClr")
        if scheme is not None:
            return {"type": "theme_ref", "ref": scheme.get("val", "?")}
        srgb = bg_ref.find(f"{{{A_NS}}}srgbClr")
        if srgb is not None:
            val = srgb.get("val")
            if val:
                return {"type": "theme_ref", "color": f"#{val.upper()}"}

    return None


def _find_bg_in_theme_fill_styles(
    pptx_path: str,
) -> dict[str, Any] | None:
    """Extract background from theme's bgFillStyleLst as final fallback."""
    try:
        with zipfile.ZipFile(pptx_path, "r") as zf:
            theme_xml = zf.read("ppt/theme/theme1.xml")
    except (KeyError, zipfile.BadZipFile):
        return None

    root = etree.fromstring(theme_xml)
    fmt_scheme = root.find(f"{{{A_NS}}}themeElements/{{{A_NS}}}fmtScheme")
    if fmt_scheme is None:
        return None

    fill_style_lst = fmt_scheme.find(f"{{{A_NS}}}bgFillStyleLst")
    if fill_style_lst is None:
        return None

    # bgFillStyleLst contains <a:solidFill>, <a:gradFill> etc. as children
    # Take the first fill style (index 0 = default background)
    for child in fill_style_lst:
        tag = etree.QName(child.tag).localname

        if tag == "solidFill":
            color_elem = list(child)[0] if list(child) else None
            if color_elem is not None:
                converted = _resolve_color_element(color_elem)
                if converted:
                    return {
                        "type": "solid",
                        "color": converted,
                        "inherited_from": "theme",
                    }
        elif tag == "gradFill":
            stops: list[dict[str, Any]] = []
            gs_lst = child.find(f"{{{A_NS}}}gsLst")
            if gs_lst is not None:
                for gs in gs_lst.findall(f"{{{A_NS}}}gs"):
                    pos_val = gs.get("pos")
                    pos = int(pos_val) / 100000.0 if pos_val else 0.0
                    color_child = list(gs)[0] if list(gs) else None
                    color = _resolve_color_element(color_child) if color_child is not None else None
                    stops.append({"position": pos, "color": color})

            path = child.get("path", "circle")
            return {
                "type": "gradient",
                "path_type": path,
                "stops": stops,
                "inherited_from": "theme",
            }

        # Only return the first valid fill
        break

    return None


def extract_slide_background(
    slide,  # python-pptx Slide object
    pptx_path: str = "",
) -> dict[str, Any] | None:
    """Safely extract slide background avoiding python-pptx bug #1126.

    Walks the background inheritance chain:
      1. Slide-level <p:bg> (found via cSld parent)
      2. Slide layout <p:bg>
      3. Slide master <p:bg>
      4. Theme <bgFillStyleLst> (ultimate fallback)

    Uses lxml.find() on the parent elements — never iterates children of
    slide.background._element (which returns cSld, not bg, per bug #1126).

    Args:
        slide: A python-pptx Slide object.
        pptx_path: Path to the .pptx file (needed for theme fallback).

    Returns:
        Background dict (e.g., {"type": "solid", "color": "#FFFFFF"}) or
        None if no background found at any level.
    """
    # --- Level 1: Slide-level background ---
    # slide.background._element returns <p:cSld> due to bug #1126
    c_sld = slide.background._element
    bg_elem = _find_bg_from_element(c_sld)

    if bg_elem is not None:
        result = _parse_bg_fill(bg_elem)
        if result is not None:
            result.setdefault("inherited_from", "slide")
            return result

    # --- Level 2: Slide layout background ---
    try:
        layout = slide.slide_layout
        layout_root = layout._element  # <p:sldLayout>
        bg_elem = _find_bg_from_element(layout_root)
        if bg_elem is not None:
            result = _parse_bg_fill(bg_elem)
            if result is not None:
                result.setdefault("inherited_from", "layout")
                return result
    except (AttributeError, Exception):
        pass

    # --- Level 3: Slide master background ---
    try:
        master = slide.slide_layout.slide_master
        master_root = master._element  # <p:sldMaster>
        bg_elem = _find_bg_from_element(master_root)
        if bg_elem is not None:
            result = _parse_bg_fill(bg_elem)
            if result is not None:
                result.setdefault("inherited_from", "master")
                return result
    except (AttributeError, Exception):
        pass

    # --- Level 4: Theme bgFillStyleLst (final fallback) ---
    if pptx_path:
        result = _find_bg_in_theme_fill_styles(pptx_path)
        if result is not None:
            return result

    return None


__all__ = [
    "extract_theme_colors",
    "extract_theme_fonts",
    "extract_slide_background",
]
