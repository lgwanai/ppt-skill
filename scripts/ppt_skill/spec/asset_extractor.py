"""Asset extraction — backgrounds, images, and reusable visual elements from PPTX.

Extracts:
  1. Background images from slide / layout / master levels (blipFill with embed)
  2. Picture shapes (png, jpg, emf, etc.)
  3. Shape fill images (shapes with picture fills)
  4. Background solid/gradient fills with hex colors

All images saved to assets/ directory with descriptive filenames.
"""

from __future__ import annotations

import io
import re
from pathlib import Path
from typing import Any

from lxml import etree

# OOXML namespaces
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


# ── Background image extraction ──────────────────────────────────────


def extract_background_image(slide, spec_dir: Path,
                              page_idx: int) -> dict[str, Any]:
    """Extract slide background image or color from inheritance chain.

    Walks: slide → slide layout → slide master → theme.

    Returns:
        {
            "type": "solid" | "image" | "gradient" | "none",
            "color": "#HEX" | "",
            "image": "assets/bg_000.png" | "",
            "description": "..." 
        }
    """
    result: dict[str, Any] = {
        "type": "solid",
        "color": "#FFFFFF",
        "image": "",
        "description": "White background",
    }

    # Level 1: Slide-level <p:bg>
    try:
        c_sld = slide.background._element
        bg = _find_bg(c_sld)
        if bg is not None:
            r = _parse_bg(bg, slide, spec_dir, page_idx, "slide")
            if r:
                return r
    except Exception:
        pass

    # Level 2: Slide layout
    try:
        layout = slide.slide_layout
        bg = _find_bg(layout._element)
        if bg is not None:
            r = _parse_bg(bg, slide, spec_dir, page_idx, "layout")
            if r:
                return r
    except Exception:
        pass

    # Level 3: Slide master
    try:
        master = slide.slide_layout.slide_master
        bg = _find_bg(master._element)
        if bg is not None:
            r = _parse_bg(bg, slide, spec_dir, page_idx, "master")
            if r:
                return r
    except Exception:
        pass

    return result


def _find_bg(parent: etree._Element) -> etree._Element | None:
    """Find <p:bg> element in a parent XML element."""
    return parent.find(f"{{{P_NS}}}bg")


def _parse_bg(bg_elem: etree._Element, slide, spec_dir: Path,
              page_idx: int, source: str) -> dict[str, Any] | None:
    """Parse a <p:bg> element."""
    bg_pr = bg_elem.find(f"{{{P_NS}}}bgPr")
    if bg_pr is None:
        return None

    result: dict[str, Any] = {}

    # Image fill (blipFill)
    blip_fill = bg_pr.find(f"{{{A_NS}}}blipFill")
    if blip_fill is not None:
        img_path = _extract_blip(blip_fill, slide, spec_dir, f"bg_{page_idx:03d}")
        if img_path:
            return {
                "type": "image",
                "color": "",
                "image": img_path,
                "description": f"Background image (from {source})",
            }

    # Gradient fill
    grad_fill = bg_pr.find(f"{{{A_NS}}}gradFill")
    if grad_fill is not None:
        stops = []
        gs_lst = grad_fill.find(f"{{{A_NS}}}gsLst")
        if gs_lst is not None:
            for gs in gs_lst.findall(f"{{{A_NS}}}gs"):
                pos = int(gs.get("pos", 0)) / 100000.0
                color = _resolve_color(gs[0]) if list(gs) else None
                stops.append({"pos": round(pos, 2), "color": color})
        return {
            "type": "gradient",
            "color": "",
            "image": "",
            "description": f"Gradient background ({len(stops)} stops, from {source})",
            "gradient_stops": stops,
        }

    # Solid fill
    solid = bg_pr.find(f"{{{A_NS}}}solidFill")
    if solid is not None and list(solid):
        color = _resolve_color(list(solid)[0])
        if color and color.startswith("#"):
            return {
                "type": "solid",
                "color": color,
                "image": "",
                "description": f"Solid background {color} (from {source})",
            }

    return None


def _extract_blip(blip_fill: etree._Element, slide, spec_dir: Path,
                  base_name: str) -> str | None:
    """Extract image from a blipFill element to assets/ directory."""
    blip = blip_fill.find(f"{{{A_NS}}}blip")
    if blip is None:
        return None

    embed_id = blip.get(f"{{{R_NS}}}embed")
    if not embed_id:
        return None

    try:
        rel = slide.part.rels[embed_id]
        image = rel.target_part
        blob = image.blob
        content_type = image.content_type

        ext = _content_type_to_ext(content_type)
        assets_dir = spec_dir / "assets"
        assets_dir.mkdir(parents=True, exist_ok=True)

        filename = f"{base_name}.{ext}"
        filepath = assets_dir / filename
        filepath.write_bytes(blob)
        return f"assets/{filename}"
    except Exception:
        return None


# ── Image extraction from shapes ─────────────────────────────────────


def extract_shape_images(shape, spec_dir: Path,
                          page_idx: int) -> list[dict[str, Any]]:
    """Extract all images from a shape (Picture, image fill, group children).

    Returns list of {type, src, width_px, height_px, position, description}.
    """
    results: list[dict] = []

    # Direct Picture shapes
    try:
        from pptx.shapes.picture import Picture
        if isinstance(shape, Picture):
            result = _save_picture(shape, spec_dir, page_idx, len(results))
            if result:
                results.append(result)
            return results
    except ImportError:
        pass

    # Shapes with image fill
    try:
        if hasattr(shape, "fill"):
            from pptx.enum.dml import MSO_FILL_TYPE
            if shape.fill.type == MSO_FILL_TYPE.PICTURE:
                result = _extract_fill_image(shape, spec_dir, page_idx, len(results))
                if result:
                    results.append(result)
    except Exception:
        pass

    # Group shapes — recurse into children
    try:
        if hasattr(shape, "shapes"):
            for child in shape.shapes:
                child_results = extract_shape_images(child, spec_dir, page_idx)
                results.extend(child_results)
    except Exception:
        pass

    return results


def _save_picture(picture, spec_dir: Path, page_idx: int,
                   img_idx: int) -> dict[str, Any] | None:
    """Save a Picture shape's image to assets/."""
    try:
        image = picture.image
        blob = image.blob
        ext = _content_type_to_ext(image.content_type)

        assets_dir = spec_dir / "assets"
        assets_dir.mkdir(parents=True, exist_ok=True)

        name = f"p{page_idx:02d}_{img_idx:02d}"
        filename = f"{name}.{ext}"
        filepath = assets_dir / filename
        filepath.write_bytes(blob)

        # Get position info
        x = (picture.left or 0) / 914400  # EMU to inches
        y = (picture.top or 0) / 914400
        w = (picture.width or 0) / 914400
        h = (picture.height or 0) / 914400

        return {
            "type": "image",
            "src": f"assets/{filename}",
            "width_px": getattr(image, "width", 0) or 0,
            "height_px": getattr(image, "height", 0) or 0,
            "position": {"x": round(x, 2), "y": round(y, 2),
                         "w": round(w, 2), "h": round(h, 2)},
            "description": f"Image: {filename} ({w:.1f}x{h:.1f} in, at {x:.1f},{y:.1f})",
            "shape_name": picture.name,
        }
    except Exception:
        return None


def _extract_fill_image(shape, spec_dir: Path, page_idx: int,
                         img_idx: int) -> dict[str, Any] | None:
    """Extract image fill from a shape using XML."""
    try:
        xml = shape._element.xml
        # Find blipFill embed reference
        match = re.search(r'<a:blip[^>]*r:embed="([^"]*)"', xml)
        if not match:
            return None

        embed_id = match.group(1)
        rel = shape.part.rels.get(embed_id)
        if not rel:
            return None

        image = rel.target_part
        blob = image.blob
        ext = _content_type_to_ext(image.content_type)

        assets_dir = spec_dir / "assets"
        assets_dir.mkdir(parents=True, exist_ok=True)

        name = f"p{page_idx:02d}_{img_idx:02d}"
        filename = f"{name}.{ext}"
        filepath = assets_dir / filename
        filepath.write_bytes(blob)

        x = (shape.left or 0) / 914400
        y = (shape.top or 0) / 914400
        w = (shape.width or 0) / 914400
        h = (shape.height or 0) / 914400

        return {
            "type": "image_fill",
            "src": f"assets/{filename}",
            "position": {"x": round(x, 2), "y": round(y, 2),
                         "w": round(w, 2), "h": round(h, 2)},
            "description": f"Fill image: {filename} (in shape '{shape.name}')",
            "shape_name": shape.name,
        }
    except Exception:
        return None


# ── Helpers ────────────────────────────────────────────────────────────


def _content_type_to_ext(content_type: str) -> str:
    return {
        "image/png": "png",
        "image/jpeg": "jpg",
        "image/gif": "gif",
        "image/bmp": "bmp",
        "image/tiff": "tiff",
        "image/x-emf": "emf",
        "image/x-wmf": "wmf",
        "image/svg+xml": "svg",
    }.get(content_type, "bin")


def _resolve_color(color_elem: etree._Element) -> str | None:
    """Resolve OOXML color element to hex string."""
    if color_elem is None:
        return None
    tag = etree.QName(color_elem.tag).localname
    if tag == "srgbClr":
        val = color_elem.get("val")
        return f"#{val.upper()}" if val else None
    if tag == "sysClr":
        val = color_elem.get("lastClr") or color_elem.get("val")
        return f"#{val.upper()}" if val else None
    if tag == "schemeClr":
        return f"@scheme:{color_elem.get('val')}"
    return None
