"""Asset extraction — backgrounds, icons, and reusable decorative elements.

Extracts and CLASSIFIES visual assets:
  - BACKGROUND: Full-slide backgrounds (keep)
  - ICON: Small images in margins/corners, typically repeated (keep)
  - DECORATION: Accent shapes, lines, decorative image fills (keep)
  - CHART: Data visualizations (DISCARD — content-dependent)
  - CONTENT_IMAGE: Photos, screenshots in content area (DISCARD — content-dependent)

Classification rules:
  - area < 3% + in margin → ICON
  - area < 15% + not main content → DECORATION
  - chart/table shapes → DISCARD
  - everything else in content zone → DISCARD

All kept assets saved to specs/<name>/assets/.
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Any

from lxml import etree

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


class AssetType(str, Enum):
    BACKGROUND = "background"
    ICON = "icon"
    DECORATION = "decoration"
    CHART = "chart"
    CONTENT_IMAGE = "content_image"


@dataclass
class AssetInfo:
    type: AssetType
    src: str
    description: str
    shape_name: str = ""
    position: dict = None
    size_ratio: float = 0.0  # fraction of slide area

    def to_dict(self) -> dict:
        d = {"type": self.type.value, "src": self.src, "description": self.description}
        if self.shape_name:
            d["shape_name"] = self.shape_name
        if self.position:
            d["position"] = self.position
        return d


# ── Classification ────────────────────────────────────────────────────


def classify_image(img_w_emu: int, img_h_emu: int,
                   slide_w_emu: int, slide_h_emu: int,
                   shape_name: str = "",
                   is_chart_shape: bool = False) -> AssetType:
    """Classify an image asset based on size and position.

    Uses heuristics:
        - Charts/tables    → DISCARD
        - < 3% slide area  → ICON (if in margins)
        - < 15% slide area → DECORATION
        - > 50% slide area → BACKGROUND
        - Everything else  → CONTENT_IMAGE (discard)
    """
    if is_chart_shape:
        return AssetType.CHART

    if slide_w_emu <= 0 or slide_h_emu <= 0:
        return AssetType.CONTENT_IMAGE

    slide_area = slide_w_emu * slide_h_emu
    img_area = img_w_emu * img_h_emu
    ratio = img_area / slide_area

    name_lower = shape_name.lower()

    # Very small → likely icon if named like one
    if ratio < 0.03:
        if any(kw in name_lower for kw in ["icon", "logo", "图标"]):
            return AssetType.ICON
        return AssetType.DECORATION

    # Small decorative elements
    if ratio < 0.15:
        if any(kw in name_lower for kw in ["icon", "logo", "图标", "decor", "装饰"]):
            return AssetType.ICON
        return AssetType.DECORATION

    # Very large → background
    if ratio > 0.5:
        return AssetType.BACKGROUND

    # Medium size in content area → content image (discard)
    return AssetType.CONTENT_IMAGE


# ── Background extraction ─────────────────────────────────────────────


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


# ── Shape image extraction ────────────────────────────────────────────


def extract_shape_assets(shape, spec_dir: Path, page_idx: int,
                          slide_w_emu: int = 0, slide_h_emu: int = 0,
                          img_counter: list[int] | None = None
                          ) -> list[AssetInfo]:
    """Extract and classify images from a shape (Picture, image fill, group children).

    Only keeps: BACKGROUND, ICON, DECORATION.
    Discards: CHART, CONTENT_IMAGE.
    """
    if img_counter is None:
        img_counter = [0]
    results: list[AssetInfo] = []

    # Direct Picture shapes
    try:
        from pptx.shapes.picture import Picture
        if isinstance(shape, Picture):
            asset = _extract_picture_asset(shape, spec_dir, page_idx,
                                            img_counter, slide_w_emu, slide_h_emu)
            if asset and asset.type in (AssetType.BACKGROUND, AssetType.ICON,
                                         AssetType.DECORATION):
                results.append(asset)
            return results
    except ImportError:
        pass

    # Skip chart/table shapes
    try:
        if hasattr(shape, "has_table") and shape.has_table:
            return results
        if hasattr(shape, "chart"):
            shape.chart
            return results
    except Exception:
        pass

    # Shapes with image fill
    try:
        if hasattr(shape, "fill"):
            from pptx.enum.dml import MSO_FILL_TYPE
            if shape.fill.type == MSO_FILL_TYPE.PICTURE:
                asset = _extract_fill_asset(shape, spec_dir, page_idx,
                                             img_counter, slide_w_emu, slide_h_emu)
                if asset and asset.type in (AssetType.BACKGROUND, AssetType.ICON,
                                             AssetType.DECORATION):
                    results.append(asset)
    except Exception:
        pass

    # Group shapes — recurse
    try:
        if hasattr(shape, "shapes"):
            for child in shape.shapes:
                child_results = extract_shape_assets(
                    child, spec_dir, page_idx, slide_w_emu, slide_h_emu, img_counter
                )
                results.extend(child_results)
    except Exception:
        pass

    return results


def _extract_picture_asset(picture, spec_dir: Path, page_idx: int,
                            img_counter: list[int],
                            slide_w_emu: int, slide_h_emu: int) -> AssetInfo | None:
    """Extract and classify a Picture shape."""
    try:
        image = picture.image
        blob = image.blob
        ext = _content_type_to_ext(image.content_type)

        # Classify first (don't save charts/content)
        asset_type = classify_image(
            img_w_emu=(picture.width or 0),
            img_h_emu=(picture.height or 0),
            slide_w_emu=slide_w_emu,
            slide_h_emu=slide_h_emu,
            shape_name=picture.name,
        )

        if asset_type in (AssetType.CHART, AssetType.CONTENT_IMAGE):
            return AssetInfo(type=asset_type, src="", description=f"[DISCARDED] {picture.name}")

        # Save
        assets_dir = spec_dir / "assets"
        assets_dir.mkdir(parents=True, exist_ok=True)
        prefix = asset_type.value[:3]  # "ico", "dec", "bac"
        n = img_counter[0]
        img_counter[0] += 1
        filename = f"{prefix}_{page_idx:02d}_{n:02d}.{ext}"
        (assets_dir / filename).write_bytes(blob)

        x = (picture.left or 0) / 914400
        y = (picture.top or 0) / 914400
        w = (picture.width or 0) / 914400
        h = (picture.height or 0) / 914400

        return AssetInfo(
            type=asset_type,
            src=f"assets/{filename}",
            shape_name=picture.name,
            position={"x": round(x, 2), "y": round(y, 2), "w": round(w, 2), "h": round(h, 2)},
            description=f"{asset_type.value}: {filename} ({w:.1f}x{h:.1f}in, at {x:.1f},{y:.1f})",
            size_ratio=((picture.width or 0) * (picture.height or 0)) / (slide_w_emu * slide_h_emu) if slide_w_emu and slide_h_emu else 0,
        )
    except Exception:
        return None


def _extract_fill_asset(shape, spec_dir: Path, page_idx: int,
                         img_counter: list[int],
                         slide_w_emu: int, slide_h_emu: int) -> AssetInfo | None:
    """Extract and classify a shape's image fill."""
    try:
        # Classify first
        asset_type = classify_image(
            img_w_emu=(shape.width or 0),
            img_h_emu=(shape.height or 0),
            slide_w_emu=slide_w_emu,
            slide_h_emu=slide_h_emu,
            shape_name=shape.name,
        )

        if asset_type in (AssetType.CHART, AssetType.CONTENT_IMAGE):
            return AssetInfo(type=asset_type, src="", description=f"[DISCARDED] {shape.name} fill")

        # Extract blip embed from XML
        xml = shape._element.xml
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
        prefix = asset_type.value[:3]
        n = img_counter[0]
        img_counter[0] += 1
        filename = f"{prefix}_{page_idx:02d}_{n:02d}.{ext}"
        (assets_dir / filename).write_bytes(blob)

        x = (shape.left or 0) / 914400
        y = (shape.top or 0) / 914400
        w = (shape.width or 0) / 914400
        h = (shape.height or 0) / 914400

        return AssetInfo(
            type=asset_type,
            src=f"assets/{filename}",
            shape_name=shape.name,
            position={"x": round(x, 2), "y": round(y, 2), "w": round(w, 2), "h": round(h, 2)},
            description=f"{asset_type.value}: {filename} (fill in '{shape.name}')",
            size_ratio=((shape.width or 0) * (shape.height or 0)) / (slide_w_emu * slide_h_emu) if slide_w_emu and slide_h_emu else 0,
        )
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
