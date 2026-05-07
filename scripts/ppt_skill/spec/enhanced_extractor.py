"""Enhanced SpecExtractor — directory-based spec extraction with VL model analysis.

Outputs a spec directory:
  specs/<name>/
  ├── spec.yaml           # Colors, fonts, metadata
  ├── pages/              # Per-page layout specs
  │   ├── cover/
  │   │   ├── page_0.yaml + page_0.png + page_0.svg
  │   ├── toc/
  │   ├── transition/
  │   ├── content/
  │   │   ├── left_right/
  │   │   │   ├── page_5.yaml + page_5.png + page_5.svg
  │   │   ├── top_bottom/
  │   │   └── ...
  │   └── end_page/
  ├── assets/             # Extracted reusable elements
  │   ├── background_0.png
  │   ├── logo_3.png
  │   └── ...
  └── logic.yaml          # Presentation logic & narrative

Integration:
  1. python-pptx + lxml: theme extraction (colors, fonts, backgrounds)
  2. python-pptx: shape extraction, layout names, text content
  3. VL model: page type classification, layout sub-type, region analysis
  4. Asset extraction: background images, logos, decorative shapes
  5. Logic analysis: narrative structure, density rhythm, sections
"""

from __future__ import annotations

import datetime
import os
import shutil
import yaml
from dataclasses import asdict, is_dataclass
from pathlib import Path
from typing import Any

from pptx import Presentation

from ppt_skill.spec.spec_model import (
    ColorPalette,
    DesignSpec,
    LayoutSubType,
    PageElement,
    PageLayoutSpec,
    PageType,
    PresentationLogic,
    Region,
    Typography,
    VLModelConfig,
)
from ppt_skill.spec.theme import (
    extract_theme_colors,
    extract_theme_fonts,
    extract_slide_background,
)
from ppt_skill.spec.screenshot import (
    extract_slide_regions,
    generate_slide_preview,
    _emu_to_px,
)
from ppt_skill.spec.vision import VLClient


# ── Page type detection (programmatic ─ VL model refines this) ──────


_PAGE_TYPE_KEYWORDS: dict[str, list[str]] = {
        "cover": ["title slide", "title", "cover", "封面", "首页", "标题幻灯片",
                   "title slide", "title"],
        "toc": ["table of contents", "contents", "agenda", "目录", "content",
                "outline", "目 录"],
        "transition": ["section", "divider", "section header", "过渡", "分隔",
                       "section divider", "transition", "章节", "部分"],
        "end_page": ["thank", "thanks", "q&a", "contact", "end", "close",
                     "谢谢", "感谢", "致谢", "尾页", "结束"],
    }


_COVER_TITLE_KEYWORDS = ["关于", "项目", "方案", "计划", "申请", "报告", "立项",
                          "预算", "采购", "招标", "述职", "总结", "规划"]


def _detect_page_type_programmatic(layout_name: str, slide_index: int,
                                    total_slides: int,
                                    regions: list) -> tuple[str, float]:
    """Quick programmatic page type detection before VL refinement."""
    name_lower = layout_name.lower()
    text_content = " ".join(r.text_preview for r in regions if r.text_preview)
    text_lower = text_content.lower()

    # --- Cover detection: first slide with title content ---
    if slide_index == 0:
        # Strong signal: layout name matches
        for kw in _PAGE_TYPE_KEYWORDS["cover"]:
            if kw in name_lower:
                return "cover", 0.9
        # Medium signal: first slide has title-like text AND Chinese doc keywords
        title_regions = [r for r in regions if r.shape_type == "text" and len(r.text_preview) > 5]
        if title_regions:
            full_text = " ".join(r.text_preview for r in title_regions)
            if any(kw in full_text for kw in _COVER_TITLE_KEYWORDS):
                return "cover", 0.75
            if len(title_regions) >= 2 and total_slides > 3:
                return "cover", 0.65

    # --- End page: last slide ---
    if slide_index == total_slides - 1:
        for kw in ["thank", "thanks", "q&a", "contact", "谢谢", "感谢"]:
            if kw in text_lower:
                return "end_page", 0.9
        # Short text on last page → likely end
        if len(text_content) < 100 and total_slides > 3:
            return "end_page", 0.6

    # TOC patterns
    toc_keywords = ["table of contents", "contents", "agenda", "目录", "outline"]
    if any(kw in text_lower for kw in toc_keywords):
        return "toc", 0.8
    # Check for list-like content in early slides
    if slide_index <= 2 and "\n" in text_content:
        lines = [l.strip() for l in text_content.split("\n") if l.strip()]
        if len(lines) >= 3 and all(len(l) < 80 for l in lines):
            return "toc", 0.5

    # Transition patterns
    for kw in ["section", "part", "partie", "章节", "部分"]:
        if kw in text_lower and len(text_content) < 200:
            return "transition", 0.7

    return "content", 0.5


def _detect_layout_sub_type_programmatic(regions: list) -> tuple[str, float]:
    """Quick programmatic layout sub-type detection."""
    text_regions = [r for r in regions if r.shape_type == "text"]
    image_regions = [r for r in regions if r.shape_type == "image"]

    if not regions:
        return "full_width", 0.5

    total_w = max(r.x + r.w for r in regions) if regions else 10000
    if total_w == 0:
        total_w = 1

    # Check for image + text patterns
    if image_regions and text_regions:
        img_left = min(r.x for r in image_regions)
        text_center = sum(r.x + r.w / 2 for r in text_regions) / len(text_regions)
        img_center = sum(r.x + r.w / 2 for r in image_regions) / len(image_regions)

        if img_center < text_center * 0.7:
            return "image_left", 0.7
        if text_center < img_center * 0.7:
            return "image_right", 0.7
        return "left_right", 0.6

    # Multi-column detection
    if len(text_regions) >= 2:
        xs = sorted(r.x for r in text_regions)
        gaps = [xs[i + 1] - xs[i] for i in range(len(xs) - 1)]
        if gaps and max(gaps) > total_w * 0.3:
            # Wide gap → two-column
            if len(xs) <= 2:
                return "left_right", 0.6
            else:
                return "grid", 0.5

    # Three-column check
    if len(text_regions) == 3:
        return "left_middle_right", 0.6

    # Single text area → top_bottom or full_width
    return "top_bottom", 0.6


# ── Asset extraction ────────────────────────────────────────────────


def _extract_background_image(slide, spec_dir: Path, page_index: int) -> str | None:
    """Extract slide background image if present."""
    try:
        bg = slide.background
        fill = bg.fill
        if fill.type is not None:
            import io
            from pptx.enum.dml import MSO_THEME_COLOR
            fill_xml = fill._fill._fill
            blip_els = fill_xml.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip")
            for blip in blip_els:
                embed = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                if embed:
                    rel = slide.part.rels[embed]
                    image_bytes = rel.target_part.blob
                    ext = rel.target_part.content_type.split("/")[-1]
                    if ext == "jpeg":
                        ext = "jpg"
                    assets_dir = spec_dir / "assets"
                    assets_dir.mkdir(parents=True, exist_ok=True)
                    filename = f"background_{page_index}.{ext}"
                    filepath = assets_dir / filename
                    filepath.write_bytes(image_bytes)
                    return f"assets/{filename}"
    except Exception:
        pass
    return None


def _extract_images_from_slide(slide, spec_dir: Path,
                                page_index: int) -> list[dict]:
    """Extract embedded images from a slide for reuse."""
    try:
        from pptx.shapes.picture import Picture
    except ImportError:
        return []

    assets: list[dict] = []
    try:
        assets_dir = spec_dir / "assets"
        assets_dir.mkdir(parents=True, exist_ok=True)

        for shape in slide.shapes:
            if hasattr(shape, "image"):
                try:
                    image = shape.image
                    ext = image.content_type.split("/")[-1]
                    if ext == "jpeg":
                        ext = "jpg"
                    filename = f"img_{page_index}_{shape.shape_id}.{ext}"
                    filepath = assets_dir / filename
                    filepath.write_bytes(image.blob)

                    # Convert EMU to inches for position/size
                    assets.append({
                        "type": "image",
                        "src": f"assets/{filename}",
                        "x": (shape.left or 0) / 914400,
                        "y": (shape.top or 0) / 914400,
                        "width": (shape.width or 0) / 914400,
                        "height": (shape.height or 0) / 914400,
                    })
                except Exception:
                    continue

    except Exception:
        pass

    return assets


# ── Page-level spec save ─────────────────────────────────────────────


def _save_page_spec(spec: PageLayoutSpec, spec_dir: Path,
                    page_index: int, prs):
    """Save a single page's layout spec to the spec directory."""
    # Determine subdirectory
    base = spec_dir / "pages" / spec.page_type.value
    if spec.page_type == PageType.CONTENT and spec.layout_sub_type.value:
        base = base / spec.layout_sub_type.value
    base.mkdir(parents=True, exist_ok=True)

    page_name = f"page_{page_index}"

    # Save YAML spec
    yaml_path = base / f"{page_name}.yaml"
    data = {
        "page_type": spec.page_type.value,
        "layout_sub_type": spec.layout_sub_type.value,
        "description": spec.description,
        "layout_description": spec.vl_analysis or spec.description,
        "width_emu": spec.width_emu,
        "height_emu": spec.height_emu,
        "background_color": spec.background_color,
        "background_description": spec.background_description,
        "has_background_image": spec.has_background_image,
        "regions": [
            {
                "x": r.x, "y": r.y,
                "width": r.width, "height": r.height,
                "role": r.role,
                "content_type": r.content_type,
                "description": r.description,
            }
            for r in spec.regions
        ],
        "elements": [
            {
                "element_type": e.element_type,
                "role": e.role,
                "src": e.src,
                "position": e.position,
                "size": e.size,
                "description": e.description,
            }
            for e in spec.elements
        ],
    }
    yaml_path.write_text(
        yaml.dump(data, default_flow_style=False, sort_keys=False, allow_unicode=True),
        encoding="utf-8",
    )

    # Generate screenshot previews
    try:
        png_path = base / f"{page_name}.png"
        svg_path = base / f"{page_name}.svg"
        generate_slide_preview(prs, page_index, png_path, svg_path)
    except Exception:
        pass


def _save_logic(logic: PresentationLogic, spec_dir: Path):
    """Save presentation logic analysis."""
    logic_path = spec_dir / "logic.yaml"
    data = {
        "page_sequence": logic.page_sequence,
        "sections": logic.sections,
        "narrative_pattern": logic.narrative_pattern,
        "story_arc": logic.story_arc,
        "density_sequence": logic.density_sequence,
        "transition_style": logic.transition_style,
        "transition_positions": logic.transition_positions,
        "avg_content_per_page": logic.avg_content_per_page,
        "consistent_header": logic.consistent_header,
        "consistent_footer": logic.consistent_footer,
        "consistent_margins": logic.consistent_margins,
    }
    logic_path.write_text(
        yaml.dump(data, default_flow_style=False, sort_keys=False, allow_unicode=True),
        encoding="utf-8",
    )


# ── Config loading ──────────────────────────────────────────────────


def _load_config(config_path: Path | None = None) -> VLModelConfig:
    """Load config from config.txt (dotenv format) in the current directory."""
    if config_path is None:
        config_path = Path("config.txt")
    if not config_path.exists():
        return VLModelConfig(enabled=False)

    config = {}
    try:
        for line in config_path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if "=" in line:
                key, _, value = line.partition("=")
                config[key.strip()] = value.strip().strip('"').strip("'")
    except Exception:
        pass

    enabled = config.get("VL_ENABLED", "false").lower() == "true"

    return VLModelConfig(
        provider=config.get("VL_PROVIDER", "openai"),
        model=config.get("VL_MODEL", "gpt-4o"),
        api_key=config.get("VL_API_KEY", ""),
        api_base=config.get("VL_API_BASE", ""),
        max_tokens=int(config.get("VL_MAX_TOKENS", "4096")),
        temperature=float(config.get("VL_TEMPERATURE", "0.1")),
        enabled=enabled,
    )


# ── Enhanced SpecExtractor ──────────────────────────────────────────


class SpecExtractor:
    """Extract a comprehensive design spec from a reference PPTX file.

    Outputs a directory-based spec::

        specs/<name>/
        ├── spec.yaml
        ├── pages/{cover,toc,transition,content,end_page}/
        ├── assets/
        └── logic.yaml

    Usage::

        extractor = SpecExtractor()
        spec = extractor.extract("reference.pptx")
        extractor.save(spec)
    """

    def __init__(self, config_path: Path | None = None):
        self.config = _load_config(config_path)
        self.vl_client: VLClient | None = None
        if self.config.enabled:
            self.vl_client = VLClient(self.config)

    def extract(self, pptx_path: Path | str) -> DesignSpec:
        """Extract design spec from a PPTX file."""
        pptx_path = Path(pptx_path)
        prs = Presentation(str(pptx_path))

        # ── Metadata ──
        spec_name = pptx_path.stem
        metadata = {
            "name": spec_name,
            "source_pptx": str(pptx_path.absolute()),
            "extracted_at": datetime.datetime.now().isoformat(),
            "slide_count": len(prs.slides),
            "format": f"{prs.slide_width}x{prs.slide_height}",
        }

        # ── Theme extraction ──
        pptx_path_str = str(pptx_path.absolute())
        theme_colors = extract_theme_colors(pptx_path_str)  # raw dict for scheme resolution
        colors = ColorPalette.from_theme_scheme(theme_colors)
        fonts_data = extract_theme_fonts(pptx_path_str)
        typography = Typography(
            heading_family=fonts_data.get("majorFont", ""),
            body_family=fonts_data.get("minorFont", ""),
        )

        # ── Per-page analysis ──
        pages: list[PageLayoutSpec] = []
        page_types_found: set[str] = set()
        layout_sub_types_found: set[str] = set()
        logic_sequence: list[str] = []
        density_sequence: list[str] = []
        all_assets: list[dict] = []
        total_slides = len(prs.slides)

        total_chars = 0
        for slide_idx, slide in enumerate(prs.slides):
            # Extract shape regions
            regions = extract_slide_regions(slide)

            # Count content characters
            char_count = sum(len(r.text_preview) for r in regions if r.shape_type == "text")
            total_chars += char_count

            # Programmatic page type detection
            layout_name = slide.slide_layout.name if slide.slide_layout else ""
            prog_type, prog_conf = _detect_page_type_programmatic(
                layout_name, slide_idx, total_slides, regions
            )

            # Programmatic layout sub-type
            prog_sub_type, prog_sub_conf = _detect_layout_sub_type_programmatic(regions)

            # VL analysis (if enabled)
            vl_result = None
            if self.vl_client and self.config.enabled:
                try:
                    # Generate preview image
                    spec_dir = Path("specs") / spec_name
                    preview_dir = spec_dir / "pages" / prog_type / "previews"
                    preview_dir.mkdir(parents=True, exist_ok=True)
                    png_path = preview_dir / f"slide_{slide_idx}.png"
                    svg_path = preview_dir / f"slide_{slide_idx}.svg"
                    generate_slide_preview(prs, slide_idx, png_path, svg_path)

                    if png_path.exists():
                        vl_result = self.vl_client.analyze_layout(png_path)
                except Exception:
                    pass

            # Merge VL result with programmatic detection
            if vl_result and vl_result.page_type:
                page_type = _str_to_page_type(vl_result.page_type)
                layout_sub_type = _str_to_layout_sub_type(vl_result.layout_sub_type)
            else:
                page_type = _str_to_page_type(prog_type)
                layout_sub_type = _str_to_layout_sub_type(prog_sub_type)
                # Create a simple description from shape analysis
                text_region_count = sum(1 for r in regions if r.shape_type == "text")
                image_count = sum(1 for r in regions if r.shape_type == "image")
                vl_desc = (
                    f"Layout with {len(regions)} shapes: "
                    f"{text_region_count} text areas, {image_count} images. "
                    f"Type: {prog_type}, Sub-type: {prog_sub_type}"
                )

            page_types_found.add(page_type.value)
            if page_type == PageType.CONTENT:
                layout_sub_types_found.add(layout_sub_type.value)

            # Build regions from VL result
            spec_regions: list[Region] = []
            if vl_result and vl_result.regions:
                for vr in vl_result.regions:
                    spec_regions.append(Region(
                        x=vr.get("x", 0) / 100.0,
                        y=vr.get("y", 0) / 100.0,
                        width=vr.get("width", 0) / 100.0,
                        height=vr.get("height", 0) / 100.0,
                        role=vr.get("role", ""),
                        content_type=vr.get("content_type", ""),
                        description=vr.get("description", ""),
                    ))
            else:
                # Programmatic regions from shape positions
                for r in regions:
                    spec_regions.append(Region(
                        x=r.x / prs.slide_width if prs.slide_width else 0,
                        y=r.y / prs.slide_height if prs.slide_height else 0,
                        width=r.w / prs.slide_width if prs.slide_width else 0,
                        height=r.h / prs.slide_height if prs.slide_height else 0,
                        role=_infer_role(r),
                        content_type=r.shape_type,
                        description=r.text_preview,
                    ))

            # Extract assets
            assets = _extract_images_from_slide(
                slide, Path("specs") / spec_name, slide_idx
            )
            all_assets.extend(assets)

            # Background
            bg_result = extract_slide_background(slide)
            bg_color = ""
            bg_desc = ""
            if bg_result and bg_result.get("color"):
                bg_color = bg_result["color"]
                bg_desc = bg_result.get("description", "")
                # Resolve @scheme: references using extracted theme colors
                if bg_color.startswith("@scheme:"):
                    scheme_key = bg_color.split(":", 1)[1]
                    resolved = theme_colors.get(scheme_key, "")
                    if resolved and resolved.startswith("#"):
                        bg_color = resolved
            else:
                # No explicit background → default white background
                bg_color = "#FFFFFF"
                bg_desc = "White background (PowerPoint default)"

            # Build page spec
            page_spec = PageLayoutSpec(
                page_type=page_type,
                layout_sub_type=layout_sub_type,
                description=vl_result.layout_description if vl_result else vl_desc,
                vl_analysis=vl_result.raw_response if vl_result else "",
                width_emu=int(prs.slide_width),
                height_emu=int(prs.slide_height),
                regions=spec_regions,
                elements=_extract_elements(slide, assets),
                background_description=(
                    vl_result.background_description if vl_result else bg_desc
                ),
                background_color=bg_color,
                has_background_image=False,  # Detected in save phase
            )

            pages.append(page_spec)
            logic_sequence.append(page_type.value)

            # Density classification
            if char_count < 100:
                density_sequence.append("breathing")
            elif char_count < 500:
                density_sequence.append("dense")
            else:
                density_sequence.append("anchor")

        # ── Presentation logic ──
        logic = PresentationLogic(
            page_sequence=logic_sequence,
            sections=_build_sections(pages),
            density_sequence=density_sequence,
            avg_content_per_page=total_chars // max(total_slides, 1),
        )

        # VL logic analysis
        if self.vl_client and self.config.enabled and len(pages) <= 20:
            try:
                summaries = [
                    {
                        "page_type": p.page_type.value,
                        "summary": p.description[:200],
                        "density": density_sequence[i],
                    }
                    for i, p in enumerate(pages)
                ]
                logic_data = self.vl_client.analyze_logic(summaries)
                if logic_data:
                    logic.narrative_pattern = logic_data.get("narrative_pattern", "")
                    logic.story_arc = logic_data.get("story_arc", {})
                    logic.transition_style = logic_data.get("transition_style", "")
            except Exception:
                pass

        return DesignSpec(
            metadata=metadata,
            colors=colors,
            typography=typography,
            logic=logic,
            page_types_found=sorted(page_types_found),
            layout_sub_types_found=sorted(layout_sub_types_found),
            pages=pages,
            asset_count=len(all_assets),
        )

    def save(self, spec: DesignSpec, base_dir: str = "specs"):
        """Save the spec as a directory structure."""
        spec_dir = Path(base_dir) / spec.metadata.get("name", "spec")
        spec_dir.mkdir(parents=True, exist_ok=True)

        # ── spec.yaml ──
        spec_data = {
            "metadata": spec.metadata,
            "colors": spec.colors.to_dict(),
            "typography": {
                "heading_family": spec.typography.heading_family,
                "body_family": spec.typography.body_family,
                "heading_sizes": spec.typography.heading_sizes,
                "body_sizes": spec.typography.body_sizes,
            },
            "page_types_found": spec.page_types_found,
            "layout_sub_types_found": spec.layout_sub_types_found,
            "asset_count": spec.asset_count,
        }
        (spec_dir / "spec.yaml").write_text(
            yaml.dump(spec_data, default_flow_style=False,
                     sort_keys=False, allow_unicode=True),
            encoding="utf-8",
        )

        # ── logic.yaml ──
        _save_logic(spec.logic, spec_dir)

        # ── Per-page specs ──
        # Need the original PPTX to generate previews
        src_path = spec.metadata.get("source_pptx", "")
        if src_path and Path(src_path).exists():
            prs = Presentation(src_path)
        else:
            prs = None

        for i, page in enumerate(spec.pages):
            try:
                _save_page_spec(page, spec_dir, i, prs)
            except Exception:
                pass

        print(f"Spec saved to: {spec_dir}/")
        print(f"  Pages: {len(spec.pages)} ({', '.join(spec.page_types_found)})")
        print(f"  Content sub-types: {', '.join(spec.layout_sub_types_found) or 'none'}")
        print(f"  Assets: {spec.asset_count}")

    def save_page(self, spec: DesignSpec, page_index: int, prs=None):
        """Save a single page from the spec."""
        if page_index >= len(spec.pages):
            return
        spec_dir = Path("specs") / spec.metadata.get("name", "spec")
        _save_page_spec(spec.pages[page_index], spec_dir, page_index, prs)


# ── Helper functions ─────────────────────────────────────────────────


def _str_to_page_type(s: str) -> PageType:
    mapping = {
        "cover": PageType.COVER,
        "toc": PageType.TOC,
        "transition": PageType.TRANSITION,
        "content": PageType.CONTENT,
        "end_page": PageType.END_PAGE,
        "end": PageType.END_PAGE,
        "blank": PageType.BLANK,
    }
    return mapping.get(s.lower().strip(), PageType.CONTENT)


def _str_to_layout_sub_type(s: str) -> LayoutSubType:
    mapping = {
        "left_right": LayoutSubType.LEFT_RIGHT,
        "top_bottom": LayoutSubType.TOP_BOTTOM,
        "left_middle_right": LayoutSubType.LEFT_MIDDLE_RIGHT,
        "full_width": LayoutSubType.FULL_WIDTH,
        "grid": LayoutSubType.GRID,
        "image_left": LayoutSubType.IMAGE_LEFT,
        "image_right": LayoutSubType.IMAGE_RIGHT,
        "quote": LayoutSubType.QUOTE,
        "chart": LayoutSubType.CHART,
        "custom": LayoutSubType.CUSTOM,
    }
    return mapping.get(s.lower().strip().replace("-", "_"), LayoutSubType.CUSTOM)


def _infer_role(region) -> str:
    """Infer the semantic role of a shape region."""
    text = region.text_preview.lower()
    stype = region.shape_type

    if stype == "image":
        return "image"
    if stype == "table":
        return "table"
    if stype == "chart":
        return "chart"

    # Text region role inference
    if any(kw in text for kw in ["title", "标题", "introduction", "introducing"]):
        return "title"
    if any(kw in text for kw in ["page", "页码", "slide", "©"]):
        return "footer"
    if any(kw in text for kw in ["logo", "公司", "company", "brand"]):
        return "logo"
    if len(text) < 20 and "\n" in text:
        return "title"
    if len(text) < 10:
        return "decoration"

    return "body"


def _extract_elements(slide, assets: list[dict]) -> list[PageElement]:
    """Extract reusable visual elements from a slide."""
    elements: list[PageElement] = []
    for asset in assets:
        elements.append(PageElement(
            element_type="image",
            role="image",
            src=asset.get("src", ""),
            position={"x": asset.get("x", 0), "y": asset.get("y", 0)},
            size={"width": asset.get("width", 0), "height": asset.get("height", 0)},
        ))
    return elements


def _build_sections(pages: list[PageLayoutSpec]) -> list[dict[str, Any]]:
    """Group pages into logical sections based on page types."""
    sections: list[dict] = []
    current: dict | None = None

    for i, page in enumerate(pages):
        if page.page_type in (PageType.TRANSITION, PageType.TOC):
            if current and current.get("slides"):
                sections.append(current)
            current = {"name": f"Section {len(sections) + 1}", "slides": [i],
                       "page_types": [page.page_type.value]}
        elif page.page_type == PageType.COVER:
            if current and current.get("slides"):
                sections.append(current)
            current = {"name": "Opening", "slides": [i],
                       "page_types": [page.page_type.value]}
        elif page.page_type == PageType.END_PAGE:
            if current and current.get("slides"):
                current["slides"].append(i)
                current["page_types"].append(page.page_type.value)
                sections.append(current)
            else:
                sections.append({"name": "Closing", "slides": [i],
                                "page_types": [page.page_type.value]})
            current = None
        else:
            if current is None:
                current = {"name": f"Section {len(sections) + 1}", "slides": [],
                           "page_types": []}
            current["slides"].append(i)
            current["page_types"].append(page.page_type.value)

    if current and current.get("slides"):
        sections.append(current)

    return sections


# ── Legacy-compatible functions ─────────────────────────────────────


def _dataclass_to_dict(obj) -> dict:
    """Convert dataclass to dict for YAML serialization."""
    from enum import Enum
    if is_dataclass(obj):
        result = {}
        for field_name in obj.__dataclass_fields__:
            value = getattr(obj, field_name)
            result[field_name] = _dataclass_to_dict(value)
        return result
    elif isinstance(obj, Enum):
        return obj.value
    elif isinstance(obj, list):
        return [_dataclass_to_dict(item) for item in obj]
    elif isinstance(obj, dict):
        return {k: _dataclass_to_dict(v) for k, v in obj.items()}
    else:
        return obj
