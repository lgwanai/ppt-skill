"""Design spec data model — complete design blueprint for PPT reproduction.

Every element captured in this spec must be independently reproducible —
given this YAML, one should be able to redraw the original slide with
pixel-level fidelity on style (not content text).

Spec directory:
  specs/<name>/
  ├── spec.yaml        # Palette, fonts, page list, logic
  ├── pages/           # Per-page blueprint YAMLs
  │   ├── cover/
  │   ├── toc/
  │   ├── content/{sub_type}/
  │   └── end_page/
  ├── assets/          # Extracted images, backgrounds
  └── logic.yaml       # Narrative structure
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from enum import Enum
from typing import Any


# ── Enums ───────────────────────────────────────────────────────────

class PageType(str, Enum):
    COVER = "cover"
    TOC = "toc"
    TRANSITION = "transition"
    CONTENT = "content"
    END_PAGE = "end_page"
    BLANK = "blank"


class LayoutSubType(str, Enum):
    LEFT_RIGHT = "left_right"
    TOP_BOTTOM = "top_bottom"
    LEFT_MIDDLE_RIGHT = "left_middle_right"
    FULL_WIDTH = "full_width"
    GRID = "grid"
    IMAGE_LEFT = "image_left"
    IMAGE_RIGHT = "image_right"
    QUOTE = "quote"
    CHART = "chart"
    CUSTOM = "custom"


class ElementType(str, Enum):
    TEXT = "text"
    IMAGE = "image"
    SHAPE = "shape"
    GROUP = "group"
    TABLE = "table"
    CHART = "chart"
    PLACEHOLDER = "placeholder"


class SemanticRole(str, Enum):
    TITLE = "title"
    SUBTITLE = "subtitle"
    BODY = "body"
    FOOTER = "footer"
    PAGE_NUMBER = "page_number"
    LOGO = "logo"
    DECORATION = "decoration"
    IMAGE = "image"
    ICON = "icon"
    ACCENT = "accent"
    DIVIDER = "divider"
    HEADER = "header"


class DensityLabel(str, Enum):
    BREATHING = "breathing"
    DENSE = "dense"
    ANCHOR = "anchor"


# ── Style dataclasses ────────────────────────────────────────────────


@dataclass
class TextStyle:
    """Complete text formatting specification."""
    font_family: str = ""
    font_size_pt: float = 0.0
    font_weight: str = "normal"     # "normal" | "bold"
    font_italic: bool = False
    font_underline: bool = False
    font_color: str = ""            # HEX, e.g. "#4472C4"
    text_alignment: str = "left"    # "left" | "center" | "right" | "justify"
    line_spacing: float = 1.0       # multiplier (1.0 = single, 1.5 = 1.5 lines)
    letter_spacing: float = 0.0     # pt

    def to_dict(self) -> dict:
        return {
            "font_family": self.font_family,
            "font_size_pt": self.font_size_pt,
            "font_weight": self.font_weight,
            "font_italic": self.font_italic,
            "font_underline": self.font_underline,
            "font_color": self.font_color,
            "text_alignment": self.text_alignment,
            "line_spacing": self.line_spacing,
            "letter_spacing": self.letter_spacing,
        }


@dataclass
class ShapeStyle:
    """Shape/rectangle visual style."""
    shape_type: str = "rect"        # "rect" | "round_rect" | "circle" | "line" | "triangle" | "arrow" | "parallelogram" | "custom"
    fill_color: str = ""            # HEX, empty = transparent
    fill_opacity: float = 1.0
    stroke_color: str = ""          # HEX, empty = no stroke
    stroke_width_pt: float = 0.0
    corner_radius_pt: float = 0.0   # for round_rect

    def to_dict(self) -> dict:
        return {
            "shape_type": self.shape_type,
            "fill_color": self.fill_color,
            "fill_opacity": self.fill_opacity,
            "stroke_color": self.stroke_color,
            "stroke_width_pt": self.stroke_width_pt,
            "corner_radius_pt": self.corner_radius_pt,
        }


@dataclass
class ImageStyle:
    """Image placement specification."""
    src: str = ""                   # relative path in assets/
    original_width_px: int = 0
    original_height_px: int = 0
    crop_left: float = 0.0
    crop_right: float = 0.0
    crop_top: float = 0.0
    crop_bottom: float = 0.0
    brightness: float = 0.0
    contrast: float = 0.0

    def to_dict(self) -> dict:
        return {
            "src": self.src,
            "original_width_px": self.original_width_px,
            "original_height_px": self.original_height_px,
            "crop": {
                "left": self.crop_left,
                "right": self.crop_right,
                "top": self.crop_top,
                "bottom": self.crop_bottom,
            },
        }


# ── Position ─────────────────────────────────────────────────────────


@dataclass
class Pos:
    """Normalized position (0–1) in page coordinates."""
    x: float = 0.0
    y: float = 0.0
    w: float = 0.0
    h: float = 0.0

    def to_dict(self) -> dict:
        return {"x": round(self.x, 4), "y": round(self.y, 4),
                "w": round(self.w, 4), "h": round(self.h, 4)}


# ── Element — the core unit ──────────────────────────────────────────


@dataclass
class Element:
    """A single visual element on a slide — fully specified for reproduction."""
    id: int = 0
    element_type: ElementType = field(default=ElementType.SHAPE)
    semantic_role: SemanticRole = field(default=SemanticRole.BODY)
    position: Pos = field(default_factory=Pos)

    # Content
    text: str = ""                  # Full text content (not truncated!)
    text_style: TextStyle | None = None
    shape_style: ShapeStyle | None = None
    image_style: ImageStyle | None = None

    # Children (for GROUP elements)
    children: list[Element] = field(default_factory=list)

    # Metadata
    z_order: int = 0
    visible: bool = True
    shape_name: str = ""            # Original PowerPoint shape name

    def to_dict(self) -> dict:
        d: dict[str, Any] = {
            "id": self.id,
            "type": self.element_type.value,
            "role": self.semantic_role.value,
            "position": self.position.to_dict(),
            "z_order": self.z_order,
        }
        if self.text:
            d["text"] = self.text
        if self.text_style:
            d["text_style"] = self.text_style.to_dict()
        if self.shape_style:
            d["shape_style"] = self.shape_style.to_dict()
        if self.image_style:
            d["image_style"] = self.image_style.to_dict()
        if self.children:
            d["children"] = [c.to_dict() for c in self.children]
        if self.shape_name:
            d["shape_name"] = self.shape_name
        return d


# ── Page spec — complete slides blueprint ────────────────────────────


@dataclass
class PageSpec:
    """Complete page blueprint — everything needed to redraw this slide."""
    page_type: PageType = field(default=PageType.CONTENT)
    layout_sub_type: LayoutSubType = field(default=LayoutSubType.FULL_WIDTH)

    # Canvas
    width_emu: int = 0
    height_emu: int = 0
    width_inches: float = 0.0
    height_inches: float = 0.0

    # Background
    background_color: str = ""
    background_image: str = ""
    background_type: str = "solid"
    background_description: str = ""
    gradient_stops: list[dict] = field(default_factory=list)

    # Visual hierarchy summary
    visual_hierarchy: list[str] = field(default_factory=list)
    # ["title: 等线 Light 36pt bold #000000", "body: 等线 14pt #333333", "footer: 等线 10pt #999999"]

    # Elements
    elements: list[Element] = field(default_factory=list)

    # Descriptive
    layout_description: str = ""

    @property
    def element_count(self) -> int:
        return len(self.elements)

    def to_dict(self) -> dict:
        return {
            "page_type": self.page_type.value,
            "layout_sub_type": self.layout_sub_type.value,
            "canvas": {
                "width_emu": self.width_emu,
                "height_emu": self.height_emu,
                "width_inches": round(self.width_inches, 2),
                "height_inches": round(self.height_inches, 2),
            },
            "background": {
                "type": self.background_type,
                "color": self.background_color,
                "image": self.background_image,
                "description": self.background_description,
                "gradient_stops": self.gradient_stops,
            },
            "visual_hierarchy": self.visual_hierarchy,
            "layout_description": self.layout_description,
            "elements": [e.to_dict() for e in self.elements],
        }


# ── Color & Typography ──────────────────────────────────────────────


@dataclass
class ColorPalette:
    accent1: str = ""
    accent2: str = ""
    accent3: str = ""
    accent4: str = ""
    accent5: str = ""
    accent6: str = ""
    dark1: str = ""      # Usually background
    dark2: str = ""
    light1: str = ""     # Usually text
    light2: str = ""
    hyperlink: str = ""
    followed_hyperlink: str = ""

    def to_dict(self) -> dict[str, str]:
        return {k: v for k, v in self.__dict__.items() if v}

    @classmethod
    def from_theme_scheme(cls, scheme_colors: dict[str, str]) -> ColorPalette:
        return cls(
            dark1=scheme_colors.get("dk1", ""),
            dark2=scheme_colors.get("dk2", ""),
            light1=scheme_colors.get("lt1", ""),
            light2=scheme_colors.get("lt2", ""),
            accent1=scheme_colors.get("accent1", ""),
            accent2=scheme_colors.get("accent2", ""),
            accent3=scheme_colors.get("accent3", ""),
            accent4=scheme_colors.get("accent4", ""),
            accent5=scheme_colors.get("accent5", ""),
            accent6=scheme_colors.get("accent6", ""),
            hyperlink=scheme_colors.get("hlink", ""),
            followed_hyperlink=scheme_colors.get("folHlink", ""),
        )


@dataclass
class Typography:
    heading_family: str = ""
    body_family: str = ""

    def to_dict(self) -> dict:
        return {"heading_family": self.heading_family, "body_family": self.body_family}


# ── Logic ────────────────────────────────────────────────────────────


@dataclass
class PresentationLogic:
    page_sequence: list[str] = field(default_factory=list)
    sections: list[dict] = field(default_factory=list)
    narrative_pattern: str = ""
    story_arc: dict[str, int] = field(default_factory=dict)
    density_sequence: list[str] = field(default_factory=list)
    transition_style: str = ""
    avg_content_per_page: int = 0

    @property
    def density_profile(self): return self.density_sequence

    @property
    def sequencing_pattern(self): return self.page_sequence

    def to_dict(self) -> dict:
        return {
            "page_sequence": self.page_sequence,
            "sections": self.sections,
            "narrative_pattern": self.narrative_pattern,
            "story_arc": self.story_arc,
            "density_sequence": self.density_sequence,
            "transition_style": self.transition_style,
            "avg_content_per_page": self.avg_content_per_page,
        }


# ── Root Spec ────────────────────────────────────────────────────────


@dataclass
class DesignSpec:
    metadata: dict = field(default_factory=dict)
    palette: ColorPalette = field(default_factory=ColorPalette)
    typography: Typography = field(default_factory=Typography)
    pages: list[PageSpec] = field(default_factory=list)
    logic: PresentationLogic = field(default_factory=PresentationLogic)
    page_types_found: list[str] = field(default_factory=list)
    layout_sub_types_found: list[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "metadata": self.metadata,
            "palette": self.palette.to_dict(),
            "typography": self.typography.to_dict(),
            "page_types_found": self.page_types_found,
            "layout_sub_types_found": self.layout_sub_types_found,
            "pages_summary": [
                f"page_{i}: {p.page_type.value}/{p.layout_sub_type.value}"
                for i, p in enumerate(self.pages)
            ],
        }


# ── VL config ────────────────────────────────────────────────────────


@dataclass
class VLModelConfig:
    provider: str = "openai"
    model: str = "gpt-4o"
    api_key: str = ""
    api_base: str = ""
    max_tokens: int = 4096
    temperature: float = 0.1
    enabled: bool = False


# ── Legacy compatibility ─────────────────────────────────────────────

class SlideType(str, Enum):
    TITLE = "title"
    CONTENT = "content"
    SECTION_DIVIDER = "section_divider"
    IMAGE_TEXT = "image_text"
    DATA = "data"


PresentationRhythm = PresentationLogic


__all__ = [
    "ColorPalette", "DensityLabel", "DesignSpec",
    "Element", "ElementType", "ImageStyle",
    "LayoutSubType", "PageSpec", "PageType",
    "Pos", "PresentationLogic", "PresentationRhythm",
    "SemanticRole", "ShapeStyle", "SlideType",
    "TextStyle", "Typography", "VLModelConfig",
]
