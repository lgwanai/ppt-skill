"""Enhanced design spec data model — page-level extraction with VL analysis support.

Page types: COVER, TOC, TRANSITION, CONTENT, END_PAGE
Content sub-layouts: LEFT_RIGHT, TOP_BOTTOM, LEFT_MIDDLE_RIGHT, FULL_WIDTH, GRID, etc.

Spec directory: specs/<name>/ ├── spec.yaml ├── pages/{cover,toc,transition,content,end_page}/
├── assets/ └── logic.yaml
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


class DensityLabel(str, Enum):
    BREATHING = "breathing"
    DENSE = "dense"
    ANCHOR = "anchor"


# ── Color & Typography ──────────────────────────────────────────────

@dataclass
class ColorPalette:
    background1: str = field(default="")
    background2: str = field(default="")
    text1: str = field(default="")
    text2: str = field(default="")
    accent1: str = field(default="")
    accent2: str = field(default="")
    accent3: str = field(default="")
    accent4: str = field(default="")
    accent5: str = field(default="")
    accent6: str = field(default="")
    hyperlink: str = field(default="")
    followed_hyperlink: str = field(default="")

    def to_dict(self) -> dict[str, str]:
        return {k: v for k, v in self.__dict__.items() if v}

    @classmethod
    def from_theme_scheme(cls, scheme_colors: dict[str, str]) -> ColorPalette:
        return cls(
            background1=scheme_colors.get("dk1", ""),
            background2=scheme_colors.get("dk2", ""),
            text1=scheme_colors.get("lt1", ""),
            text2=scheme_colors.get("lt2", ""),
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
    heading_family: str = field(default="")
    body_family: str = field(default="")
    heading_sizes: dict[str, float] = field(default_factory=dict)
    body_sizes: dict[str, float] = field(default_factory=dict)


# ── Layout region & element ─────────────────────────────────────────

@dataclass
class Region:
    x: float = 0.0
    y: float = 0.0
    width: float = 0.0
    height: float = 0.0
    role: str = ""
    content_type: str = ""
    description: str = ""


@dataclass
class PageElement:
    element_type: str = ""
    role: str = ""
    src: str = ""
    position: dict[str, float] = field(default_factory=dict)
    size: dict[str, float] = field(default_factory=dict)
    description: str = ""


# ── Page layout spec ────────────────────────────────────────────────

@dataclass
class PageLayoutSpec:
    page_type: PageType = field(default=PageType.CONTENT)
    layout_sub_type: LayoutSubType = field(default=LayoutSubType.FULL_WIDTH)
    description: str = ""
    vl_analysis: str = ""
    width_emu: int = 0
    height_emu: int = 0
    regions: list[Region] = field(default_factory=list)
    elements: list[PageElement] = field(default_factory=list)
    title_region: Region | None = None
    body_region: Region | None = None
    footer_region: Region | None = None
    background_description: str = ""
    background_color: str = ""
    has_background_image: bool = False
    screenshot: str = ""


# ── Presentation logic ──────────────────────────────────────────────

@dataclass
class PresentationLogic:
    page_sequence: list[str] = field(default_factory=list)
    sections: list[dict[str, Any]] = field(default_factory=list)
    narrative_pattern: str = ""
    story_arc: dict[str, int] = field(default_factory=dict)
    density_sequence: list[str] = field(default_factory=list)
    avg_content_per_page: int = 0
    transition_style: str = ""
    transition_positions: list[int] = field(default_factory=list)
    consistent_header: bool = True
    consistent_footer: bool = True
    consistent_margins: bool = True

    # Backward-compatible aliases
    @property
    def density_profile(self): return self.density_sequence
    @property
    def sequencing_pattern(self): return self.page_sequence
    @property
    def slide_types(self): return self.page_sequence


# ── VL model config ─────────────────────────────────────────────────

@dataclass
class VLModelConfig:
    provider: str = "openai"
    model: str = "gpt-4o"
    api_key: str = ""
    api_base: str = ""
    max_tokens: int = 4096
    temperature: float = 0.1
    enabled: bool = False


# ── Backward-compatible legacy types ────────────────────────────────

class SlideType(str, Enum):
    TITLE = "title"
    CONTENT = "content"
    SECTION_DIVIDER = "section_divider"
    IMAGE_TEXT = "image_text"
    DATA = "data"


@dataclass
class SlideSpec:
    slide_index: int = 0
    slide_type: SlideType = field(default=SlideType.CONTENT)
    layout_name: str = ""
    density: DensityLabel = field(default=DensityLabel.DENSE)
    char_count: int = 0
    image_count: int = 0
    shape_count: int = 0
    background: dict[str, Any] | None = None


@dataclass
class LayoutMargins:
    top: float = 0.0
    bottom: float = 0.0
    left: float = 0.0
    right: float = 0.0
    title_x: float = 0.0
    title_y: float = 0.0
    title_width: float = 0.0
    title_height: float = 0.0


@dataclass
class SlideLayoutSpec:
    slide_type: SlideType = field(default=SlideType.CONTENT)
    margins: LayoutMargins = field(default_factory=LayoutMargins)
    title_position: dict[str, float] = field(default_factory=dict)
    content_positions: list[dict[str, Any]] = field(default_factory=list)


PresentationRhythm = PresentationLogic


# ── Top-level DesignSpec (defined last — depends on all above) ──────

@dataclass
class DesignSpec:
    metadata: dict[str, Any] = field(default_factory=dict)
    colors: ColorPalette = field(default_factory=ColorPalette)
    typography: Typography = field(default_factory=Typography)
    logic: PresentationLogic = field(default_factory=PresentationLogic)
    page_types_found: list[str] = field(default_factory=list)
    layout_sub_types_found: list[str] = field(default_factory=list)
    pages: list[PageLayoutSpec] = field(default_factory=list)
    asset_count: int = 0
    source_config: dict[str, Any] = field(default_factory=dict)
    slides_legacy: list[Any] = field(default_factory=list)

    @property
    def rhythm(self): return self.logic

    @rhythm.setter
    def rhythm(self, value):
        self.logic = value

    @property
    def slides(self): return self.slides_legacy or self.pages

    @slides.setter
    def slides(self, value):
        self.slides_legacy = value

    def spec_dir(self) -> str:
        return self.metadata.get("name", "spec")

    def spec_root(self) -> str:
        return os.path.join("specs", self.spec_dir())

    def to_dict(self) -> dict:
        """Serialize to dict with backward-compatible rhythm/slides keys."""
        slides_data = []
        for i, s in enumerate(self.slides):
            if hasattr(s, "slide_index"):
                si = s.slide_index
            else:
                si = i
            if hasattr(s, "slide_type"):
                st = s.slide_type
            elif hasattr(s, "page_type"):
                st = s.page_type
            else:
                st = "content"
            slides_data.append({
                "slide_index": si,
                "slide_type": st.value if hasattr(st, "value") else str(st),
                "layout_name": getattr(s, "layout_name", ""),
                "density": getattr(s, "density", "dense").value if hasattr(getattr(s, "density", None), "value") else str(getattr(s, "density", "dense")),
                "char_count": getattr(s, "char_count", 0),
            })

        result = {
            "metadata": self.metadata,
            "colors": self.colors.to_dict() if hasattr(self.colors, "to_dict") else {},
            "typography": {
                "heading_family": self.typography.heading_family,
                "body_family": self.typography.body_family,
                "heading_sizes": self.typography.heading_sizes,
                "body_sizes": self.typography.body_sizes,
            },
            "rhythm": {
                "sequencing_pattern": self.logic.sequencing_pattern,
                "density_profile": self.logic.density_profile,
                "story_arc": self.logic.story_arc,
            },
            "slides": slides_data,
            "page_types_found": self.page_types_found,
            "layout_sub_types_found": self.layout_sub_types_found,
            "source_config": self.source_config,
        }
        return result


__all__ = [
    "ColorPalette", "DensityLabel", "DesignSpec",
    "LayoutMargins", "LayoutSubType",
    "PageElement", "PageLayoutSpec", "PageType",
    "PresentationLogic", "PresentationRhythm",
    "Region", "SlideLayoutSpec", "SlideSpec", "SlideType",
    "Typography", "VLModelConfig",
]
