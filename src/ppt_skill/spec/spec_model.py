"""Design spec data model — dataclass schemas for the extraction-generation contract.

All types are plain Python dataclasses (NOT Pydantic) to keep dependencies
minimal. Every field has default values so DesignSpec() can be constructed
partially and populated incrementally by extraction modules.

These schemas serialize to YAML for project-local storage and are consumed
by Phase 4 (PPT Generation) to faithfully reproduce extracted designs.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any


class SlideType(str, Enum):
    """Slide type classification based on content structure and layout name."""

    TITLE = "title"
    CONTENT = "content"
    SECTION_DIVIDER = "section_divider"
    IMAGE_TEXT = "image_text"
    DATA = "data"


class DensityLabel(str, Enum):
    """Content density classification per slide.

    Breathing: < 20th percentile character count (sparse, visual-oriented)
    Dense: 20th–80th percentile (balanced text and visuals)
    Anchor: > 80th percentile (text-heavy, detail-oriented)
    """

    BREATHING = "breathing"
    DENSE = "dense"
    ANCHOR = "anchor"


@dataclass
class ColorPalette:
    """12-color palette extracted from the PPTX theme's clrScheme.

    Maps OOXML scheme names to semantic field names:
      dk1 → background1, dk2 → background2
      lt1 → text1, lt2 → text2
      accent1–6 → accent1–6
      hlink → hyperlink, folHlink → followed_hyperlink

    All values are HEX strings with '#' prefix (e.g., "#4472C4").
    Empty string means the color was not found in the theme.
    """

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
        """Return all non-empty color fields as a flat dict."""
        return {k: v for k, v in self.__dict__.items() if v}

    @classmethod
    def from_theme_scheme(cls, scheme_colors: dict[str, str]) -> ColorPalette:
        """Build ColorPalette from raw theme scheme color dict.

        Args:
            scheme_colors: Dict mapping OOXML scheme names (dk1, lt1, accent1, etc.)
                           to HEX color strings (with '#' prefix).
        """
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
    """Font families and size hierarchy extracted from the theme's fontScheme.

    heading_family / body_family come from majorFont/minorFont in theme1.xml.
    heading_sizes / body_sizes provide the font size hierarchy in points (float).
    """

    heading_family: str = field(default="")
    body_family: str = field(default="")
    heading_sizes: dict[str, float] = field(default_factory=dict)
    body_sizes: dict[str, float] = field(default_factory=dict)


@dataclass
class LayoutMargins:
    """Slide-level margin and title positioning data.

    All measurements in inches (converted from EMU by extraction code).
    """

    top: float = field(default=0.0)
    bottom: float = field(default=0.0)
    left: float = field(default=0.0)
    right: float = field(default=0.0)
    title_x: float = field(default=0.0)
    title_y: float = field(default=0.0)
    title_width: float = field(default=0.0)
    title_height: float = field(default=0.0)


@dataclass
class SlideLayoutSpec:
    """Spatial layout specification for a classified slide type.

    Captures margins, title positioning, and content area positions
    extracted from the PPTX slide and layout XML.
    """

    slide_type: SlideType = field(default=SlideType.CONTENT)
    margins: LayoutMargins = field(default_factory=LayoutMargins)
    title_position: dict[str, float] = field(default_factory=dict)
    content_positions: list[dict[str, Any]] = field(default_factory=list)


@dataclass
class SlideSpec:
    """Per-slide analysis data combining classification, density, and background.

    Each slide in the source PPTX produces one SlideSpec instance.
    """

    slide_index: int = field(default=0)
    slide_type: SlideType = field(default=SlideType.CONTENT)
    layout_name: str = field(default="")
    density: DensityLabel = field(default=DensityLabel.DENSE)
    char_count: int = field(default=0)
    image_count: int = field(default=0)
    shape_count: int = field(default=0)
    background: dict[str, Any] | None = field(default=None)


@dataclass
class PresentationRhythm:
    """Overall presentation logic — sequencing, density flow, and story arc.

    sequencing_pattern: Ordered list of slide types (e.g., ["title", "content", ...]).
    density_profile: Ordered list of density labels matching slide order.
    story_arc: Distribution of slides across narrative sections
               ({"opening": N, "development": N, "climax": N, "closing": N}).
    """

    sequencing_pattern: list[str] = field(default_factory=list)
    density_profile: list[str] = field(default_factory=list)
    story_arc: dict[str, int] = field(default_factory=dict)


@dataclass
class DesignSpec:
    """Root design specification — seeded by extraction, consumed by generation.

    This is the primary serialization target. YAML dumps of DesignSpec
    instances are stored in the specs/ directory and loaded by Phase 4.

    All fields have defaults so extraction can populate incrementally.
    """

    metadata: dict[str, Any] = field(default_factory=dict)
    colors: ColorPalette = field(default_factory=ColorPalette)
    typography: Typography = field(default_factory=Typography)
    slides: list[SlideSpec] = field(default_factory=list)
    rhythm: PresentationRhythm = field(default_factory=PresentationRhythm)
    source_config: dict[str, Any] = field(default_factory=dict)


__all__ = [
    "ColorPalette",
    "DensityLabel",
    "DesignSpec",
    "LayoutMargins",
    "PresentationRhythm",
    "SlideLayoutSpec",
    "SlideSpec",
    "SlideType",
    "Typography",
]
