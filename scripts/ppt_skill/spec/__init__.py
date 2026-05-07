"""Spec extraction module — design spec data model, theme extraction, and analysis.

Enhanced with page-level classification, VL model analysis, and directory-based
spec output. See enhanced_extractor.py for the new SpecExtractor interface.
"""

from ppt_skill.spec.spec_model import (
    ColorPalette,
    DensityLabel,
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

__all__ = [
    "ColorPalette",
    "DensityLabel",
    "DesignSpec",
    "LayoutSubType",
    "PageElement",
    "PageLayoutSpec",
    "PageType",
    "PresentationLogic",
    "Region",
    "Typography",
    "VLModelConfig",
]
