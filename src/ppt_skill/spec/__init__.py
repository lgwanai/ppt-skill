"""Spec extraction module — design spec data model, theme extraction, and analysis.

This package defines the contract between Phase 2 (spec extraction) and
Phase 4 (PPT generation). All extraction modules populate the dataclass
schemas defined in spec_model.py, which serialize to YAML for storage
and reuse.
"""

from ppt_skill.spec.spec_model import (
    ColorPalette,
    DensityLabel,
    DesignSpec,
    LayoutMargins,
    PresentationRhythm,
    SlideLayoutSpec,
    SlideSpec,
    SlideType,
    Typography,
)

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
