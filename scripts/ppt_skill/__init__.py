"""PPT Skill — SVG to native-shape PowerPoint conversion pipeline."""

__version__ = "0.1.0"
__author__ = "ppt-skill contributors"

from ppt_skill.quality import SVGQualityChecker
from ppt_skill.config import CANVAS_FORMATS, DESIGN_COLORS, FONT_SIZES, LAYOUT_MARGINS

__all__ = [
    "CANVAS_FORMATS",
    "DESIGN_COLORS",
    "FONT_SIZES",
    "LAYOUT_MARGINS",
    "SVGQualityChecker",
    "__version__",
]
