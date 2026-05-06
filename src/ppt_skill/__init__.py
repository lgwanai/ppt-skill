"""PPT Skill — SVG to native-shape PowerPoint conversion pipeline."""

__version__ = "0.1.0"
__author__ = "ppt-skill contributors"

from ppt_skill.quality import SVGQualityChecker
from ppt_skill.config import CANVAS_FORMATS

__all__ = ["SVGQualityChecker", "CANVAS_FORMATS", "__version__"]
