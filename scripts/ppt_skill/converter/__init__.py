"""ppt_skill.converter — SVG to PPTX conversion package (forked from ppt-master).

Public API:
    - main(): CLI entry point
    - convert_svg_to_slide_shapes(): SVG -> DrawingML slide XML
    - create_pptx_with_native_svg(): Build PPTX from SVG files

Template resolution: templates/ lives at project root (sibling to src/).
Use: Path(__file__).resolve().parent.parent.parent.parent / 'assets' / 'templates' / 'icons'
"""

from .cli import main
from .converter import convert_svg_to_slide_shapes
from .builder import create_pptx_with_native_svg

__all__ = [
    'main',
    'convert_svg_to_slide_shapes',
    'create_pptx_with_native_svg',
]
