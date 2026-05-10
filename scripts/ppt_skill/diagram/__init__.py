"""Diagram generation module for PPT skill.

Provides SVG diagram generation for:
- Architecture diagrams
- Data flow diagrams
- Flowcharts
- Sequence diagrams
- Class diagrams (UML)
- ER diagrams
- Network topology
- Timeline/Gantt charts
"""

from ppt_skill.diagram.generator import (
    build_svg,
    parse_style,
    STYLE_PROFILES,
    DEFAULT_VIEWBOX,
)
from ppt_skill.diagram.diagram_generator import (
    DiagramGenerator,
    DiagramNode,
    DiagramArrow,
    DiagramContainer,
    DiagramResult,
)

__all__ = [
    # Low-level API
    "build_svg",
    "parse_style",
    "STYLE_PROFILES",
    "DEFAULT_VIEWBOX",
    # High-level API
    "DiagramGenerator",
    "DiagramNode",
    "DiagramArrow",
    "DiagramContainer",
    "DiagramResult",
]
