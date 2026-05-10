"""Diagram generator wrapper for PPT integration.

Provides high-level API for generating SVG diagrams to embed in PPT slides.

Usage:
    generator = DiagramGenerator()
    svg = generator.generate_architecture({
        "title": "System Architecture",
        "nodes": [...],
        "arrows": [...],
    })
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional

from ppt_skill.diagram.generator import (
    build_svg,
    parse_style,
    STYLE_PROFILES,
    DEFAULT_VIEWBOX,
)


@dataclass
class DiagramNode:
    """Node in a diagram."""

    id: str
    label: str
    x: float
    y: float
    width: float = 180
    height: float = 76
    kind: str = "rect"  # rect, cylinder, hexagon, document, folder, terminal
    sublabel: str = ""
    type_label: str = ""
    fill: str = ""
    stroke: str = ""

    def to_dict(self) -> dict:
        result = {
            "id": self.id,
            "label": self.label,
            "x": self.x,
            "y": self.y,
            "width": self.width,
            "height": self.height,
            "kind": self.kind,
        }
        if self.sublabel:
            result["sublabel"] = self.sublabel
        if self.type_label:
            result["type_label"] = self.type_label
        if self.fill:
            result["fill"] = self.fill
        if self.stroke:
            result["stroke"] = self.stroke
        return result


@dataclass
class DiagramArrow:
    """Arrow connecting nodes."""

    source: str
    target: str
    label: str = ""
    flow: str = "control"  # control, write, read, data, async, feedback
    source_port: str = ""
    target_port: str = ""

    def to_dict(self) -> dict:
        result = {
            "source": self.source,
            "target": self.target,
            "flow": self.flow,
        }
        if self.label:
            result["label"] = self.label
        if self.source_port:
            result["source_port"] = self.source_port
        if self.target_port:
            result["target_port"] = self.target_port
        return result


@dataclass
class DiagramContainer:
    """Container/grouping box in a diagram."""

    x: float
    y: float
    width: float
    height: float
    label: str = ""
    subtitle: str = ""

    def to_dict(self) -> dict:
        result = {
            "x": self.x,
            "y": self.y,
            "width": self.width,
            "height": self.height,
        }
        if self.label:
            result["label"] = self.label
        if self.subtitle:
            result["subtitle"] = self.subtitle
        return result


@dataclass
class DiagramResult:
    """Result of diagram generation."""

    svg: str
    width: float
    height: float
    style: int = 1
    diagram_type: str = "architecture"


class DiagramGenerator:
    """Generate SVG diagrams for PPT slides."""

    def __init__(self, style: int = 1):
        self.style = style

    def generate(
        self,
        diagram_type: str,
        title: str,
        nodes: List[DiagramNode],
        arrows: List[DiagramArrow],
        containers: List[DiagramContainer] = None,
        subtitle: str = "",
        legend: List[dict] = None,
        style: int = None,
    ) -> DiagramResult:
        """Generate an SVG diagram.

        Args:
            diagram_type: architecture, data-flow, flowchart, sequence, etc.
            title: Diagram title
            nodes: List of DiagramNode objects
            arrows: List of DiagramArrow objects
            containers: Optional grouping containers
            subtitle: Optional subtitle
            legend: Optional legend entries
            style: Style number (1-7), defaults to self.style

        Returns:
            DiagramResult with SVG content
        """
        style_index = style or self.style
        width, height = DEFAULT_VIEWBOX.get(diagram_type, (960, 600))

        data = {
            "title": title,
            "style": style_index,
            "nodes": [n.to_dict() for n in nodes],
            "arrows": [a.to_dict() for a in arrows],
        }

        if subtitle:
            data["subtitle"] = subtitle
        if containers:
            data["containers"] = [c.to_dict() for c in containers]
        if legend:
            data["legend"] = legend

        svg = build_svg(diagram_type, data)
        return DiagramResult(
            svg=svg,
            width=width,
            height=height,
            style=style_index,
            diagram_type=diagram_type,
        )

    def generate_architecture(
        self,
        title: str,
        layers: List[Dict[str, Any]],
        style: int = None,
    ) -> DiagramResult:
        """Generate an architecture diagram from layer structure.

        Args:
            title: Diagram title
            layers: List of layer dicts with 'name' and 'components'
            style: Style number (1-7)

        Returns:
            DiagramResult with SVG content
        """
        nodes: List[DiagramNode] = []
        arrows: List[DiagramArrow] = []
        containers: List[DiagramContainer] = []

        y_offset = 120
        layer_height = 100
        component_width = 160
        gap = 40

        for layer_idx, layer in enumerate(layers):
            layer_name = layer.get("name", f"Layer {layer_idx + 1}")
            components = layer.get("components", [])

            # Container for layer
            container_width = len(components) * (component_width + gap) + gap
            containers.append(DiagramContainer(
                x=40,
                y=y_offset - 20,
                width=container_width,
                height=layer_height + 40,
                label=layer_name,
            ))

            # Components in layer
            x_offset = 80
            for comp_idx, comp in enumerate(components):
                comp_id = f"node-{layer_idx}-{comp_idx}"
                comp_name = comp.get("name", f"Component {comp_idx + 1}")
                comp_type = comp.get("type", "")

                nodes.append(DiagramNode(
                    id=comp_id,
                    label=comp_name,
                    x=x_offset,
                    y=y_offset,
                    width=component_width,
                    height=layer_height - 20,
                    type_label=comp_type,
                ))

                # Arrow to previous layer
                if layer_idx > 0 and comp_idx < len(layers[layer_idx - 1].get("components", [])):
                    prev_id = f"node-{layer_idx - 1}-{comp_idx}"
                    arrows.append(DiagramArrow(
                        source=prev_id,
                        target=comp_id,
                        flow="control",
                    ))

                x_offset += component_width + gap

            y_offset += layer_height + 60

        return self.generate(
            diagram_type="architecture",
            title=title,
            nodes=nodes,
            arrows=arrows,
            containers=containers,
            style=style,
        )

    def generate_flowchart(
        self,
        title: str,
        steps: List[Dict[str, Any]],
        style: int = None,
    ) -> DiagramResult:
        """Generate a flowchart from step structure.

        Args:
            title: Diagram title
            steps: List of step dicts with 'label', 'type' (process/decision/io)
            style: Style number (1-7)

        Returns:
            DiagramResult with SVG content
        """
        nodes: List[DiagramNode] = []
        arrows: List[DiagramArrow] = []

        y_offset = 120
        step_height = 60
        step_width = 160
        gap = 80

        for step_idx, step in enumerate(steps):
            step_id = f"step-{step_idx}"
            step_label = step.get("label", f"Step {step_idx + 1}")
            step_type = step.get("type", "process")

            # Shape based on type
            kind = "rect"
            if step_type == "decision":
                kind = "hexagon"
            elif step_type == "io":
                kind = "document"

            nodes.append(DiagramNode(
                id=step_id,
                label=step_label,
                x=400 - step_width / 2,  # Center horizontally
                y=y_offset,
                width=step_width if step_type != "decision" else 120,
                height=step_height,
                kind=kind,
            ))

            # Arrow to next step
            if step_idx < len(steps) - 1:
                next_id = f"step-{step_idx + 1}"
                arrows.append(DiagramArrow(
                    source=step_id,
                    target=next_id,
                    flow="control",
                    source_port="bottom",
                    target_port="top",
                ))

            y_offset += step_height + gap

        return self.generate(
            diagram_type="flowchart",
            title=title,
            nodes=nodes,
            arrows=arrows,
            style=style,
        )

    def generate_sequence(
        self,
        title: str,
        participants: List[str],
        messages: List[Dict[str, Any]],
        style: int = None,
    ) -> DiagramResult:
        """Generate a sequence diagram.

        Args:
            title: Diagram title
            participants: List of participant names
            messages: List of message dicts with 'from', 'to', 'label'
            style: Style number (1-7)

        Returns:
            DiagramResult with SVG content
        """
        nodes: List[DiagramNode] = []
        arrows: List[DiagramArrow] = []

        # Participants as lifelines
        x_offset = 120
        lifeline_width = 120
        gap = 200

        for p_idx, participant in enumerate(participants):
            nodes.append(DiagramNode(
                id=f"participant-{p_idx}",
                label=participant,
                x=x_offset,
                y=80,
                width=lifeline_width,
                height=40,
                type_label="participant",
            ))
            x_offset += lifeline_width + gap

        # Messages
        y_offset = 150
        message_gap = 50

        for m_idx, message in enumerate(messages):
            from_idx = participants.index(message.get("from", ""))
            to_idx = participants.index(message.get("to", ""))
            label = message.get("label", "")

            from_x = 120 + from_idx * (lifeline_width + gap) + lifeline_width / 2
            to_x = 120 + to_idx * (lifeline_width + gap) + lifeline_width / 2

            arrows.append(DiagramArrow(
                source=f"participant-{from_idx}",
                target=f"participant-{to_idx}",
                label=label,
                flow="control",
                source_port="right" if from_idx < to_idx else "left",
                target_port="left" if from_idx < to_idx else "right",
            ))

            y_offset += message_gap

        return self.generate(
            diagram_type="sequence",
            title=title,
            nodes=nodes,
            arrows=arrows,
            style=style,
        )

    def save(self, result: DiagramResult, path: Path | str) -> Path:
        """Save diagram SVG to file."""
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(result.svg, encoding="utf-8")
        return path


__all__ = [
    "DiagramGenerator",
    "DiagramNode",
    "DiagramArrow",
    "DiagramContainer",
    "DiagramResult",
]