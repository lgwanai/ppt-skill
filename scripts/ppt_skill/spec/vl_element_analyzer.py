"""VL-based element analysis — sends slide image + element properties to VL model.

For each slide in a PPTX:
  1. Extract all element properties (position, size, fonts, colors, borders, backgrounds)
     using python-pptx
  2. Render the slide as a PNG image
  3. Send both to VL model with a structured prompt
  4. VL returns element role descriptions, group relationships, and page classification

The output JSON contains NO text content — only element attributes, roles,
and relationships. This becomes the "clone prototype" for reproducing the slide.

Usage:
    analyzer = VLElementAnalyzer()
    result = analyzer.analyze(slide_image_path, elements_data)
    for elem in result["elements"]:
        print(elem["role"], elem["description"])
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path

from ppt_skill.config import get_vl_client, get_vl_config


# ---------------------------------------------------------------------------
# VL Analysis Prompt
# ---------------------------------------------------------------------------

ELEMENT_ANALYSIS_PROMPT = """You are a PPT design spec analyzer. You will receive:
1. A slide screenshot image
2. A numbered list of all elements detected from the PPTX file, with their visual properties

Your task: Compare the IMAGE with the ELEMENT DATA, and for each element determine:
  - What VISUAL ROLE does this element play in the layout?
  - What TYPE of content does it contain or represent?
  - If multiple elements work together as a group, identify the group.

## How to analyze each element

Look at the screenshot and match each numbered element to what you see:
  - Is it a title? (large text, prominent position, usually at top)
  - Is it body text? (paragraphs, bullet points, multi-line text)
  - Is it a subtitle? (smaller text below the title)
  - Is it a background image? (full-slide image behind everything)
  - Is it a logo/icon? (small image in corner or margin, repeated across slides)
  - Is it a decorative element? (lines, shapes, color accents, dividers)
  - Is it a table? (grid of rows and columns)
  - Is it a chart? (data visualization)
  - Is it an image in the content area? (photo, screenshot, illustration)

## Rules
- Always compare the element data (position, size, font, colors) with what you
  actually see in the image. If a text element has large bold font at the top
  of the slide, it's a "title". If it has small text with bullets, it's "body".
- For layout-level images (from layout/master), identify if they are background
  (full-slide size) or decoration (small, corner/margin positions).
- Identify compound groups: e.g. "elements 3,4,5 form a process flowchart"
- Use consistent role terminology (see list above)
- DO NOT include text content in your output — describe visual role only

## Element Data (extracted from PPTX)

{element_context}

## Output Format

Return ONLY valid JSON with no markdown wrapping:

{{
  "layout_sub_type": "left_right|top_bottom|full_width|grid|image_left|image_right|quote|chart|custom",
  "background": {{
    "type": "solid|gradient|image|none",
    "color": "#HEX or empty",
    "description": "Visual description of background"
  }},
  "elements": [
    {{
      "id": 1,
      "role": "title",
      "type_category": "text|shape|image|table|chart|group|decoration",
      "description": "Slide title: large bold dark blue text, centered at top",
      "visual_weight": "high|medium|low",
      "style_notes": "Font color #0070C0, bold, centered alignment"
    }}
  ],
  "element_groups": [
    {{
      "element_ids": [3, 4, 5],
      "group_role": "process_flow|card|list|diagram|table|comparison|timeline|layout_grid|decoration",
      "description": "Three rounded rectangles with connecting arrows forming a horizontal process"
    }}
  ],
  "design_patterns": [
    "Consistent card pattern with shadow and rounded corners",
    "Two-column text layout with icon accents on left"
  ]
}}

IMPORTANT:
- Every element MUST have a role assigned. Do not skip any element.
- Base each element's role on what you see in the image matched to its properties.
- element_ids in groups MUST reference valid element IDs from the input.
- DO NOT include text content. Describe visual properties only.
- Do NOT classify page_type. Focus on element roles.
"""


# ---------------------------------------------------------------------------
# Element context builder
# ---------------------------------------------------------------------------


def _build_element_context(elements_data: list[dict], indent: str = "") -> str:
    """Build a formatted text description of all elements for the VL prompt.

    Shows hierarchy via indentation for nested children.
    """
    lines: list[str] = []
    for el in elements_data:
        eid = el.get("id", 0)
        etype = el.get("element_type", "unknown")
        pos = el.get("position", {})
        x = pos.get("x", 0)
        y = pos.get("y", 0)
        w = pos.get("w", 0)
        h = pos.get("h", 0)

        lines.append(f"{indent}Element #{eid}:")
        lines.append(f"{indent}  Type: {etype}")
        lines.append(f"{indent}  Position: x={x:.3f}, y={y:.3f}, w={w:.3f}, h={h:.3f} (normalized 0-1)")
        lines.append(f"{indent}  Z-order: {el.get('z_order', 0)}")

        # Show semantic role if available
        role = el.get("semantic_role")
        if role:
            lines.append(f"{indent}  Inferred role: {role}")

        text_style = el.get("text_style")
        shape_style = el.get("shape_style")

        if text_style:
            ts = text_style
            font_info = f"{ts.get('font_family', '?')}, {ts.get('font_size_pt', 0)}pt"
            if ts.get("font_weight") == "bold":
                font_info += ", bold"
            if ts.get("font_italic"):
                font_info += ", italic"
            lines.append(f"{indent}  Font: {font_info}")
            lines.append(f"{indent}  Font color: {ts.get('font_color', '?')}")
            lines.append(f"{indent}  Alignment: {ts.get('text_alignment', '?')}")

        if shape_style:
            ss = shape_style
            lines.append(f"{indent}  Shape type: {ss.get('shape_type', '?')}")
            if ss.get("fill_color"):
                lines.append(f"{indent}  Fill: {ss.get('fill_color')} (opacity: {ss.get('fill_opacity', 1.0)})")
            if ss.get("stroke_color"):
                lines.append(f"{indent}  Stroke: {ss.get('stroke_color')} {ss.get('stroke_width_pt', 0)}pt")
            if ss.get("corner_radius_pt", 0) > 0:
                lines.append(f"{indent}  Corner radius: {ss.get('corner_radius_pt')}pt")

        image_style = el.get("image_style")
        if image_style:
            lines.append(f"{indent}  Image dimensions: {image_style.get('original_width_px', '?')}x{image_style.get('original_height_px', '?')}px")

        # Layout-level image info
        if el.get("from_layout"):
            lines.append(f"{indent}  Source: {el['from_layout']} (layout/master inherited)")
        if el.get("is_background"):
            lines.append(f"{indent}  Full-slide background image")

        # Table info
        table_data = el.get("table", {})
        if table_data:
            lines.append(f"{indent}  Table: {table_data.get('rows', 0)} rows x {table_data.get('cols', 0)} columns")
            for cell in table_data.get("cells", [])[:5]:  # Show first 5 cells
                lines.append(f"{indent}    Cell[{cell.get('row',0)},{cell.get('col',0)}]: '{cell.get('text','')[:40]}'")
                if cell.get("text_style"):
                    cts = cell.get("text_style", {})
                    lines.append(f"{indent}      font={cts.get('font_family','')} {cts.get('font_size_pt','')}pt")
                if cell.get("fill_color"):
                    lines.append(f"{indent}      fill={cell.get('fill_color')}")

        # Children (recursive, indented)
        children = el.get("children", [])
        if children:
            lines.append(f"{indent}  Children ({len(children)} elements):")
            child_text = _build_element_context(children, indent + "    ")
            lines.append(child_text)

        lines.append("")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# VL Analysis
# ---------------------------------------------------------------------------


@dataclass
class ElementAnalysisResult:
    """Result of VL-based element analysis for one slide."""

    page_type: str = "content"
    layout_sub_type: str = "full_width"
    background: dict = field(default_factory=lambda: {"type": "solid", "color": "#FFFFFF", "description": ""})
    elements: list[dict] = field(default_factory=list)
    element_groups: list[dict] = field(default_factory=list)
    design_patterns: list[str] = field(default_factory=list)
    raw_response: str = ""


class VLElementAnalyzer:
    """Analyze slide elements using VL model to determine roles and relationships."""

    def __init__(self):
        self.config = get_vl_config()
        self.client = get_vl_client() if self.config.get("api_key") else None
        self.enabled = self.config.get("enabled", False) and bool(self.client)

    def analyze(
        self,
        slide_image: Path | str,
        elements_data: list[dict],
    ) -> ElementAnalysisResult:
        """Analyze a single slide's elements using VL model.

        Args:
            slide_image: Path to the rendered slide PNG image.
            elements_data: List of element dicts with properties (from extract_element).

        Returns:
            ElementAnalysisResult with VL-assigned roles and group descriptions.
        """
        if not self.enabled:
            return self._fallback_from_programmatic(elements_data)

        element_context = _build_element_context(elements_data)
        prompt = ELEMENT_ANALYSIS_PROMPT.format(element_context=element_context)

        try:
            result = self._call_vl(slide_image, prompt)
            return result
        except Exception as e:
            return self._fallback_from_programmatic(elements_data, str(e))

    def _call_vl(self, image_path: Path | str, prompt: str) -> ElementAnalysisResult:
        """Call the VL model with image + prompt."""
        import base64
        from openai import OpenAI

        image_path = Path(image_path)
        if not image_path.exists():
            return self._fallback_from_programmatic([], "Image not found")

        mime_type = "image/png" if image_path.suffix == ".png" else "image/jpeg"
        with open(image_path, "rb") as f:
            img_b64 = base64.b64encode(f.read()).decode("utf-8")

        client = self.client
        model = self.config.get("model", "gpt-4o")
        max_tokens = int(self.config.get("max_tokens", "4096"))

        response = client.chat.completions.create(
            model=model,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:{mime_type};base64,{img_b64}",
                            "detail": "high",
                        },
                    },
                ],
            }],
            max_tokens=max_tokens,
            temperature=0.1,
        )

        text = response.choices[0].message.content
        return self._parse_response(text)

    def _parse_response(self, text: str) -> ElementAnalysisResult:
        """Parse VL response JSON into ElementAnalysisResult."""
        data = _parse_json(text)
        if not data:
            return ElementAnalysisResult(raw_response=text)

        return ElementAnalysisResult(
            layout_sub_type=data.get("layout_sub_type", "full_width"),
            background=data.get("background", {"type": "solid", "color": "#FFFFFF", "description": ""}),
            elements=data.get("elements", []),
            element_groups=data.get("element_groups", []),
            design_patterns=data.get("design_patterns", []),
            raw_response=text,
        )

    def _fallback_from_programmatic(
        self,
        elements_data: list[dict],
        error: str = "",
    ) -> ElementAnalysisResult:
        """Fallback: use programmatic semantic roles when VL is unavailable.

        Uses element_type and position heuristics to infer role:
          - id=1000+ with is_background=true → background
          - id=1000+ from layout → decoration
          - element_type=table → table
          - semantic_role from programmatic inference
        """
        result = ElementAnalysisResult(
            raw_response=f"VL fallback: {error}" if error else "VL disabled"
        )

        for el in elements_data:
            eid = el.get("id", 0)
            elem_type = el.get("element_type", "shape")
            role = el.get("semantic_role", "body")

            # Better fallback heuristics
            if el.get("is_background"):
                role = "background"
            elif el.get("from_layout"):
                role = "decoration"
            elif eid >= 1000:
                role = role or "decoration"

            type_category = elem_type
            if elem_type == "table":
                type_category = "table"
            elif elem_type == "image":
                type_category = "image"
            elif elem_type == "group":
                type_category = "group"

            result.elements.append({
                "id": eid,
                "role": role,
                "type_category": type_category,
                "description": f"{role}: position=({el.get('position',{}).get('x',0):.2f},{el.get('position',{}).get('y',0):.2f}) size=({el.get('position',{}).get('w',0):.2f}x{el.get('position',{}).get('h',0):.2f})",
                "visual_weight": "high" if role == "title" else "medium",
                "style_notes": "",
            })

        return result


# ---------------------------------------------------------------------------
# JSON parsing helper
# ---------------------------------------------------------------------------


def _parse_json(text: str) -> dict | None:
    """Extract JSON from model response, handling markdown code blocks."""
    if not text:
        return None

    text = text.strip()

    # Try direct JSON parse
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Try extracting from markdown code fence
    match = re.search(r'```(?:json)?\s*\n(.*?)\n?```', text, re.DOTALL)
    if match:
        try:
            return json.loads(match.group(1).strip())
        except json.JSONDecodeError:
            pass

    # Try finding outermost braces
    brace_start = text.find("{")
    if brace_start >= 0:
        depth = 0
        for i in range(brace_start, len(text)):
            if text[i] == "{":
                depth += 1
            elif text[i] == "}":
                depth -= 1
                if depth == 0:
                    try:
                        return json.loads(text[brace_start : i + 1])
                    except json.JSONDecodeError:
                        pass
                    break

    return None


__all__ = ["VLElementAnalyzer", "ElementAnalysisResult", "ELEMENT_ANALYSIS_PROMPT"]
