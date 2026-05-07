"""Vision-language model analysis for PPT layout understanding.

Sends slide screenshots to a VL model (GPT-4o, Claude, Gemini, Ollama) to
produce structured descriptions of page layout, visual regions, design
patterns, and presentation logic.

Supports two analysis modes:
  1. layout: Describe page structure, regions, elements, layout sub-type
  2. logic: Analyze narrative flow, sequencing, storytelling patterns

Output is structured JSON consumed by the enhanced SpecExtractor.
"""

from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass
from pathlib import Path

from ppt_skill.spec.spec_model import (
    LayoutSubType,
    PageType,
    Region,
    VLModelConfig,
)

# ── Layout analysis prompt ───────────────────────────────────────────


LAYOUT_ANALYSIS_PROMPT = """Analyze this presentation slide image. Describe IN DETAIL:

1. **Page Type**: What kind of page is this?
   - cover: title slide (presentation title, subtitle, author, date)
   - toc: table of contents (section listing)
   - transition: section divider (section name/number)
   - content: body content slide
   - end_page: closing slide (thank you, contact, Q&A)

2. **Layout Structure** (for content pages):
   - left_right: text area on one side, visual on the other
   - top_bottom: heading at top, content below
   - left_middle_right: three columns
   - full_width: single full-width content area
   - grid: tiled cards or grid items
   - image_left: prominent image on left, text on right
   - image_right: prominent image on right, text on left
   - quote: centered quote or key statement
   - chart: data visualization dominant

3. **Regions**: List each distinct visual region with:
   - Approximate position (x, y, width, height as percentages 0-100)
   - Role: "title", "body", "image", "chart", "table", "decoration", "footer", "logo"
   - Content type: "text", "image", "chart", "table", "icon", "shape"

4. **Text Areas**: Where is the title? Where is the body text? Where is the footer/page number?

5. **Visual Style**:
   - background: color (HEX if possible), gradient, or image
   - are there decorative elements? (lines, shapes, icons)
   - color accents used
   - font hierarchy observed (title size, body size)

6. **Layout Pattern Description**: Describe the overall layout arrangement in 2-3 natural
   language sentences. This will be used as a design reference for generating similar pages.

IMPORTANT: Return your analysis as a JSON object with these exact keys:
{
  "page_type": "cover|toc|transition|content|end_page",
  "layout_sub_type": "left_right|top_bottom|left_middle_right|full_width|grid|image_left|image_right|quote|chart|custom",
  "regions": [
    {"x": 10, "y": 5, "width": 80, "height": 15, "role": "title", "content_type": "text", "description": "..."}
  ],
  "title_position": {"x": 0, "y": 0, "width": 0, "height": 0},
  "body_position": {"x": 0, "y": 0, "width": 0, "height": 0},
  "background_description": "...",
  "decoration_description": "...",
  "layout_description": "A natural language description of the layout arrangement..."
}

Use percentages (0-100) for all positions. Include all regions visible in the slide.
"""

# ── Logic analysis prompt ────────────────────────────────────────────


LOGIC_ANALYSIS_PROMPT = """Analyze the presentation logic across these {slide_count} slides.

For each slide, I'll provide: page_type, content_summary, density.

Based on the page type sequence, analyze:

1. **Section Structure**: How is the presentation organized into sections?
   Group slides into logical sections with names.

2. **Narrative Pattern**: What story structure does this follow?
   - problem_solution: present problem → propose solution → evidence → conclusion
   - chronological: timeline-based progression
   - pyramid: broad intro → narrow detail → broad conclusion
   - modular: self-contained sections
   - compare_contrast: A vs B comparison

3. **Story Arc**: Distribution across narrative stages.
   {opening: N, buildup: N, climax: N, resolution: N}

4. **Transition Style**: How does one slide flow into the next?
   - smooth: gradual topic progression
   - sectioned: clear section breaks
   - varied: mixed pacing

5. **Density Rhythm**: Note the pattern of dense vs breathing pages.
   Is there a deliberate rhythm?

Return JSON:
{
  "sections": [{"name": "...", "slides": [0,1,2], "page_types": [...]}],
  "narrative_pattern": "problem_solution|chronological|pyramid|modular",
  "story_arc": {"opening": 0, "buildup": 0, "climax": 0, "resolution": 0},
  "transition_style": "smooth|sectioned|varied",
  "density_observation": "..."
}
"""

# ── VL client ────────────────────────────────────────────────────────


@dataclass
class VLAnalysisResult:
    page_type: str = ""
    layout_sub_type: str = ""
    regions: list[dict] = None
    layout_description: str = ""
    background_description: str = ""
    decoration_description: str = ""
    raw_response: str = ""

    def __post_init__(self):
        if self.regions is None:
            self.regions = []


class VLClient:
    """Client for vision-language model API calls."""

    def __init__(self, config: VLModelConfig):
        self.config = config

    def analyze_layout(self, image_path: Path) -> VLAnalysisResult:
        """Send a slide image to VL model for layout analysis."""
        if not self.config.enabled:
            return VLAnalysisResult(
                layout_description="VL analysis disabled"
            )

        if self.config.provider == "openai":
            return self._analyze_openai(image_path, LAYOUT_ANALYSIS_PROMPT)
        elif self.config.provider == "anthropic":
            return self._analyze_anthropic(image_path, LAYOUT_ANALYSIS_PROMPT)
        elif self.config.provider == "gemini":
            return self._analyze_gemini(image_path, LAYOUT_ANALYSIS_PROMPT)
        elif self.config.provider == "ollama":
            return self._analyze_ollama(image_path, LAYOUT_ANALYSIS_PROMPT)
        else:
            return VLAnalysisResult(
                layout_description=f"Unknown provider: {self.config.provider}"
            )

    def analyze_logic(self, page_summaries: list[dict]) -> dict:
        """Analyze presentation logic from page summaries."""
        prompt = LOGIC_ANALYSIS_PROMPT.format(slide_count=len(page_summaries))

        # Build page context for the prompt
        page_context = "Page sequence:\n"
        for i, p in enumerate(page_summaries):
            page_context += (
                f"  Slide {i + 1}: type={p.get('page_type', '?')}, "
                f"summary={p.get('summary', '')}, "
                f"density={p.get('density', 'dense')}\n"
            )
        prompt += "\n" + page_context

        if self.config.provider == "openai":
            return self._analyze_text_openai(prompt)
        elif self.config.provider == "anthropic":
            return self._analyze_text_anthropic(prompt)
        else:
            return {"sections": [], "narrative_pattern": ""}

    # ── OpenAI ──────────────────────────────────────────────────────

    def _analyze_openai(self, image_path: Path, prompt: str) -> VLAnalysisResult:
        try:
            from openai import OpenAI
            import base64

            client = OpenAI(
                api_key=self.config.api_key or os.environ.get("OPENAI_API_KEY"),
                base_url=self.config.api_base or None,
            )

            with open(image_path, "rb") as f:
                image_data = base64.b64encode(f.read()).decode("utf-8")
            mime = "image/png" if image_path.suffix == ".png" else "image/jpeg"

            response = client.chat.completions.create(
                model=self.config.model,
                messages=[{
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {"type": "image_url", "image_url": {
                            "url": f"data:{mime};base64,{image_data}",
                            "detail": "high"
                        }},
                    ],
                }],
                max_tokens=self.config.max_tokens,
                temperature=self.config.temperature,
            )
            text = response.choices[0].message.content
            return self._parse_layout_response(text)
        except ImportError:
            return VLAnalysisResult(
                layout_description="openai package not installed. pip install openai"
            )
        except Exception as e:
            return VLAnalysisResult(
                layout_description=f"VL analysis error: {e}"
            )

    def _analyze_text_openai(self, prompt: str) -> dict:
        try:
            from openai import OpenAI
            client = OpenAI(
                api_key=self.config.api_key or os.environ.get("OPENAI_API_KEY"),
                base_url=self.config.api_base or None,
            )
            response = client.chat.completions.create(
                model=self.config.model,
                messages=[{"role": "user", "content": prompt}],
                max_tokens=self.config.max_tokens,
                temperature=self.config.temperature,
            )
            return _parse_json_response(response.choices[0].message.content)
        except Exception:
            return {}

    # ── Anthropic ────────────────────────────────────────────────────

    def _analyze_anthropic(self, image_path: Path, prompt: str) -> VLAnalysisResult:
        try:
            from anthropic import Anthropic
            import base64

            client = Anthropic(
                api_key=self.config.api_key or os.environ.get("ANTHROPIC_API_KEY"),
            )

            with open(image_path, "rb") as f:
                image_data = base64.b64encode(f.read()).decode("utf-8")
            mime = "image/png" if image_path.suffix == ".png" else "image/jpeg"

            response = client.messages.create(
                model=self.config.model,
                max_tokens=self.config.max_tokens,
                messages=[{
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {"type": "image", "source": {
                            "type": "base64",
                            "media_type": mime,
                            "data": image_data,
                        }},
                    ],
                }],
            )
            return self._parse_layout_response(response.content[0].text)
        except ImportError:
            return VLAnalysisResult(
                layout_description="anthropic package not installed. pip install anthropic"
            )
        except Exception as e:
            return VLAnalysisResult(layout_description=f"VL analysis error: {e}")

    def _analyze_text_anthropic(self, prompt: str) -> dict:
        try:
            from anthropic import Anthropic
            client = Anthropic(
                api_key=self.config.api_key or os.environ.get("ANTHROPIC_API_KEY"),
            )
            response = client.messages.create(
                model=self.config.model,
                max_tokens=self.config.max_tokens,
                messages=[{"role": "user", "content": prompt}],
            )
            return _parse_json_response(response.content[0].text)
        except Exception:
            return {}

    # ── Gemini ───────────────────────────────────────────────────────

    def _analyze_gemini(self, image_path: Path, prompt: str) -> VLAnalysisResult:
        try:
            import google.generativeai as genai
            genai.configure(api_key=self.config.api_key or os.environ.get("GEMINI_API_KEY"))

            import PIL.Image
            image = PIL.Image.open(image_path)

            model = genai.GenerativeModel(self.config.model)
            response = model.generate_content([prompt, image])
            return self._parse_layout_response(response.text)
        except ImportError:
            return VLAnalysisResult(
                layout_description="google-generativeai not installed. pip install google-generativeai"
            )
        except Exception as e:
            return VLAnalysisResult(layout_description=f"VL analysis error: {e}")

    # ── Ollama ───────────────────────────────────────────────────────

    def _analyze_ollama(self, image_path: Path, prompt: str) -> VLAnalysisResult:
        try:
            from openai import OpenAI
            import base64

            client = OpenAI(
                base_url=self.config.api_base or "http://localhost:11434/v1",
                api_key="ollama",
            )
            with open(image_path, "rb") as f:
                image_data = base64.b64encode(f.read()).decode("utf-8")

            response = client.chat.completions.create(
                model=self.config.model,
                messages=[{
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {"type": "image_url", "image_url": {
                            "url": f"data:image/png;base64,{image_data}"
                        }},
                    ],
                }],
                max_tokens=self.config.max_tokens,
            )
            return self._parse_layout_response(response.choices[0].message.content)
        except Exception as e:
            return VLAnalysisResult(layout_description=f"VL analysis error: {e}")

    # ── Response parsing ─────────────────────────────────────────────

    def _parse_layout_response(self, text: str) -> VLAnalysisResult:
        """Parse VL model response into structured VLAnalysisResult."""
        data = _parse_json_response(text)

        return VLAnalysisResult(
            page_type=data.get("page_type", "content"),
            layout_sub_type=data.get("layout_sub_type", "full_width"),
            regions=data.get("regions", []),
            layout_description=data.get("layout_description", ""),
            background_description=data.get("background_description", ""),
            decoration_description=data.get("decoration_description", ""),
            raw_response=text,
        )


# ── JSON parsing helpers ─────────────────────────────────────────────


def _parse_json_response(text: str) -> dict:
    """Try to parse JSON from model response (may be wrapped in markdown)."""
    if not text:
        return {}

    # Try direct parse
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Try extracting from markdown code block
    match = re.search(r'```(?:json)?\s*\n?(.*?)\n?```', text, re.DOTALL)
    if match:
        try:
            return json.loads(match.group(1))
        except json.JSONDecodeError:
            pass

    # Try finding JSON object boundaries
    for start, end in [("{", "}")]:
        si = text.find(start)
        if si >= 0:
            depth = 0
            for i in range(si, len(text)):
                if text[i] == start:
                    depth += 1
                elif text[i] == end:
                    depth -= 1
                    if depth == 0:
                        try:
                            return json.loads(text[si:i + 1])
                        except json.JSONDecodeError:
                            pass
                        break

    return {}
