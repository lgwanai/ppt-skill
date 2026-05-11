"""Content outline generator using prompt-ppt-content.md principles.

Uses agent loop to generate professional PPT content outline following WPS model:
- W (What): Navigation label
- P (Point): Core conclusion statement
- S (Support): Supporting evidence

Supports two modes:
  1. Topic-only mode: Input is just a topic name -> signals outer agent to
     call deerflow-skill for research, then generates outline with enriched content.
  2. Rich-content mode: Input has sufficient detail -> generates outline directly.

Usage:
    generator = OutlineGenerator()
    # Detect if input needs research delegation:
    result = generator.assess_and_delegate("AI in Healthcare 2024")
    # Or generate directly from rich content:
    outline = generator.generate(prompt_ppt_content_rules + enriched_content)
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml


# ---------------------------------------------------------------------------
# Sufficiency detection
# ---------------------------------------------------------------------------


def is_topic_only(user_input: str) -> bool:
    """Detect if input is just a topic name (vs detailed content).

    Returns True when the input looks like a bare topic that needs
    research before outline generation can proceed.

    Heuristics:
    - Structure markers (bullets, headers, numbered lists) indicate rich content
    - Word count <= 10 with no structure -> topic-only
    - Word count 11-30 with few lines and no structure -> topic-only
    """
    text = user_input.strip()
    if not text:
        return True

    words = text.split()
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    # Check for structure markers first — any presence means rich content
    structure_markers = ["- ", "* ", "1. ", "## ", "# ", "> ", "| ", "· ", "• "]
    for line in lines:
        for marker in structure_markers:
            if line.startswith(marker):
                return False

    # No structure markers found — assess by word count
    if len(words) <= 10:
        return True

    if len(words) <= 30 and len(lines) <= 3:
        return True

    return False


def build_delegation_signal(query: str) -> str:
    """Build a structured JSON signal telling the outer agent to call deerflow-skill.

    The outer agent (Trae orchestrator) parses this signal and delegates
    content research to deerflow-skill, then returns enriched content to
    ppt-skill for outline generation.
    """
    signal = {
        "type": "delegate",
        "skill": "deerflow-skill",
        "query": query,
        "context": (
            "Research this presentation topic comprehensively. "
            "Gather key sections, trends, data points, examples, case studies, "
            "and supporting evidence needed for a professional slide deck. "
            "The content will be used to generate a PPT outline following "
            "the WPS (What/Point/Support) model."
        ),
    }
    return json.dumps(signal, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# Data models
# ---------------------------------------------------------------------------


@dataclass
class SlideOutline:
    """Single slide content outline with WPS structure."""

    slide_number: int
    title: str
    w: str  # Navigation/section label
    p: str  # Core point/conclusion
    s: list[str]  # Supporting evidence
    page_type: str = "content"  # cover, toc, transition, content, emphasis, end
    section_name: str = ""
    notes: str = ""

    def to_dict(self) -> dict:
        return {
            "slide_number": self.slide_number,
            "title": self.title,
            "w": self.w,
            "p": self.p,
            "s": self.s,
            "page_type": self.page_type,
            "section_name": self.section_name,
            "notes": self.notes,
        }

    @classmethod
    def from_dict(cls, data: dict) -> SlideOutline:
        return cls(
            slide_number=data.get("slide_number", 0),
            title=data.get("title", ""),
            w=data.get("w", ""),
            p=data.get("p", ""),
            s=data.get("s", []),
            page_type=data.get("page_type", "content"),
            section_name=data.get("section_name", ""),
            notes=data.get("notes", ""),
        )


@dataclass
class ContentOutlineResult:
    """Complete content outline result."""

    presentation_title: str
    slides: list[SlideOutline]
    sections: list[str] = field(default_factory=list)
    metadata: dict = field(default_factory=dict)

    def to_markdown(self) -> str:
        """Convert outline to markdown (no W/P/S labels, follows ##/### rules).

        ## [section] = chapter divider (NOT a page)
        ### [title] = actual page
        ### 转场页：[title] = transition page
        ### 强调页：[title] = emphasis page
        ## 封面 / ## 结尾页 = special pages
        """
        lines = [f"# {self.presentation_title}", ""]

        # Cover
        cover = next((s for s in self.slides if s.page_type == "cover"), None)
        if cover:
            lines.append("## 封面")
            lines.append(f"- 主标题：{cover.title}")
            if cover.p:
                lines.append(f"- 副标题：{cover.p}")
            if cover.notes:
                lines.append(f"- {cover.notes}")
            lines.append("")

        # TOC
        toc = next((s for s in self.slides if s.page_type == "toc"), None)
        if toc and self.sections:
            lines.append("## 目录")
            for i, section in enumerate(self.sections, 1):
                lines.append(f"{i}. {section}")
            lines.append("")

        lines.append("---")
        lines.append("")

        # Sections
        current_section = ""
        for slide in self.slides:
            # Section divider (only on first slide of each section)
            if slide.section_name and slide.section_name != current_section:
                current_section = slide.section_name
                lines.append(f"## {current_section}")
                lines.append("")

            if slide.page_type == "transition":
                lines.append(f"### 转场页：{slide.title}")
                lines.append("")
            elif slide.page_type == "emphasis":
                lines.append(f"### 强调页：{slide.p or slide.title}")
                lines.append("")
            elif slide.page_type == "content":
                lines.append(f"### {slide.title}")
                if slide.p:
                    lines.append(slide.p)
                    lines.append("")
                for s_item in slide.s:
                    lines.append(f"- {s_item}")
                lines.append("")
            elif slide.page_type == "end":
                lines.append("## 结尾页")
                lines.append(f"- 致谢：{slide.title}")
                if slide.notes:
                    lines.append(f"- {slide.notes}")
                lines.append("")

        return "\n".join(lines)

    def to_yaml(self) -> str:
        """Convert outline to YAML format."""
        data = {
            "presentation_title": self.presentation_title,
            "sections": self.sections,
            "slides": [s.to_dict() for s in self.slides],
            "metadata": self.metadata,
        }
        return yaml.dump(data, default_flow_style=False, sort_keys=False, allow_unicode=True)

    @classmethod
    def from_yaml(cls, yaml_str: str) -> ContentOutlineResult:
        """Load outline from YAML string."""
        data = yaml.safe_load(yaml_str) or {}
        return cls(
            presentation_title=data.get("presentation_title", ""),
            sections=data.get("sections", []),
            slides=[SlideOutline.from_dict(s) for s in data.get("slides", [])],
            metadata=data.get("metadata", {}),
        )


class OutlineGenerator:
    """Generate PPT content outline using agent loop and content principles."""

    def __init__(self):
        self.content_prompt = self._load_content_prompt()
        self.max_iterations = 5

    def _load_content_prompt(self) -> str:
        """Load prompt-ppt-content.md content."""
        prompt_path = Path("references/prompt-ppt-content.md")
        if prompt_path.exists():
            return prompt_path.read_text(encoding="utf-8")
        return ""

    def assess_and_delegate(self, user_input: str) -> str:
        """Assess input and either generate outline or signal deerflow-skill delegation.

        Returns:
            If input is topic-only: JSON delegation signal for outer agent.
            If input is sufficient: markdown outline string.
        """
        if is_topic_only(user_input):
            return build_delegation_signal(user_input)

        outline = self.generate(user_input)
        return outline.to_markdown()

    def generate(
        self,
        user_input: str,
        mode: str = "auto",
        questions_callback: Any = None,
    ) -> ContentOutlineResult:
        """Generate content outline from user input.

        Args:
            user_input: Topic, article, or existing outline
            mode: "auto" (ask questions if needed), "skip_questions" (direct generation)
            questions_callback: Function to collect answers (for interactive mode)

        Returns:
            ContentOutlineResult with slides following WPS model
        """
        # The actual LLM-powered outline generation happens in the outer agent.
        # This method provides the programmatic placeholder for CLI/test use.

        return self._generate_placeholder(user_input)

    def build_content_prompt(self, user_input: str) -> str:
        """Build a structured prompt that combines user input with content principles.

        The outer agent uses this prompt to execute the LLM agent loop:
        1. Apply prompt-ppt-content.md WPS rules
        2. Structure each slide with W (What), P (Point), S (Support)
        3. Validate and optimize
        """
        return (
            f"## User Input\n\n{user_input}\n\n"
            f"## Content Principles\n\n{self.content_prompt}\n\n"
            f"Generate a professional PPT outline following the WPS model above."
        )

    def _generate_placeholder(self, user_input: str) -> ContentOutlineResult:
        """Generate minimal outline for programmatic use."""
        lines = [l.strip() for l in user_input.strip().splitlines() if l.strip()]
        title = lines[0] if lines else "Presentation"

        return ContentOutlineResult(
            presentation_title=title,
            sections=["Section 1"],
            slides=[
                SlideOutline(
                    slide_number=1,
                    title=title,
                    w="Cover",
                    p="",
                    s=[],
                    page_type="cover",
                ),
                SlideOutline(
                    slide_number=2,
                    title="Table of Contents",
                    w="Navigation",
                    p="Overview of presentation",
                    s=["Section 1"],
                    page_type="toc",
                ),
                SlideOutline(
                    slide_number=3,
                    title="Section 1",
                    w="Section 1",
                    p="Key point about the topic",
                    s=["Supporting evidence 1", "Supporting evidence 2"],
                    page_type="content",
                    section_name="Section 1",
                ),
                SlideOutline(
                    slide_number=4,
                    title="Thank You",
                    w="End",
                    p="",
                    s=[],
                    page_type="end",
                ),
            ],
            metadata={"mode": "placeholder"},
        )

    def save(self, outline: ContentOutlineResult, path: Path | str) -> Path:
        """Save outline to file."""
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)

        content = outline.to_yaml() if path.suffix == ".yaml" else outline.to_markdown()
        path.write_text(content, encoding="utf-8")
        return path

    @staticmethod
    def load(path: Path | str) -> ContentOutlineResult:
        """Load outline from markdown file (##/### format)."""
        path = Path(path)
        content = path.read_text(encoding="utf-8")

        if path.suffix in (".yaml", ".yml"):
            return ContentOutlineResult.from_yaml(content)

        lines = content.splitlines()
        title = ""
        slides: list[SlideOutline] = []
        sections: list[str] = []
        current_section = ""

        i = 0
        while i < len(lines):
            line = lines[i].strip()

            if line.startswith("# "):
                title = line[2:].strip()
            elif line.startswith("## 封面"):
                # Find body lines after ## 封面
                i += 1
                bt = title; bp = ""; bn = ""
                while i < len(lines) and lines[i].strip() and not lines[i].strip().startswith("## "):
                    t = lines[i].strip()
                    if t.startswith("- 主标题："): bt = t.replace("- 主标题：", "").strip()
                    elif t.startswith("- 副标题："): bp = t.replace("- 副标题：", "").strip()
                    elif t.startswith("- 演讲者") or t.startswith("- 日期"): bn = t[2:].strip()
                    elif t.startswith("- "): bn = t[2:].strip()
                    i += 1
                slides.append(SlideOutline(0, bt, "", bp, [], "cover", notes=bn))
                continue
            elif line.startswith("## 目录") or line.startswith("## 结尾页"):
                # TOC / End
                pt = "toc" if "目录" in line else "end"
                i += 1
                st = title
                sn = ""
                while i < len(lines) and lines[i].strip() and not lines[i].strip().startswith("## "):
                    t = lines[i].strip()
                    if t.startswith("- "): sn = t[2:].strip()
                    i += 1
                slides.append(SlideOutline(0, st, "", "", [], pt, notes=sn))
                continue
            elif line.startswith("## ") and not line.startswith("## 封") and not line.startswith("## 目") and not line.startswith("## 结"):
                current_section = line[3:].strip()
                sections.append(current_section)
            elif line.startswith("### 转场页"):
                st = line.replace("### 转场页：", "").replace("### 转场页", "").strip()
                i += 1
                body = []
                while i < len(lines) and lines[i].strip() and not lines[i].strip().startswith("### ") and not lines[i].strip().startswith("## "):
                    t = lines[i].strip()
                    if not t.startswith("---"): body.append(t)
                    i += 1
                slides.append(SlideOutline(len(slides)+1, st, current_section, "", body, "transition", current_section))
                continue
            elif line.startswith("### 强调页"):
                st = line.replace("### 强调页：", "").replace("### 强调页", "").strip()
                slides.append(SlideOutline(len(slides)+1, st, current_section, "", [], "emphasis", current_section))
            elif line.startswith("### "):
                st = line[4:].strip()
                i += 1
                body = []; core = ""
                while i < len(lines) and lines[i].strip() and not lines[i].strip().startswith("### ") and not lines[i].strip().startswith("## "):
                    t = lines[i].strip()
                    if t.startswith("- "): body.append(t[2:].strip())
                    elif t and not t.startswith("---") and not core: core = t
                    elif t and not t.startswith("---"): body.append(t)
                    i += 1
                slides.append(SlideOutline(len(slides)+1, st, current_section, core, body, "content", current_section))
                continue
            i += 1

        sections = list(dict.fromkeys(s for s in sections if s))
        return ContentOutlineResult(title, slides, sections)


__all__ = [
    "OutlineGenerator", "ContentOutlineResult", "SlideOutline",
    "is_topic_only", "build_delegation_signal",
]
