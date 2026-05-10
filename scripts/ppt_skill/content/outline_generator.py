"""Content outline generator using prompt-ppt-content.md principles.

Uses agent loop to generate professional PPT content outline following WPS model:
- W (What): Navigation label
- P (Point): Core conclusion statement
- S (Support): Supporting evidence

Usage:
    generator = OutlineGenerator()
    outline = generator.generate("AI in Healthcare 2024")
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml


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
        """Convert outline to markdown format."""
        lines = [f"# {self.presentation_title}", ""]

        # Cover
        cover = next((s for s in self.slides if s.page_type == "cover"), None)
        if cover:
            lines.append("## Cover")
            lines.append(f"- Main title: {cover.title}")
            lines.append(f"- Subtitle: {cover.p}")
            lines.append(f"- Notes: {cover.notes}")
            lines.append("")

        # TOC
        toc = next((s for s in self.slides if s.page_type == "toc"), None)
        if toc:
            lines.append("## Table of Contents")
            for i, section in enumerate(self.sections, 1):
                lines.append(f"{i}. {section}")
            lines.append("")

        lines.append("---")
        lines.append("")

        # Sections
        current_section = ""
        for slide in self.slides:
            if slide.page_type in ("transition", "content"):
                if slide.section_name and slide.section_name != current_section:
                    current_section = slide.section_name
                    lines.append(f"## {current_section}")
                    lines.append("")

            if slide.page_type == "transition":
                lines.append(f"### Transition: {slide.title}")
                lines.append("")
            elif slide.page_type == "content":
                lines.append(f"### Slide {slide.slide_number}: {slide.title}")
                lines.append(f"**W (Navigation)**: {slide.w}")
                lines.append(f"**P (Core Point)**: {slide.p}")
                lines.append("**S (Support)**:")
                for s_item in slide.s:
                    lines.append(f"- {s_item}")
                lines.append("")
            elif slide.page_type == "emphasis":
                lines.append(f"### Emphasis: {slide.p}")
                lines.append("")
            elif slide.page_type == "end":
                lines.append("## End")
                lines.append(f"- {slide.title}")
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
        # This is a placeholder - the actual LLM generation happens in the skill
        # The skill uses the agent loop to:
        # 1. Assess content sufficiency
        # 2. Ask questions if needed
        # 3. Generate outline using prompt-ppt-content.md principles
        # 4. Validate and fix issues

        return self._generate_placeholder(user_input)

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
        """Load outline from file."""
        path = Path(path)
        content = path.read_text(encoding="utf-8")

        if path.suffix in (".yaml", ".yml"):
            return ContentOutlineResult.from_yaml(content)

        # Parse markdown format
        lines = content.splitlines()
        title = ""
        slides: list[SlideOutline] = []
        current_section = ""

        i = 0
        while i < len(lines):
            line = lines[i].strip()

            if line.startswith("# "):
                title = line[2:].strip()
            elif line.startswith("## "):
                current_section = line[3:].strip()
            elif line.startswith("### Slide "):
                # Parse slide
                slide_num = int(line.split()[2].rstrip(":"))
                slide_title = line.split(":", 1)[1].strip() if ":" in line else ""
                i += 1
                w = p = ""
                s: list[str] = []
                while i < len(lines) and lines[i].strip() and not lines[i].strip().startswith("###"):
                    part = lines[i].strip()
                    if part.startswith("**W"):
                        w = part.split("**: ", 1)[1] if "**: " in part else ""
                    elif part.startswith("**P"):
                        p = part.split("**: ", 1)[1] if "**: " in part else ""
                    elif part.startswith("**S"):
                        pass
                    elif part.startswith("- "):
                        s.append(part[2:])
                    i += 1
                slides.append(SlideOutline(
                    slide_number=slide_num,
                    title=slide_title,
                    w=w,
                    p=p,
                    s=s,
                    page_type="content",
                    section_name=current_section,
                ))
                continue
            i += 1

        sections = list(dict.fromkeys(s.section_name for s in slides if s.section_name))

        return ContentOutlineResult(
            presentation_title=title,
            sections=sections,
            slides=slides,
        )


__all__ = ["OutlineGenerator", "ContentOutlineResult", "SlideOutline"]
