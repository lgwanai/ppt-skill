"""ContentGatherer orchestrator — 3-phase pipeline: assess → question → generate.

Orchestrates the full content gathering pipeline: sufficiency assessment,
adaptive questioning (with 8-question budget), and outline generation.
Produces ContentOutline objects serialized as YAML artifacts in outlines/.

Usage::

    gatherer = ContentGatherer()
    outline = gatherer.gather("Make a presentation about Q3 results")
    path = gatherer.save("outlines")  # → outlines/q3-results.yaml

    # Skip questioning entirely
    outline = gatherer.gather(topic, mode="skip_questions")

    # Load a previously saved outline
    outline = ContentGatherer.load_outline("q3-results", "outlines")
"""

from __future__ import annotations

import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import yaml

from ppt_skill.content.model import (
    ContentOutline,
    OutlineLayoutType,
    QuestionSession,
    SlideEntry,
    SufficiencyResult,
)
from ppt_skill.content.questioning import (
    generate_gap_questions,
    generate_section_questions,
    identify_content_gaps,
)
from ppt_skill.content.sufficiency import assess_sufficiency
from ppt_skill.content.model import _dataclass_to_dict


# ---------------------------------------------------------------------------
# Prompt templates — tuning surface for output quality
# ---------------------------------------------------------------------------

OUTLINE_GENERATION_PROMPT = """
Generate a slide-by-slide content outline for the following presentation.

Context:
- Active spec: {spec_name}
- Available layout types: title, content, two_column, section_divider, image_text, data
- Target audience: {audience}
- Purpose: {purpose}

Content gathered:
{content_summary}

Rules:
1. First slide MUST be a TITLE slide with presentation title and subtitle
2. Each section starts with a SECTION_DIVIDER slide (section name only)
3. Content slides use CONTENT layout type by default
4. Use TWO_COLUMN for comparison/contrast content (pros/cons, before/after, etc.)
5. Use IMAGE_TEXT for slides that would benefit from imagery
6. Every slide MUST have a descriptive title AND at least 2 body points with >10 characters each
7. Total slides: follow the gathered content scope (don't pad or truncate significantly)

Output format: JSON array of slide objects:
[
  {{"title": "...", "body": ["...", "..."], "layout_type": "...", "notes": "...", "image_hint": "...", "section_name": "..."}},
  ...
]
"""

SECTION_EXTRACTION_PROMPT = """
Identify the main sections/topics from the following presentation description.

Input:
{user_input}

Return a list of section names (short, 2-5 word labels). Do NOT number them.
Example: ["Executive Summary", "Product Updates", "Financial Results", "Q4 Outlook"]

Sections only (no JSON wrapper):
"""

MINIMAL_OUTLINE_WARNING = (
    "Zero sections identified — generated minimal outline with title slide only."
)

BUDGET_EXHAUSTED_WARNING = (
    "Question budget (8) exhausted with {gap_count} gaps remaining. "
    "Outline generated with best available content."
)

SKIP_QUESTIONS_WARNING = (
    "Mode 'skip_questions' with insufficient input. "
    "Outline generated with best available content from user input alone."
)


# ---------------------------------------------------------------------------
# ContentGatherer — main orchestrator
# ---------------------------------------------------------------------------


class ContentGatherer:
    """Orchestrates content gathering: sufficiency assessment → adaptive
    questioning → outline generation.

    Usage::

        gatherer = ContentGatherer()
        outline = gatherer.gather("Make a presentation about Q3 results")
        path = gatherer.save("outlines")  # → outlines/q3-results.yaml
    """

    def __init__(self):
        self.outline: ContentOutline | None = None
        self.session: QuestionSession | None = None
        self.sufficiency: SufficiencyResult | None = None
        self.gathered_content: dict[str, list[str]] = {}

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def gather(
        self,
        user_input: str,
        mode: str = "auto",
        spec_name: str | None = None,
    ) -> ContentOutline:
        """Run full content gathering pipeline. Returns ContentOutline.

        Parameters
        ----------
        user_input : str
            Raw user-provided content text (may be multi-line).
        mode : str
            ``"auto"`` — full pipeline with adaptive questioning.
            ``"skip_questions"`` — bypass Phase 2; proceed directly from
            sufficiency to outline generation.
        spec_name : str | None
            Optional spec name. If None, auto-resolves via
            ``get_active_spec()`` from spec_commands.

        Returns
        -------
        ContentOutline
            Populated outline with slide entries, layout type recommendations,
            and metadata.

        Raises
        ------
        ValueError
            If ``user_input`` is empty or whitespace-only.
        """
        # --- Guard: empty input ---
        if not user_input.strip():
            raise ValueError(
                "ContentGatherer.gather() requires non-empty user_input. "
                "Provide a presentation topic, outline, or content description."
            )

        # --- Phase 1: Sufficiency Assessment ---
        spec_context = self._resolve_spec_context(spec_name)
        self.sufficiency = assess_sufficiency(user_input, spec_context)

        # --- Phase 2: Adaptive Questioning (if not skipped and not sufficient) ---
        if mode == "skip_questions":
            self.session = QuestionSession()
            self.gathered_content = self._extract_content_from_input(user_input)
            self.session.sections_identified = self._extract_sections(user_input)
        elif self.sufficiency.sufficient:
            # Input is rich enough — skip the interrogation
            self.session = QuestionSession()
            self.gathered_content = self._extract_content_from_input(user_input)
            self.session.sections_identified = self._extract_sections(user_input)
        else:
            self.session = QuestionSession()
            self.gathered_content = self._extract_content_from_input(user_input)
            self._run_adaptive_questioning(user_input)

        # --- Phase 3: Outline Generation ---
        self.outline = self._generate_outline(user_input, spec_name)

        return self.outline

    def save(self, outlines_dir: str | Path = "outlines") -> Path:
        """Save the generated outline as YAML and Markdown. Returns YAML file path.

        Two files are written:
          - <name>.yaml  — structured YAML for programmatic consumption
          - <name>.md    — human-readable WPS Markdown format

        Parameters
        ----------
        outlines_dir : str or Path
            Directory for outline files (created if missing).

        Returns
        -------
        Path
            Path to the written YAML file.

        Raises
        ------
        ValueError
            If ``self.outline`` is None (i.e., ``gather()`` was never called
            or the pipeline did not complete).
        """
        if self.outline is None:
            raise ValueError(
                "No outline to save. Call gather() first."
            )

        outlines_path = Path(outlines_dir)
        outlines_path.mkdir(parents=True, exist_ok=True)

        # Derive name from presentation title (kebab-case)
        raw_title = self.outline.presentation_title or "untitled"
        name = self._title_to_slug(raw_title)

        # Update outline metadata before serialization
        self.outline.metadata.update({
            "name": name,
            "saved_at": datetime.now(timezone.utc).isoformat(),
            "spec_name": self.outline.spec_name or "",
            "question_count": (
                self.session.total_asked if self.session else 0
            ),
        })

        # Serialize YAML (for programmatic use)
        data = _dataclass_to_dict(self.outline)
        yaml_text = yaml.dump(
            data,
            default_flow_style=False,
            sort_keys=False,
            allow_unicode=True,
        )
        yaml_path = outlines_path / f"{name}.yaml"
        yaml_path.write_text(yaml_text, encoding="utf-8")

        # Serialize Markdown (human-readable WPS format)
        md_text = self.outline.to_ppt_markdown()
        md_path = outlines_path / f"{name}.md"
        md_path.write_text(md_text, encoding="utf-8")

        return yaml_path

    @staticmethod
    def load_outline(
        name: str,
        outlines_dir: str | Path = "outlines",
    ) -> ContentOutline:
        """Load a previously saved ContentOutline from YAML.

        Parameters
        ----------
        name : str
            Outline name (without ``.yaml`` extension), e.g. ``"q3-results"``.
        outlines_dir : str or Path
            Directory where outline YAML files are stored.

        Returns
        -------
        ContentOutline
            Reconstructed from the YAML file via ``from_dict()``.

        Raises
        ------
        FileNotFoundError
            If ``outlines/<name>.yaml`` does not exist.
        """
        file_path = Path(outlines_dir) / f"{name}.yaml"

        if not file_path.is_file():
            raise FileNotFoundError(
                f"Outline '{name}' not found at {file_path}"
            )

        with open(file_path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f)

        if data is None:
            data = {}

        return ContentOutline.from_dict(data)

    # ------------------------------------------------------------------
    # Phase 2 — Adaptive Questioning
    # ------------------------------------------------------------------

    def _run_adaptive_questioning(self, user_input: str) -> None:
        """Execute the adaptive questioning flow driven by prompt templates.

        Extracts sections from ``user_input``, generates section-level
        overview questions, then gap-fill questions with remaining budget.

        In production, the LLM runtime processes the prompt templates and
        returns structured answers that populate ``self.gathered_content``.
        In test/programmatic use, the content extracted from the input
        text serves as a baseline.
        """
        session = self.session
        if session is None:
            return

        # --- Extract sections ---
        sections = self._extract_sections(user_input)
        session.sections_identified = sections

        # --- Phase 2a: Section-level overview questions ---
        section_questions = generate_section_questions(
            session,
            sections,
            context="Establishing content for each presentation section.",
        )
        # In the LLM runtime, answers are collected interactively here.
        # For programmatic use, the questions themselves serve as the
        # structure — the content was already extracted from input.

        # --- Phase 2b: Gap-fill questions ---
        gaps = identify_content_gaps(self.gathered_content, sections)

        if gaps and session.can_ask():
            gap_questions = generate_gap_questions(
                session,
                gaps,
                context="Filling content gaps for each section.",
            )
            # Again, LLM runtime collects answers interactively.

        # --- Budget exhausted warning ---
        # Check if gaps remain after budget exhausted
        remaining_gaps = identify_content_gaps(
            self.gathered_content, sections
        )
        if remaining_gaps and not session.can_ask():
            self._emit_warning(
                BUDGET_EXHAUSTED_WARNING.format(
                    gap_count=sum(len(v) for v in remaining_gaps.values())
                )
            )

    def _emit_warning(self, message: str) -> None:
        """Emit a warning to stderr (non-fatal informational message)."""
        print(f"[ContentGatherer] {message}", file=sys.stderr)

    # ------------------------------------------------------------------
    # Phase 3 — Outline Generation
    # ------------------------------------------------------------------

    def _generate_outline(
        self,
        user_input: str,
        spec_name: str | None = None,
    ) -> ContentOutline:
        """Generate a ContentOutline from gathered content and user input.

        Applies layout type rules (first slide → TITLE, section beginnings →
        SECTION_DIVIDER, etc.), validates the result, and retries with
        issue feedback if validation fails.

        Parameters
        ----------
        user_input : str
            Original user input text.
        spec_name : str | None
            Resolved spec name for metadata.

        Returns
        -------
        ContentOutline
            Validated outline with slide entries and layout types.
        """
        resolved_spec = spec_name or self._try_get_active_spec()

        # --- Compile content summary ---
        content_summary = self._build_content_summary(user_input)

        # --- Determine title ---
        title = self._derive_title(user_input)

        # --- Generate slide entries ---
        slides = self._build_slide_entries(user_input, content_summary)

        # --- Construct outline ---
        outline = ContentOutline(
            presentation_title=title,
            spec_name=resolved_spec or "",
            sections=self.session.sections_identified if self.session else [],
            slides=slides,
            metadata={
                "slide_count": len(slides),
                "sufficiency": (
                    "sufficient"
                    if self.sufficiency and self.sufficiency.sufficient
                    else "insufficient"
                ),
                "question_count": (
                    self.session.total_asked if self.session else 0
                ),
            },
        )

        # --- Validate, retry on failure ---
        issues = outline.validate()
        if issues:
            # Attempt regeneration once with issue feedback
            slides_v2 = self._build_slide_entries(
                user_input, content_summary, issues=issues
            )
            outline.slides = slides_v2
            outline.metadata["validation_issues"] = issues
            outline.metadata["slide_count"] = len(slides_v2)

            # Revalidate — if still failing, fix individual slides
            issues_v2 = outline.validate()
            if issues_v2:
                # Fix slides with short body by supplementing from title
                for slide in outline.slides:
                    meaningful = [b for b in slide.body if len(b.strip()) > 10]
                    if not meaningful and slide.title:
                        slide.body = [slide.title]
                outline.metadata["quality_warning"] = (
                    f"Outline has {len(issues_v2)} slides with short content — auto-supplemented"
                )

        # --- Zero sections edge case ---
        if not self.session or not self.session.sections_identified:
            self._emit_warning(MINIMAL_OUTLINE_WARNING)
            outline.metadata["quality_warning"] = MINIMAL_OUTLINE_WARNING

        # --- skip_questions insufficient warning ---
        if (
            self.sufficiency
            and not self.sufficiency.sufficient
            and self.session
            and self.session.total_asked == 0
        ):
            outline.metadata["quality_warning"] = SKIP_QUESTIONS_WARNING

        return outline

    def _build_slide_entries(
        self,
        user_input: str,
        content_summary: str,
        issues: list[str] | None = None,
    ) -> list[SlideEntry]:
        """Build SlideEntry instances from content summary using prompt template.

        This method constructs the OUTLINE_GENERATION_PROMPT with gathered
        content. In production, the LLM runtime evaluates the prompt and
        returns a JSON array of slide objects which are parsed into
        SlideEntry instances. In programmatic mode, it constructs slides
        directly from the section and content structure.
        """
        sections = self.session.sections_identified if self.session else []

        # If no sections identified, create a minimal outline
        if not sections:
            return self._build_minimal_outline(user_input)

        slides: list[SlideEntry] = []
        slide_num = 0
        prev_layout: OutlineLayoutType | None = None

        # First slide: TITLE
        slide_num += 1
        slides.append(SlideEntry(
            slide_number=slide_num,
            title=self._derive_title(user_input),
            body=[f"{self._derive_subtitle(user_input)}"],
            layout_type=OutlineLayoutType.TITLE,
        ))

        for section in sections:
            # Section divider
            slide_num += 1
            slides.append(SlideEntry(
                slide_number=slide_num,
                title=section,
                body=[],  # Section dividers have no body
                layout_type=OutlineLayoutType.SECTION_DIVIDER,
                section_name=section,
            ))

            # Content slide for this section
            content_points = self.gathered_content.get(section, [])
            if not content_points:
                # Derive from user input
                content_points = self._section_points_from_input(
                    user_input, section
                )

            if content_points:
                # Determine layout type based on content
                layout = self._determine_layout_type(
                    content_points, prev_layout
                )

                # Split into multiple slides if needed (~3-5 bullets per slide)
                chunks = self._chunk_points(content_points, max_per_slide=5)
                for chunk in chunks:
                    slide_num += 1
                    slides.append(SlideEntry(
                        slide_number=slide_num,
                        title=section,
                        body=chunk,
                        layout_type=layout,
                        section_name=section,
                        image_hint=self._derive_image_hint(
                            section, chunk
                        ),
                    ))
                    prev_layout = layout
            else:
                # Minimal content slide
                slide_num += 1
                slides.append(SlideEntry(
                    slide_number=slide_num,
                    title=section,
                    body=[f"Content for {section}"],
                    layout_type=OutlineLayoutType.CONTENT,
                    section_name=section,
                ))

        return slides

    def _build_minimal_outline(self, user_input: str) -> list[SlideEntry]:
        """Build a minimal outline when zero sections are identified."""
        title = self._derive_title(user_input)
        return [
            SlideEntry(
                slide_number=1,
                title=title,
                body=[user_input[:200] if len(user_input) > 200 else user_input],
                layout_type=OutlineLayoutType.TITLE,
            ),
        ]

    def _determine_layout_type(
        self,
        points: list[str],
        prev_layout: OutlineLayoutType | None = None,
    ) -> OutlineLayoutType:
        """Determine the best layout type for a set of bullet points.

        Heuristics:
        - Points comparing/contrasting → TWO_COLUMN
        - Points mentioning images or visuals → IMAGE_TEXT
        - Points mentioning data, stats, or charts → DATA
        - Default → CONTENT
        """
        combined = " ".join(points).lower()

        comparison_keywords = [
            "vs", "versus", "comparison", "compare", "before",
            "after", "pros", "cons", "versus", "difference",
        ]
        image_keywords = [
            "image", "photo", "screenshot", "diagram", "picture",
            "visual", "graphic", "illustration",
        ]
        data_keywords = [
            "data", "chart", "graph", "statistic", "metric",
            "percentage", "number", "table", "figure", "%",
        ]

        for kw in comparison_keywords:
            if kw in combined:
                return OutlineLayoutType.TWO_COLUMN

        for kw in image_keywords:
            if kw in combined:
                return OutlineLayoutType.IMAGE_TEXT

        for kw in data_keywords:
            if kw in combined:
                return OutlineLayoutType.DATA

        return OutlineLayoutType.CONTENT

    def _derive_image_hint(
        self,
        section: str,
        points: list[str],
    ) -> str:
        """Derive an image hint keyword from section name and content."""
        # Use keywords from the section and points
        if points and any(
            kw in " ".join(points).lower()
            for kw in ["chart", "graph", "data"]
        ):
            img_word = next(
                kw for kw in points
                if any(
                    kw2 in kw.lower()
                    for kw2 in ["chart", "graph", "data"]
                )
            )
            return img_word[:40]

        return ""

    def _chunk_points(
        self,
        points: list[str],
        max_per_slide: int = 5,
    ) -> list[list[str]]:
        """Split a list of bullet points into chunks of max_per_slide.

        Preserves table blocks (lines starting with |) — a table started
        on one slide stays entirely on that slide.
        """
        chunks: list[list[str]] = []
        current: list[str] = []
        in_table = False

        for point in points:
            is_table_line = point.strip().startswith("|") and point.strip().endswith("|")
            is_table_header = is_table_line and not in_table

            # If this is a table header and current chunk has items, flush it
            if is_table_header and current and len(current) >= max_per_slide * 0.5:
                chunks.append(current)
                current = []
                in_table = False

            # Table mode: keep accumulating until table ends
            if is_table_line:
                in_table = True
                current.append(point)
                continue
            elif in_table:
                # Table ended, flush current chunk
                in_table = False
                if current:
                    chunks.append(current)
                    current = []
                current.append(point)
                continue

            # Normal point
            current.append(point)
            if len(current) >= max_per_slide:
                chunks.append(current)
                current = []

        if current:
            chunks.append(current)

        return chunks if chunks else [[]]

    def _derive_title(self, user_input: str) -> str:
        """Derive a presentation title from user input text."""
        lines = [l.strip() for l in user_input.strip().splitlines() if l.strip()]
        if lines:
            first = lines[0]
            first = first.lstrip("#").strip()  # Remove markdown heading markers
            return first[:100] if len(first) > 100 else first
        return "Presentation"

    def _derive_subtitle(self, user_input: str) -> str:
        """Derive a subtitle from the second line, or provide a default."""
        lines = [l.strip() for l in user_input.strip().splitlines() if l.strip()]
        speech_markers = ["讲师口述", "讲师身份", "讲师现场", "讲师发问", "演示建议", "> **讲师"]
        import re
        for line in lines[1:]:
            clean = line.lstrip("#>").strip()
            clean = re.sub(r"[（(]?\d+分钟[）)]?", "", clean).strip()
            if clean and len(clean) > 5 and not any(m in clean for m in speech_markers):
                return clean[:100]
        return ""

    # ------------------------------------------------------------------
    # Content extraction helpers
    # ------------------------------------------------------------------

    def _extract_content_from_input(
        self,
        user_input: str,
    ) -> dict[str, list[str]]:
        """Extract PPT bullet-style content from user input text.

        Parses lines that look like section headers and gathers following
        lines as content points. Filters out lecture/speech text,
        blockquotes, time markers, and presenter instructions.

        Returns
        -------
        dict[str, list[str]]
            Mapping of section name → list of content point strings.
        """
        sections = self._extract_sections(user_input)
        result: dict[str, list[str]] = {}

        lines = user_input.strip().splitlines()
        current_section: str | None = None
        speech_markers = ["讲师口述", "讲师身份", "讲师现场", "讲师发问", "演示建议", "> **讲师"]

        for line in lines:
            stripped = line.strip()
            if not stripped:
                continue

            # Skip speech/lecture markers entirely
            if any(m in stripped for m in speech_markers):
                continue
            # Skip blockquotes (> ...) — these are lecture scripts, not slide content
            if stripped.startswith(">"):
                continue
            # Skip time-only markers like "(5分钟)" or "（2分钟）"
            import re
            if re.match(r"^[（(]?\d+分钟[）)]?", stripped):
                continue
            # Skip horizontal rules
            if re.match(r"^[-—]{3,}$", stripped):
                continue
            # Skip bare markdown headers
            if re.match(r"^#+\s+", stripped):
                continue

            # Check if this line is a section header
            is_header = False
            for section in sections:
                if section and section.lower() in stripped.lower():
                    current_section = section
                    is_header = True
                    if section not in result:
                        result[section] = []
                    break

            if is_header:
                continue

            # Not a header — append to current section
            if current_section:
                clean = self._clean_bullet(stripped)
                # Filter: skip speech, blockquotes, short fragments
                if clean and len(clean) > 5 and not any(m in clean for m in speech_markers):
                    result.setdefault(current_section, []).append(clean)
            else:
                if sections and sections[0] not in result:
                    result.setdefault(sections[0], [])
                clean = self._clean_bullet(stripped)
                if clean and len(clean) > 5 and not any(m in clean for m in speech_markers):
                    if sections:
                        result.setdefault(sections[0], []).append(clean)

        # Ensure every section has at least an empty list
        for section in sections:
            result.setdefault(section, [])

        return result

    def _extract_sections(self, user_input: str) -> list[str]:
        """Extract section/topic names from user input.

        Heuristic: looks for ##/### headers, numbered items (1., 2.),
        and short section titles.

        Skips: lecturer instructions, bold-formatted arguments, quotes.
        """
        lines = [l.strip() for l in user_input.strip().splitlines() if l.strip()]
        sections: list[str] = []

        # Lines that are NOT valid section headers
        speech_markers = ["讲师口述", "讲师身份", "讲师现场", "演示建议", "**讲师"]
        skip_patterns = [
            r"^\*.*\*$",           # *italic emphasis lines*
            r"^>.+",                # blockquotes
            r"^[-—]+$",             # horizontal rules
            r"^\d+分钟[）)]?$",     # "2分钟）"
            r"^[（(]?\d+分钟[）)]?",  # time markers like "(5分钟)"
            r"^#+\s*$",             # bare headers
        ]

        # Priority: ## and ### headers are the most reliable section markers
        import re
        for line in lines:
            m = re.match(r"^##+\s+(.+)$", line)
            if m:
                name = m.group(1).strip().rstrip(":")
                # Strip time markers like "（2分钟）"
                name = re.sub(r"[（(]?\d+分钟[）)]?", "", name).strip()
                if name and len(name) > 1 and not any(m in name for m in speech_markers):
                    sections.append(name)

        # Fallback: numbered structure → "1. 开场与前情回顾"
        if not sections:
            for line in lines:
                # Skip speech/lecture lines
                if any(m in line for m in speech_markers):
                    continue
                # Skip patterns
                skip = False
                for sp in skip_patterns:
                    if re.match(sp, line):
                        skip = True
                        break
                if skip:
                    continue

                m = re.match(
                    r'^(?:\d+[\.\)]\s*|Section\s*\d+[:\-]\s*)(.+)$',
                    line,
                )
                if m:
                    name = m.group(1).strip().rstrip(":")
                    if name and len(name) > 2:
                        sections.append(name)

        # Pattern 2: Short lines that look like headers (≤40 chars, no speech markers)
        if not sections:
            for line in lines:
                if any(m in line for m in speech_markers):
                    continue
                if (
                    3 <= len(line) <= 40
                    and not line.startswith(">")
                    and not line.startswith("*")
                    and not line.endswith("。")
                    and not line.endswith("：")
                ):
                    sections.append(line)

        # Pattern 3: Fall back to groups of content as "Section N"
        if not sections:
            sections = [f"Section {i + 1}" for i in range(min(3, max(2, len(lines) // 3)))]

        # Deduplicate while preserving order
        seen = set()
        unique = []
        for s in sections:
            if s.lower() not in seen:
                seen.add(s.lower())
                unique.append(s)
        return unique

    def _section_points_from_input(
        self,
        user_input: str,
        section_name: str,
    ) -> list[str]:
        """Extract points relevant to a specific section from user input."""
        points: list[str] = []
        in_section = False
        speech_markers = ["讲师口述", "讲师身份", "讲师现场", "讲师发问", "演示建议", "> **讲师"]
        for line in user_input.strip().splitlines():
            stripped = line.strip()
            if not stripped:
                in_section = False
                continue
            if any(m in stripped for m in speech_markers):
                in_section = False
                continue
            if stripped.startswith(">"):
                in_section = False
                continue

            if section_name.lower() in stripped.lower() and len(stripped) < 80:
                in_section = True
                continue

            if in_section and (
                stripped.startswith(("-", "*", "•", "·"))
                or (stripped.startswith("|") and stripped.endswith("|"))  # table
            ):
                clean = self._clean_bullet(stripped)
                if clean and len(clean) > 3 and not any(m in clean for m in speech_markers):
                    points.append(clean)

        return points if points else []

    def _build_content_summary(self, user_input: str) -> str:
        """Build a formatted content summary string for the LLM prompt."""
        parts: list[str] = []

        parts.append(f"User input summary:\n{user_input[:1000]}")

        if self.gathered_content:
            parts.append("\nGathered content by section:")
            for section, points in self.gathered_content.items():
                if points:
                    parts.append(
                        f"  {section}:\n    - "
                        + "\n    - ".join(points)
                    )

        return "\n".join(parts)

    # ------------------------------------------------------------------
    # Utility helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _clean_bullet(line: str) -> str:
        """Remove bullet markers and leading/trailing whitespace."""
        cleaned = line.strip()
        for marker in ("-", "*", "•", "·", "→", "➤", "○"):
            if cleaned.startswith(marker):
                cleaned = cleaned[len(marker):].strip()
                break
        return cleaned

    @staticmethod
    def _title_to_slug(title: str) -> str:
        """Convert a presentation title to a kebab-case filename slug."""
        slug = title.lower()
        # Replace runs of non-alphanumeric with single dash
        import re
        slug = re.sub(r'[^a-z0-9]+', '-', slug)
        slug = slug.strip('-')
        return slug or "untitled"

    @staticmethod
    def _try_get_active_spec() -> str | None:
        """Try to get the active spec name; return None if unavailable."""
        try:
            from ppt_skill.cli.spec_commands import get_active_spec
            return get_active_spec()
        except Exception:
            return None

    def _resolve_spec_context(
        self,
        spec_name: str | None = None,
    ) -> dict | None:
        """Resolve spec context for sufficiency assessment.

        If ``spec_name`` is provided, loads that spec. Otherwise, tries
        ``get_active_spec()``. Returns a dict with spec metadata or None.
        """
        name = spec_name or self._try_get_active_spec()
        if not name:
            return None

        try:
            spec_path = Path("specs") / f"{name}.yaml"
            if not spec_path.is_file():
                return None

            with open(spec_path, "r", encoding="utf-8") as f:
                data = yaml.safe_load(f) or {}

            return {
                "name": name,
                "slide_count": len(data.get("slides", [])),
                "spec_path": str(spec_path),
            }
        except Exception:
            return None


# ---------------------------------------------------------------------------
# Module exports
# ---------------------------------------------------------------------------

__all__ = [
    "ContentGatherer",
    "OUTLINE_GENERATION_PROMPT",
    "SECTION_EXTRACTION_PROMPT",
]
