"""Content gathering data model — dataclass schemas for content outline and questioning.

All types are plain Python dataclasses (NOT Pydantic) to keep dependencies
minimal. Every field has default values so instances can be constructed
partially and populated incrementally.

These schemas define the contract between Phase 3 (Content Gathering) and
Phase 4 (PPT Generation) — ContentOutline is serialized to YAML and consumed
by the generation pipeline. Internal types (SufficiencyResult, Question,
QuestionSession) support the adaptive questioning workflow.

Design follows the same conventions as ppt_skill.spec.spec_model.py:
  - from __future__ import annotations
  - dataclasses with field(default=...) for mutable defaults
  - str-based Enums with .value string keys
  - @classmethod constructors for deserialization
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum

import yaml

from ppt_skill.spec.extractor import _dataclass_to_dict


# ---------------------------------------------------------------------------
# Phase 4 contract types
# ---------------------------------------------------------------------------


class OutlineLayoutType(str, Enum):
    """Layout types the content outline can recommend per slide.

    Constrained to values Phase 4 can render:
      TITLE, CONTENT, TWO_COLUMN, SECTION_DIVIDER, IMAGE_TEXT, DATA.
    Extra types beyond existing SlideType: TWO_COLUMN.
    """

    TITLE = "title"
    CONTENT = "content"
    TWO_COLUMN = "two_column"
    SECTION_DIVIDER = "section_divider"
    IMAGE_TEXT = "image_text"
    DATA = "data"


@dataclass
class SlideEntry:
    """One slide in the content outline — consumed by Phase 4.

    Each SlideEntry captures the complete content for a single slide:
    title, body bullets, recommended layout, speaker notes, image
    hints, and section membership.

    Every field has a default — construction possible with partial data.
    The body field is a list of strings (bullet points).
    image_hint is an optional keyword for Phase 4 to generate/find imagery.
    section_name ties the slide to a parent section.
    """

    slide_number: int = 0
    title: str = ""
    body: list[str] = field(default_factory=list)
    layout_type: OutlineLayoutType = OutlineLayoutType.CONTENT
    notes: str = ""
    image_hint: str = ""
    section_name: str = ""


@dataclass
class ContentOutline:
    """Full slide-by-slide content outline — output of Phase 3, input to Phase 4.

    This is the primary serialization target. ContentOutline captures the
    complete structured content for a presentation: metadata, title,
    audience info, section structure, and per-slide entries with
    layout recommendations.

    Usage::

        outline = ContentOutline(presentation_title="Q4 Review")
        slide = SlideEntry(slide_number=1, title="Introduction", body=["..."]))
        outline.slides.append(slide)
        issues = outline.validate()
        if not issues:
            yaml_str = outline.to_dict()
    """

    metadata: dict = field(default_factory=dict)
    presentation_title: str = ""
    presentation_subtitle: str = ""
    target_audience: str = ""
    presentation_purpose: str = ""
    sections: list[str] = field(default_factory=list)
    slides: list[SlideEntry] = field(default_factory=list)
    spec_name: str = ""

    # ------------------------------------------------------------------
    # Validation
    # ------------------------------------------------------------------

    def validate(self) -> list[str]:
        """Validate the outline and return a list of issue descriptions.

        Returns an empty list if the outline is valid.
        Checks: title present, slides non-empty, per-slide title/body/
        layout_type/slide_number correctness.
        """
        issues: list[str] = []

        # --- Global checks ---
        if not self.presentation_title.strip():
            issues.append("Missing presentation title")

        if len(self.slides) == 0:
            issues.append("No slides in outline")
            # Cannot validate per-slide if no slides
            return issues

        # --- Per-slide checks ---
        for i, slide in enumerate(self.slides):
            n = i + 1  # human-readable slide number

            # Title
            if not slide.title.strip():
                issues.append(f"Slide {n}: empty title")

            # Body: at least one entry with > 10 characters after strip
            meaningful_body = [
                b for b in slide.body if len(b.strip()) > 10
            ]
            if not meaningful_body:
                issues.append(f"Slide {n}: body too short or empty")

            # Layout type
            if slide.layout_type not in OutlineLayoutType:
                issues.append(f"Slide {n}: invalid layout_type '{slide.layout_type}'")

            # Slide number sequence
            if slide.slide_number != n:
                issues.append(
                    f"Slide {slide.slide_number}: slide_number out of sequence "
                    f"(expected {n})"
                )

        return issues

    # ------------------------------------------------------------------
    # Serialization
    # ------------------------------------------------------------------

    def to_dict(self) -> dict:
        """Serialize to a plain dict, ready for YAML output.

        Validates first — raises ValueError if the outline is invalid.
        Uses _dataclass_to_dict from ppt_skill.spec.extractor for
        recursive dataclass/Enum → dict/string conversion.
        """
        issues = self.validate()
        if issues:
            raise ValueError(
                f"ContentOutline validation failed with "
                f"{len(issues)} issue(s): {issues}"
            )

        return _dataclass_to_dict(self)

    def to_yaml(self) -> str:
        """Serialize to a YAML string.

        Calls to_dict() internally (validates first), then dumps to YAML
        with default_flow_style=False, sort_keys=False, allow_unicode=True.
        """
        data = self.to_dict()
        return yaml.dump(
            data,
            default_flow_style=False,
            sort_keys=False,
            allow_unicode=True,
        )

    @classmethod
    def from_dict(cls, data: dict) -> ContentOutline:
        """Reconstruct a ContentOutline from a YAML-parsed dict.

        Handles nested SlideEntry reconstruction and converts layout_type
        strings back to OutlineLayoutType enum values. Filters fields to
        those present in __dataclass_fields__ to be defensive against
        extra keys.
        """
        # Filter to known fields
        valid_fields = {
            k: v for k, v in data.items()
            if k in cls.__dataclass_fields__
        }

        # Reconstruct slides from nested dicts
        if "slides" in valid_fields:
            slide_entries: list[SlideEntry] = []
            for slide_data in valid_fields["slides"]:
                if isinstance(slide_data, dict):
                    # Convert layout_type string → enum
                    lt_raw = slide_data.get("layout_type", "content")
                    lt_enum = OutlineLayoutType(lt_raw) if lt_raw else OutlineLayoutType.CONTENT

                    entry = SlideEntry(
                        slide_number=slide_data.get("slide_number", 0),
                        title=slide_data.get("title", ""),
                        body=slide_data.get("body", []),
                        layout_type=lt_enum,
                        notes=slide_data.get("notes", ""),
                        image_hint=slide_data.get("image_hint", ""),
                        section_name=slide_data.get("section_name", ""),
                    )
                    slide_entries.append(entry)
                elif isinstance(slide_data, SlideEntry):
                    slide_entries.append(slide_data)
            valid_fields["slides"] = slide_entries

        return cls(**valid_fields)


# ---------------------------------------------------------------------------
# Sufficiency assessment types
# ---------------------------------------------------------------------------


@dataclass
class SufficiencyResult:
    """Result of input sufficiency assessment.

    sufficient: True if input has enough detail to skip questioning.
    confidence: Float 0.0–1.0 representing assessment confidence.
    missing_dimensions: Which aspects are lacking (e.g., "structure",
        "detail", "audience").
    section_count: How many sections/topics were identified.
    estimated_slide_count: Rough estimate of resulting slides.
    scores: Per-dimension scores dict (e.g., {"structure": 2, "detail": 1}).
    total_score: Sum of all dimension scores.
    rationale: Human-readable explanation of the assessment.

    The sufficiency threshold is documented here but enforced in the
    sufficiency assessment module (sufficiency.py), not the dataclass:
        sufficient when total_score >= 5 AND structure >= 1 AND detail >= 1
    """

    sufficient: bool = False
    confidence: float = 0.0
    missing_dimensions: list[str] = field(default_factory=list)
    section_count: int = 0
    estimated_slide_count: int = 0
    scores: dict[str, int] = field(default_factory=dict)
    total_score: int = 0
    rationale: str = ""


# ---------------------------------------------------------------------------
# Adaptive questioning types
# ---------------------------------------------------------------------------


@dataclass
class Question:
    """A single question in the adaptive questioning session.

    category: One of "structure", "detail", "audience", or "storytelling".
        The priority ordering follows: structure > detail > audience > storytelling.
    target_section: None for section-level overview questions; section name
        string for gap-fill questions targeting a specific section.
    context: Why this question is being asked (references what's already known).
    """

    id: int = 0
    category: str = ""
    text: str = ""
    target_section: str | None = None
    context: str = ""


@dataclass
class QuestionSession:
    """Tracks state across an adaptive questioning session.

    Budget starts at 8 questions. Section-level questions are asked first
    (one per section), then remaining budget is allocated to gap-fill
    questions for sections with identified knowledge gaps.

    gaps_per_section keys are section names, values are lists of gap types
    (e.g., "missing_detail", "no_examples", "no_takeaway", "missing_transition").
    """

    questions_asked: list[Question] = field(default_factory=list)
    budget_remaining: int = 8
    sections_identified: list[str] = field(default_factory=list)
    gaps_per_section: dict[str, list[str]] = field(default_factory=dict)

    @property
    def total_asked(self) -> int:
        """Total number of questions asked in this session."""
        return len(self.questions_asked)

    def can_ask(self) -> bool:
        """True if budget remains for more questions."""
        return self.budget_remaining > 0

    def mark_asked(self, question: Question) -> None:
        """Record a question as asked, decrement budget."""
        self.questions_asked.append(question)
        self.budget_remaining -= 1


__all__ = [
    "ContentOutline",
    "OutlineLayoutType",
    "Question",
    "QuestionSession",
    "SlideEntry",
    "SufficiencyResult",
]
