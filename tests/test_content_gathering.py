"""Integration tests for the full content gathering pipeline.

Tests the ContentGatherer orchestrator end-to-end: model serialization,
validation, sufficiency assessment, question budget tracking, pipeline
execution, and YAML persistence. No LLM calls — uses mode="skip_questions"
with sufficiently detailed input for pipeline tests.
"""

from __future__ import annotations

from pathlib import Path

import pytest
import yaml

from ppt_skill.content.model import (
    ContentOutline,
    OutlineLayoutType,
    Question,
    QuestionSession,
    SlideEntry,
    SufficiencyResult,
)
from ppt_skill.content.sufficiency import (
    assess_sufficiency,
    build_sufficiency_result,
)
from ppt_skill.content.questioning import (
    generate_gap_questions,
    generate_section_questions,
    identify_content_gaps,
)


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------


@pytest.fixture
def detailed_input() -> str:
    """Sufficiently detailed multi-section presentation input."""
    return """
Q3 2025 Business Review

1. Executive Summary
   - Revenue: $4.2M (up 15% YoY)
   - Customer growth: 1,200 new accounts
   - Key wins: Enterprise deal with Acme Corp ($800K)

2. Product Updates
   - Launched v2.3 with AI recommendations
   - User engagement up 22%
   - Mobile app released (iOS + Android)

3. Q4 Outlook
   - Target: $5M revenue
   - Hiring: 5 new engineers
   - Market expansion: APAC region launch

Target audience: Board of Directors
Purpose: Quarterly business review and Q4 planning
"""


@pytest.fixture
def vague_input() -> str:
    """Insufficiently detailed input — short, no structure."""
    return "Make a presentation"


# ---------------------------------------------------------------------------
# Test 1: Model serialization
# ---------------------------------------------------------------------------


class TestModelSerialization:
    """ContentOutline round-trip through dict/YAML serialization."""

    def test_outline_to_dict_and_back(self):
        """ContentOutline → to_dict() → from_dict() preserves all fields."""
        from ppt_skill.content.model import ContentOutline, SlideEntry, OutlineLayoutType

        outline = ContentOutline(
            presentation_title="Test Deck",
            presentation_subtitle="A Test",
            target_audience="Developers",
            presentation_purpose="Testing",
            sections=["Intro", "Body", "Conclusion"],
            spec_name="test-spec",
        )

        # Slide 1: TITLE
        outline.slides.append(SlideEntry(
            slide_number=1,
            title="Welcome",
            body=["Subtitle goes here"],
            layout_type=OutlineLayoutType.TITLE,
        ))

        # Slide 2: SECTION_DIVIDER
        outline.slides.append(SlideEntry(
            slide_number=2,
            title="Body",
            body=["Start of body section"],
            layout_type=OutlineLayoutType.SECTION_DIVIDER,
            section_name="Body",
        ))

        # Slide 3: CONTENT
        outline.slides.append(SlideEntry(
            slide_number=3,
            title="Key Data",
            body=["Revenue: $4.2M up 15% YoY", "Customer growth: 1,200 new accounts"],
            layout_type=OutlineLayoutType.CONTENT,
            section_name="Body",
        ))

        # to_dict → verify all fields
        data = outline.to_dict()
        assert data["presentation_title"] == "Test Deck"
        assert data["presentation_subtitle"] == "A Test"
        assert data["target_audience"] == "Developers"
        assert len(data["slides"]) == 3

        # Enum values are strings
        assert data["slides"][0]["layout_type"] == "title"
        assert data["slides"][1]["layout_type"] == "section_divider"
        assert data["slides"][2]["layout_type"] == "content"

        # from_dict reconstruct
        reconstructed = ContentOutline.from_dict(data)
        assert reconstructed.presentation_title == "Test Deck"
        assert len(reconstructed.slides) == 3
        assert reconstructed.slides[0].layout_type == OutlineLayoutType.TITLE
        assert reconstructed.slides[2].layout_type == OutlineLayoutType.CONTENT
        assert reconstructed.slides[2].body[0] == "Revenue: $4.2M up 15% YoY"

    def test_outline_yaml_round_trip(self, tmp_path: Path):
        """YAML dump → load → from_dict preserves all fields."""
        from ppt_skill.content.model import ContentOutline, SlideEntry, OutlineLayoutType

        outline = ContentOutline(
            presentation_title="YAML Test",
            metadata={"name": "yaml-test"},
        )
        outline.slides = [
            SlideEntry(
                slide_number=1,
                title="Slide One",
                body=["First point with enough characters for validation", "Second point also has sufficient length"],
                layout_type=OutlineLayoutType.CONTENT,
                image_hint="chart",
                notes="Speaker note here",
                section_name="Main",
            ),
        ]

        # Convert to dict, dump to YAML
        data = outline.to_dict()
        yaml_text = yaml.dump(data, default_flow_style=False, sort_keys=False)

        # Parse YAML and reconstruct
        data_reloaded = yaml.safe_load(yaml_text)
        reconstructed = ContentOutline.from_dict(data_reloaded)

        assert reconstructed.presentation_title == "YAML Test"
        assert len(reconstructed.slides) == 1
        assert reconstructed.slides[0].title == "Slide One"
        assert reconstructed.slides[0].body[0].startswith("First point")
        assert reconstructed.slides[0].layout_type == OutlineLayoutType.CONTENT
        assert reconstructed.slides[0].image_hint == "chart"
        assert reconstructed.slides[0].notes == "Speaker note here"


# ---------------------------------------------------------------------------
# Test 2: Model validation
# ---------------------------------------------------------------------------


class TestModelValidation:
    """ContentOutline.validate() catches structural issues."""

    def test_valid_outline_passes(self):
        """A well-formed outline should validate with zero issues."""
        outline = ContentOutline(
            presentation_title="Valid Deck",
            slides=[
                SlideEntry(
                    slide_number=1,
                    title="Title Slide",
                    body=["This is a proper subtitle with enough characters"],
                    layout_type=OutlineLayoutType.TITLE,
                ),
                SlideEntry(
                    slide_number=2,
                    title="Content Slide",
                    body=["Bullet one with some content", "Bullet two with more content"],
                    layout_type=OutlineLayoutType.CONTENT,
                ),
            ],
        )
        issues = outline.validate()
        assert issues == [], f"Expected no issues, got: {issues}"

    def test_missing_title(self):
        """Missing presentation_title should be caught."""
        outline = ContentOutline(
            presentation_title="",
            slides=[
                SlideEntry(
                    slide_number=1,
                    title="Only Slide",
                    body=["Some body content here"],
                ),
            ],
        )
        issues = outline.validate()
        assert any("Missing presentation title" in issue for issue in issues), \
            f"Expected 'Missing presentation title' in issues, got: {issues}"

    def test_empty_slides_list(self):
        """Zero slides should produce validation error."""
        outline = ContentOutline(
            presentation_title="No Slides",
            slides=[],
        )
        issues = outline.validate()
        assert any("No slides in outline" in issue for issue in issues), \
            f"Expected 'No slides in outline' in issues, got: {issues}"

    def test_empty_slide_title(self):
        """A slide with an empty title should be caught."""
        outline = ContentOutline(
            presentation_title="Test",
            slides=[
                SlideEntry(
                    slide_number=1,
                    title="",
                    body=["Some content here for the slide body"],
                ),
            ],
        )
        issues = outline.validate()
        assert any("empty title" in issue for issue in issues), \
            f"Expected empty title issue, got: {issues}"

    def test_short_body(self):
        """A slide with body entries < 10 chars should be flagged."""
        outline = ContentOutline(
            presentation_title="Test",
            slides=[
                SlideEntry(
                    slide_number=1,
                    title="Test Slide",
                    body=["Hi"],  # too short
                ),
            ],
        )
        issues = outline.validate()
        assert any("body too short" in issue for issue in issues), \
            f"Expected body-too-short issue, got: {issues}"

    def test_invalid_layout_type(self):
        """A slide with an invalid layout_type value should be caught."""
        outline = ContentOutline(
            presentation_title="Test",
            slides=[
                SlideEntry(
                    slide_number=1,
                    title="Test Slide",
                    body=["Content here that is long enough"],
                    layout_type="nonexistent_layout",  # invalid
                ),
            ],
        )
        issues = outline.validate()
        assert any("invalid layout_type" in issue for issue in issues), \
            f"Expected invalid layout_type issue, got: {issues}"

    def test_slide_number_sequence(self):
        """Slides with out-of-sequence numbers should be flagged."""
        outline = ContentOutline(
            presentation_title="Test",
            slides=[
                SlideEntry(
                    slide_number=5,  # first slide should be 1
                    title="Slide Five",
                    body=["Content here that is long enough for the slide body to pass"],
                ),
            ],
        )
        issues = outline.validate()
        assert any("out of sequence" in issue for issue in issues), \
            f"Expected slide_number out-of-sequence issue, got: {issues}"


# ---------------------------------------------------------------------------
# Test 3: Sufficiency assessment — insufficient input
# ---------------------------------------------------------------------------


class TestSufficiencyInsufficient:
    """Vague input correctly classified as insufficient."""

    def test_vague_input_insufficient(self, vague_input: str):
        """Single short sentence → insufficient, low confidence."""
        result = assess_sufficiency(vague_input)

        assert result.sufficient is False, \
            f"Expected insufficient for: {vague_input!r}"
        assert result.confidence < 0.5, \
            f"Confidence should be < 0.5 for vague input, got {result.confidence}"

    def test_vague_input_missing_dimensions(self, vague_input: str):
        """Vague input should be missing multiple dimensions."""
        result = assess_sufficiency(vague_input)

        # Without LLM, default result has all 4 dimensions missing
        assert len(result.missing_dimensions) >= 3, \
            f"Expected at least 3 missing dimensions, got {result.missing_dimensions}"

    def test_build_sufficiency_result_insufficient(self):
        """build_sufficiency_result with low scores → insufficient."""
        scores = {"structure": 0, "detail": 0, "audience": 1, "scope": 1}
        result = build_sufficiency_result(
            scores=scores,
            rationale="Test: very vague input",
            section_count=0,
            estimated_slide_count=1,
        )

        # total = 2, structure=0, detail=0 → NOT sufficient
        assert result.sufficient is False
        assert result.total_score == 2
        assert result.confidence == pytest.approx(2.0 / 8.0)

    def test_assess_sufficiency_returns_prompt_in_rationale(self, vague_input: str):
        """When no LLM scores, the rationale contains the rubric prompt."""
        result = assess_sufficiency(vague_input)
        assert "Evaluate the following content" in result.rationale
        assert vague_input in result.rationale


# ---------------------------------------------------------------------------
# Test 4: Sufficiency assessment — sufficient input
# ---------------------------------------------------------------------------


class TestSufficiencySufficient:
    """Detailed input correctly identified as sufficient."""

    def test_build_sufficiency_result_sufficient(self):
        """build_sufficiency_result with high scores → sufficient."""
        scores = {"structure": 2, "detail": 2, "audience": 2, "scope": 2}
        result = build_sufficiency_result(
            scores=scores,
            rationale="Rich, structured content with audience/purpose stated.",
            section_count=3,
            estimated_slide_count=8,
        )

        # total = 8, structure=2, detail=2 → SUFFICIENT
        assert result.sufficient is True
        assert result.confidence >= 0.7
        assert result.section_count >= 3
        assert result.total_score == 8

    def test_build_sufficiency_result_borderline(self):
        """Scores exactly at threshold → sufficient."""
        scores = {"structure": 1, "detail": 1, "audience": 2, "scope": 1}
        result = build_sufficiency_result(
            scores=scores,
            rationale="Borderline sufficient.",
            section_count=2,
            estimated_slide_count=4,
        )

        # total = 5, structure=1, detail=1 → SUFFICIENT (right at threshold)
        assert result.sufficient is True
        assert result.total_score == 5

    def test_build_sufficiency_result_just_below(self):
        """Scores just below threshold → not sufficient."""
        scores = {"structure": 1, "detail": 1, "audience": 1, "scope": 1}
        result = build_sufficiency_result(
            scores=scores,
            rationale="Just below threshold.",
            section_count=1,
            estimated_slide_count=2,
        )

        # total = 4 → below 5 threshold
        assert result.sufficient is False
        assert result.total_score == 4

    def test_build_sufficiency_missing_dimensions(self):
        """Dimensions with score < 2 appear in missing_dimensions."""
        scores = {"structure": 2, "detail": 1, "audience": 2, "scope": 0}
        result = build_sufficiency_result(
            scores=scores,
            rationale="Partial coverage.",
            section_count=2,
            estimated_slide_count=5,
        )

        # detail=1, scope=0 → both should appear in missing_dimensions
        assert "detail" in result.missing_dimensions
        assert "scope" in result.missing_dimensions
        assert "structure" not in result.missing_dimensions  # score=2


# ---------------------------------------------------------------------------
# Test 5: Question session budget
# ---------------------------------------------------------------------------


class TestQuestionSessionBudget:
    """QuestionSession budget tracking and enforcement."""

    def test_section_questions_within_budget(self):
        """Generating questions for ≤8 sections should consume budget correctly."""
        session = QuestionSession()
        sections = ["A", "B", "C", "D", "E", "F"]

        questions = generate_section_questions(session, sections)

        # Should generate one per section (6 questions, budget remaining = 2)
        assert len(questions) == 6, f"Expected 6 questions, got {len(questions)}"
        assert session.budget_remaining == 2
        assert session.total_asked == 6

    def test_section_questions_capped_at_budget(self):
        """Generating questions for >8 sections should cap at budget."""
        session = QuestionSession()
        sections = [f"Section {i}" for i in range(12)]  # 12 sections

        questions = generate_section_questions(session, sections)

        # Budget is 8 → should cap at 8 questions
        assert len(questions) <= 8, f"Expected ≤8 questions, got {len(questions)}"
        assert session.budget_remaining >= 0
        assert session.total_asked == 8

    def test_gap_questions_respects_budget(self):
        """Gap-fill questions should not exceed remaining budget."""
        session = QuestionSession()

        # Consume 5 budget first
        session.questions_asked = [
            Question(id=i, text=f"Q{i}", category="structure")
            for i in range(1, 6)
        ]
        session.budget_remaining = 3

        gaps = {
            "Section A": ["no_examples", "missing_detail", "no_takeaway"],
            "Section B": ["missing_detail", "missing_transition"],
            "Section C": ["no_examples"],
        }

        questions = generate_gap_questions(session, gaps)

        # Should generate at most 3 questions (remaining budget)
        assert len(questions) <= 3, f"Expected ≤3 gap questions, got {len(questions)}"
        assert session.budget_remaining >= 0

    def test_exhausted_budget_no_more_questions(self):
        """With 0 budget, neither section nor gap questions should be generated."""
        session = QuestionSession()
        session.budget_remaining = 0

        q1 = generate_section_questions(session, ["A", "B"])
        assert len(q1) == 0, "Should not generate section questions with 0 budget"

        q2 = generate_gap_questions(session, {"A": ["no_examples"]})
        assert len(q2) == 0, "Should not generate gap questions with 0 budget"

    def test_can_ask_accurately_reflects_budget(self):
        """can_ask() should track budget changes."""
        session = QuestionSession()
        assert session.can_ask() is True

        # Ask 8 questions
        for i in range(8):
            q = Question(id=i + 1, text=f"Q{i+1}", category="detail")
            session.mark_asked(q)

        assert session.can_ask() is False
        assert session.budget_remaining == 0
        assert session.total_asked == 8


# ---------------------------------------------------------------------------
# Test 6: Full gatherer pipeline
# ---------------------------------------------------------------------------


class TestFullGathererPipeline:
    """ContentGatherer end-to-end pipeline validation."""

    def test_pipeline_produces_valid_outline(self, detailed_input: str, tmp_path: Path):
        """gather() with skip_questions produces a valid ContentOutline."""
        from ppt_skill.content.gatherer import ContentGatherer

        gatherer = ContentGatherer()
        outline = gatherer.gather(detailed_input, mode="skip_questions")

        # Must return a ContentOutline
        assert isinstance(outline, ContentOutline)

        # Must have a presentation title
        assert outline.presentation_title, "Presentation title should not be empty"

        # Must have slides
        assert len(outline.slides) > 0, f"Expected slides, got {len(outline.slides)}"

        # First slide should be TITLE layout
        first_slide = outline.slides[0]
        assert first_slide.layout_type == OutlineLayoutType.TITLE, \
            f"First slide should be TITLE, got {first_slide.layout_type}"

        # Validate — should pass
        issues = outline.validate()
        assert issues == [], f"Expected no validation issues, got: {issues}"

        # Save to YAML
        output_path = gatherer.save(str(tmp_path))
        assert output_path.is_file(), f"YAML file not created at {output_path}"

        # Verify YAML is parseable
        with open(output_path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f)
        assert data is not None
        assert "presentation_title" in data
        assert "slides" in data
        assert len(data["slides"]) == len(outline.slides)

    def test_pipeline_sections_identified(self, detailed_input: str):
        """Multi-section input should identify sections."""
        from ppt_skill.content.gatherer import ContentGatherer

        gatherer = ContentGatherer()
        outline = gatherer.gather(detailed_input, mode="skip_questions")

        # The detailed_input has 3 explicit sections
        # At minimum, some sections should be detected
        assert len(outline.sections) >= 1, \
            f"Expected at least 1 section, got {outline.sections}"

    def test_empty_input_raises_error(self):
        """Empty input should raise ValueError."""
        from ppt_skill.content.gatherer import ContentGatherer

        gatherer = ContentGatherer()
        with pytest.raises(ValueError, match="non-empty"):
            gatherer.gather("")

        with pytest.raises(ValueError, match="non-empty"):
            gatherer.gather("   ")

    def test_save_without_gather_raises(self, tmp_path: Path):
        """Calling save() before gather() should raise ValueError."""
        from ppt_skill.content.gatherer import ContentGatherer

        gatherer = ContentGatherer()
        with pytest.raises(ValueError, match="No outline to save"):
            gatherer.save(str(tmp_path))

    def test_minimal_input_produces_outline(self):
        """Very minimal input should still produce an outline (best-effort)."""
        from ppt_skill.content.gatherer import ContentGatherer

        gatherer = ContentGatherer()
        outline = gatherer.gather("Hello World", mode="skip_questions")

        assert isinstance(outline, ContentOutline)
        assert len(outline.slides) > 0, "Even minimal input should produce at least 1 slide"


# ---------------------------------------------------------------------------
# Test 7: YAML persistence (save and load)
# ---------------------------------------------------------------------------


class TestOutlineSaveAndLoad:
    """ContentGatherer.save() and load_outline() round-trip."""

    def test_save_and_load_round_trip(self, detailed_input: str, tmp_path: Path):
        """save() → load_outline() preserves all key fields."""
        from ppt_skill.content.gatherer import ContentGatherer

        gatherer = ContentGatherer()
        outline = gatherer.gather(detailed_input, mode="skip_questions")

        # Save to YAML
        output_path = gatherer.save(str(tmp_path))

        # Get the name from metadata or filename
        name = output_path.stem

        # Load back
        loaded = ContentGatherer.load_outline(name, str(tmp_path))

        assert loaded.presentation_title == outline.presentation_title
        assert len(loaded.slides) == len(outline.slides)

        # Verify first slide layout type matches
        assert loaded.slides[0].layout_type == outline.slides[0].layout_type

        # Verify slide bodies are preserved
        for i, (orig, load) in enumerate(zip(outline.slides, loaded.slides)):
            assert load.title == orig.title, f"Slide {i+1}: title mismatch"
            assert load.body == orig.body, f"Slide {i+1}: body mismatch"
            assert load.layout_type == orig.layout_type, \
                f"Slide {i+1}: layout_type mismatch"

    def test_load_nonexistent_outline_raises(self):
        """Loading a nonexistent outline should raise FileNotFoundError."""
        from ppt_skill.content.gatherer import ContentGatherer

        with pytest.raises(FileNotFoundError):
            ContentGatherer.load_outline("nonexistent-file-xyz", "outlines")


# ---------------------------------------------------------------------------
# Test 8: Layout type heuristics
# ---------------------------------------------------------------------------


class TestLayoutTypeHeuristics:
    """Layout type determination logic."""

    def test_default_layout_is_content(self):
        """Neutral content should default to CONTENT layout."""
        from ppt_skill.content.gatherer import ContentGatherer

        gatherer = ContentGatherer()
        lt = gatherer._determine_layout_type(
            ["Regular bullet point about business strategy and operations"]
        )
        assert lt == OutlineLayoutType.CONTENT

    def test_comparison_detected(self):
        """Comparison keywords → TWO_COLUMN."""
        from ppt_skill.content.gatherer import ContentGatherer

        gatherer = ContentGatherer()
        lt = gatherer._determine_layout_type(
            ["Pros of option A", "Cons of option A", "This is a comparison of two approaches"]
        )
        assert lt == OutlineLayoutType.TWO_COLUMN

    def test_image_content_detected(self):
        """Image-related keywords → IMAGE_TEXT."""
        from ppt_skill.content.gatherer import ContentGatherer

        gatherer = ContentGatherer()
        lt = gatherer._determine_layout_type(
            ["Screenshot of the new dashboard", "Diagram showing the architecture"]
        )
        assert lt == OutlineLayoutType.IMAGE_TEXT

    def test_data_content_detected(self):
        """Data/statistics keywords → DATA."""
        from ppt_skill.content.gatherer import ContentGatherer

        gatherer = ContentGatherer()
        lt = gatherer._determine_layout_type(
            ["Revenue chart for Q3", "Statistics showing 15% growth", "Key metrics"]
        )
        assert lt == OutlineLayoutType.DATA
