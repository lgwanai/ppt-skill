"""Adaptive questioning — section-first + gap-fill with 8-question budget.

Implements the smart hybrid questioning strategy (GEN-01):
  1. Section-level overview questions (Phase 1): one per section, consumes budget.
  2. Gap-fill questions (Phase 2): allocates remaining budget to sections with most gaps first.

Both phases enforce a hard 8-question cap via QuestionSession.budget_remaining.

Design follows project conventions:
  - from __future__ import annotations
  - type hints throughout
  - docstrings for all public functions
"""

from __future__ import annotations

from ppt_skill.content.model import Question, QuestionSession

# ── Question category priority ──────────────────────────────────────────

QUESTION_CATEGORY_PRIORITY = [
    "structure",     # What are the main sections/topics?
    "detail",        # What content goes in each section?
    "audience",      # Who is this for, what's the takeaway?
    "storytelling",  # What's the narrative arc?
]

# ── Section-level overview question templates ───────────────────────────

SECTION_QUESTION_TEMPLATES = {
    "structure": "What key topics or points should the '{section}' section cover?",
    "detail": "What specific content, examples, or data should appear in '{section}'?",
    "audience": "What should the audience understand or do after the '{section}' section?",
    "storytelling": "How should '{section}' transition from the previous section?",
}

# ── Gap-fill question templates ─────────────────────────────────────────

GAP_QUESTION_TEMPLATES = {
    "missing_detail": "Can you provide 2-3 specific points or examples for '{section}'?",
    "no_examples": "What concrete example or case study illustrates '{section}'?",
    "no_takeaway": "What's the single key takeaway for the '{section}' section?",
    "missing_transition": "How does '{section}' connect to the next topic?",
}

# ── Public functions ────────────────────────────────────────────────────


def generate_section_questions(
    session: QuestionSession,
    sections: list[str],
    context: str = "",
) -> list[Question]:
    """Phase 1: Ask one overview question per identified section.

    Consumes from the question budget. First section gets a "structure"
    question; subsequent sections get "detail" questions. Stops early if
    the budget is exhausted.

    Args:
        session: The QuestionSession tracking budget and asked questions.
        sections: Ordered list of section/topic names.
        context: Optional description of why these questions are being asked.

    Returns:
        List of generated Question objects (may be fewer than len(sections)
        if budget ran out).
    """
    results: list[Question] = []

    for i, section in enumerate(sections):
        if not session.can_ask():
            break

        category = "structure" if i == 0 else "detail"
        template = SECTION_QUESTION_TEMPLATES.get(
            category, SECTION_QUESTION_TEMPLATES["detail"]
        )
        text = template.format(section=section)

        question = Question(
            id=len(session.questions_asked) + 1,
            category=category,
            text=text,
            target_section=section,
            context=context or f"Establishing content for presentation section",
        )
        results.append(question)
        session.mark_asked(question)

    return results


def generate_gap_questions(
    session: QuestionSession,
    gaps: dict[str, list[str]],
    context: str = "",
) -> list[Question]:
    """Phase 2: Allocate remaining budget to sections with most gaps.

    Sections are prioritized by gap count (descending). For each section,
    the first unresolved gap type is selected and a question is generated.

    Args:
        session: The QuestionSession tracking budget and asked questions.
        gaps: Mapping of section name → list of gap type strings
            (e.g. ``{"Results": ["no_examples", "no_takeaway"]}``).
        context: Optional description of why these questions are being asked.

    Returns:
        List of generated Question objects (may be empty if budget is
        exhausted or no gaps exist).
    """
    results: list[Question] = []

    # Work on a mutable copy to track unresolved gaps
    remaining_gaps: dict[str, list[str]] = {
        section: list(gap_list) for section, gap_list in gaps.items()
    }

    # Round-robin: each iteration picks the section with the most remaining gaps
    while session.can_ask() and any(remaining_gaps.values()):
        # Pick section with most gaps (stable sort for determinism)
        sections_sorted = sorted(
            remaining_gaps.keys(),
            key=lambda s: len(remaining_gaps[s]),
            reverse=True,
        )
        any_progress = False

        for section in sections_sorted:
            if not session.can_ask():
                break

            gap_list = remaining_gaps.get(section, [])
            if not gap_list:
                continue

            # Take first gap, look up template, fall back to missing_detail
            gap_type = gap_list.pop(0)
            template = GAP_QUESTION_TEMPLATES.get(
                gap_type, GAP_QUESTION_TEMPLATES["missing_detail"]
            )
            text = template.format(section=section)

            question = Question(
                id=len(session.questions_asked) + 1,
                category="detail",
                text=text,
                target_section=section,
                context=context or f"Filling content gap: {gap_type} for {section}",
            )
            results.append(question)
            session.mark_asked(question)
            any_progress = True

        if not any_progress:
            break

    return results


def identify_content_gaps(
    known_content: dict[str, list[str]],
    sections: list[str],
) -> dict[str, list[str]]:
    """Analyze known content and identify gaps per section.

    Gaps are determined by the amount of known content for each section:
      - No content → ``["missing_detail", "no_examples", "no_takeaway"]``
      - 1–2 points → ``["no_examples", "no_takeaway"]``
      - ≥3 points with no explicit examples → ``["no_examples"]``
      - ≥3 points with examples → ``["no_takeaway"]`` (unless takeaway present)
      - First and last sections also get ``"missing_transition"`` if applicable.

    Args:
        known_content: Mapping of section name → list of known facts/points.
        sections: Ordered list of section names (used for transition gaps).

    Returns:
        Dict of section name → list of gap type strings.
    """
    gaps: dict[str, list[str]] = {}

    for section in sections:
        points = known_content.get(section, [])
        gap_list: list[str] = []

        if not points:
            gap_list = ["missing_detail", "no_examples", "no_takeaway"]
        elif len(points) <= 2:
            gap_list = ["no_examples", "no_takeaway"]
        elif len(points) >= 3:
            # Check for examples/data indicators in the content
            has_example = any(
                any(keyword in p.lower() for keyword in ("example", "e.g.", "case study", "instance"))
                for p in points
            )
            has_takeaway = any(
                any(keyword in p.lower() for keyword in ("key takeaway", "takeaway", "summary", "conclusion"))
                for p in points
            )
            if not has_example:
                gap_list.append("no_examples")
            if not has_takeaway:
                gap_list.append("no_takeaway")

        gaps[section] = gap_list

    # Add transition gaps for boundary sections
    if len(sections) >= 2:
        if sections[0] not in gaps:
            gaps[sections[0]] = []
        if "missing_transition" not in gaps[sections[0]]:
            gaps[sections[0]].append("missing_transition")

        if sections[-1] not in gaps:
            gaps[sections[-1]] = []
        if "missing_transition" not in gaps[sections[-1]]:
            gaps[sections[-1]].append("missing_transition")

    return gaps


def get_total_questions_asked(session: QuestionSession) -> int:
    """Return the total number of questions asked in this session."""
    return session.total_asked


__all__ = [
    "GAP_QUESTION_TEMPLATES",
    "QUESTION_CATEGORY_PRIORITY",
    "SECTION_QUESTION_TEMPLATES",
    "generate_gap_questions",
    "generate_section_questions",
    "get_total_questions_asked",
    "identify_content_gaps",
]
