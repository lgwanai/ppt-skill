"""Sufficiency assessment — rubric-based evaluation of user input adequacy.

Determines whether user-provided content is detailed enough to skip the
adaptive questioning phase. Uses a 4-dimension scoring rubric (structure,
detail, audience, scope) with configurable thresholds.

This module is prompt-driven: the SUFFICIENCY_RUBRIC constant defines the
evaluation prompt that Claude (the AI runtime) follows. The Python functions
provide prompt formatting, result construction, and the assessment workflow.

Design follows project conventions:
  - from __future__ import annotations
  - type hints throughout
  - docstrings for all public functions
"""

from __future__ import annotations

from ppt_skill.content.model import SufficiencyResult


# ---------------------------------------------------------------------------
# Rubric prompt template
# ---------------------------------------------------------------------------

SUFFICIENCY_RUBRIC = """
Evaluate the following content input for a presentation. Score each dimension 0-2:

1. STRUCTURE (0-2): Are topics/sections clearly delineated?
   0 = no structure visible
   1 = some grouping implied (related topics together)
   2 = numbered/headered sections, explicit outline formatting, or slide-like structure

2. DETAIL (0-2): Does each section have supporting content?
   0 = topics only, no elaboration
   1 = some bullets or brief descriptions present
   2 = rich content with specific examples, data, or detailed explanations

3. AUDIENCE (0-2): Is target audience and purpose clear?
   0 = not stated or implied
   1 = vaguely implied by content type or domain
   2 = explicitly stated (e.g., "for executives", "training new hires")

4. SCOPE (0-2): Is the presentation scope bounded?
   0 = open-ended, unbounded topic
   1 = rough scope implied (time period, department, project)
   2 = well-defined, bounded scope

Threshold for SUFFICIENT:
- Total score >= 5
- STRUCTURE >= 1 (must have at least some organization)
- DETAIL >= 1 (must have at least some content depth)

Input:
{user_input}
"""


# ---------------------------------------------------------------------------
# Prompt formatting
# ---------------------------------------------------------------------------


def get_sufficiency_prompt(user_input: str) -> str:
    """Return the formatted sufficiency assessment prompt with embedded input.

    Args:
        user_input: Raw user-provided content text.

    Returns:
        The SUFFICIENCY_RUBRIC prompt with ``{user_input}`` replaced by the
        actual user input text.
    """
    return SUFFICIENCY_RUBRIC.replace("{user_input}", user_input)


# ---------------------------------------------------------------------------
# Result construction
# ---------------------------------------------------------------------------


def build_sufficiency_result(
    scores: dict[str, int],
    rationale: str,
    section_count: int,
    estimated_slide_count: int,
) -> SufficiencyResult:
    """Construct a SufficiencyResult from raw dimension scores.

    Applies the sufficiency threshold: ``total >= 5 AND structure >= 1 AND
    detail >= 1``. Confidence is ``total / 8.0`` (max possible = 8).

    Args:
        scores: Per-dimension scores, e.g. ``{"structure": 2, "detail": 1,
            "audience": 1, "scope": 1}``.
        rationale: Human-readable explanation of the assessment.
        section_count: How many sections/topics were identified.
        estimated_slide_count: Rough estimate of resulting slides.

    Returns:
        A populated SufficiencyResult dataclass with computed thresholds.
    """
    total = sum(scores.values())
    sufficient = (
        total >= 5
        and scores.get("structure", 0) >= 1
        and scores.get("detail", 0) >= 1
    )
    confidence = min(total / 8.0, 1.0)
    missing_dimensions = [dim for dim, score in scores.items() if score < 2]

    return SufficiencyResult(
        sufficient=sufficient,
        confidence=confidence,
        missing_dimensions=missing_dimensions,
        section_count=section_count,
        estimated_slide_count=estimated_slide_count,
        scores=scores,
        total_score=total,
        rationale=rationale,
    )


# ---------------------------------------------------------------------------
# Main assessment workflow
# ---------------------------------------------------------------------------


def assess_sufficiency(
    user_input: str,
    spec_context: dict | None = None,
) -> SufficiencyResult:
    """Assess whether user input is sufficient to skip adaptive questioning.

    Generates the sufficiency rubric prompt with the user's input embedded.
    The prompt is intended to be evaluated by Claude (the AI runtime), which
    returns structured dimension scores. Those scores are then assembled into
    a ``SufficiencyResult`` via :func:`build_sufficiency_result`.

    When no LLM evaluation is available (e.g., programmatic testing), the
    returned result uses default-zero scores and embeds the prompt text in
    the ``rationale`` field for downstream consumption.

    Args:
        user_input: Raw user-provided content text (may be multi-line).
        spec_context: Optional dict with spec metadata (e.g., slide type
            distribution) to enrich the prompt. If provided, the spec
            context is appended to the rubric prompt as additional context.

    Returns:
        A SufficiencyResult. When called without LLM scoring, scores are
        all zero (sufficient=False). The ``rationale`` field contains the
        formatted evaluation prompt that should be sent to the AI runtime.

    Edge cases:
        - Empty input → all scores 0, sufficient=False
        - Single-word input → structure=0, detail=0, audience=0, scope=0
        - Full outline → typically sufficient=True after LLM scoring
    """
    # Generate the evaluation prompt
    prompt = get_sufficiency_prompt(user_input)

    # If spec_context is available, enrich the prompt
    if spec_context:
        spec_note = "\n\nAdditional context from active design spec:\n"
        for key, value in spec_context.items():
            spec_note += f"  - {key}: {value}\n"
        prompt += spec_note

    # In the non-LLM path, return a placeholder result with the prompt
    # embedded in rationale. The orchestrator (or AI runtime) should
    # evaluate the prompt and call build_sufficiency_result() with
    # the returned scores.
    return SufficiencyResult(
        sufficient=False,
        confidence=0.0,
        missing_dimensions=["structure", "detail", "audience", "scope"],
        section_count=0,
        estimated_slide_count=0,
        scores={},
        total_score=0,
        rationale=prompt,
    )


# ---------------------------------------------------------------------------
# Module exports
# ---------------------------------------------------------------------------

__all__ = [
    "SUFFICIENCY_RUBRIC",
    "assess_sufficiency",
    "build_sufficiency_result",
    "get_sufficiency_prompt",
]
