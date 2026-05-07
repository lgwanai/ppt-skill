"""CLI content gathering commands — gather-content, generate-outline, list-outlines.

All functions use pathlib.Path for paths, write to stdout for user-facing
output, and write to stderr for errors. No interactive stdin prompts —
these are CLI commands designed for both human and programmatic consumption.

Functions are callable directly from Python (e.g., from Phase 4 code)
without requiring argparse. They will be wired to a CLI entry point
in Phase 5 (packaging).

Design follows the same stateless-function pattern as spec_commands.py.
"""

from __future__ import annotations

import sys
from pathlib import Path

import yaml

from ppt_skill.content.gatherer import ContentGatherer
from ppt_skill.content.model import ContentOutline, OutlineLayoutType
from ppt_skill.cli.spec_commands import get_active_spec


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def gather_content(
    topic: str,
    spec_name: str | None = None,
    skip_questions: bool = False,
    outlines_dir: str = "outlines",
) -> ContentOutline:
    """Gather content for a presentation and produce a ContentOutline.

    Instantiates ``ContentGatherer``, runs the full pipeline (sufficiency
    assessment → adaptive questioning → outline generation), saves the
    result to YAML, and prints a summary to stdout.

    Parameters
    ----------
    topic : str
        The presentation topic or content description.
    spec_name : str | None
        Optional spec name. If None, auto-resolves via ``get_active_spec()``.
    skip_questions : bool
        If True, passes ``mode="skip_questions"`` to bypass Phase 2.
    outlines_dir : str
        Directory for outline YAML output (default: ``"outlines"``).

    Returns
    -------
    ContentOutline
        The generated outline with validated slide entries.

    Raises
    ------
    ValueError
        If ``topic`` is empty or whitespace-only.
    """
    gatherer = ContentGatherer()

    # Resolve spec_name
    if spec_name is None:
        spec_name = get_active_spec()

    # Run pipeline
    mode = "skip_questions" if skip_questions else "auto"
    outline = gatherer.gather(topic, mode=mode, spec_name=spec_name)

    # Save to YAML
    output_path = gatherer.save(outlines_dir)

    # Print summary
    _print_gather_summary(outline, output_path)

    return outline


def generate_outline_from_summary(
    summary: str,
    title: str,
    spec_name: str | None = None,
    outlines_dir: str = "outlines",
) -> ContentOutline:
    """Generate a ContentOutline directly from a structured content summary.

    Bypasses the sufficiency assessment and questioning phases entirely —
    for programmatic use or when the caller has already gathered sufficient
    content.

    Parameters
    ----------
    summary : str
        Free-text description of all slides and their content.
    title : str
        Presentation title.
    spec_name : str | None
        Optional spec name. If None, auto-resolves via ``get_active_spec()``.
    outlines_dir : str
        Directory for outline YAML output (default: ``"outlines"``).

    Returns
    -------
    ContentOutline
        The generated outline.
    """
    gatherer = ContentGatherer()

    # Resolve spec_name
    if spec_name is None:
        spec_name = get_active_spec()

    # Run pipeline with skip_questions (content already gathered)
    outline = gatherer.gather(
        summary,
        mode="skip_questions",
        spec_name=spec_name,
    )

    # Override title if provided
    if title:
        outline.presentation_title = title

    # Save to YAML
    output_path = gatherer.save(outlines_dir)

    # Print summary
    _print_gather_summary(outline, output_path)

    return outline


def list_outlines(outlines_dir: str = "outlines") -> list[str]:
    """List all available content outlines in the outlines directory.

    Scans for ``*.yaml`` files, reads metadata from each, and prints a
    formatted list showing outline names with slide counts and creation dates.

    Parameters
    ----------
    outlines_dir : str
        Directory where outline YAML files are stored (default: ``"outlines"``).

    Returns
    -------
    list[str]
        List of outline names (without ``.yaml`` extension). Empty list if
        no outlines found.
    """
    outlines_path = Path(outlines_dir)

    if not outlines_path.is_dir():
        print(
            f"No outlines found in {outlines_dir}/ directory. "
            f"Use gather-content to create one.",
        )
        return []

    yaml_files = sorted(outlines_path.glob("*.yaml"))
    if not yaml_files:
        print(
            f"No outlines found in {outlines_dir}/ directory. "
            f"Use gather-content to create one.",
        )
        return []

    # Read metadata from each outline
    outline_entries: list[dict] = []
    for yf in yaml_files:
        try:
            with open(yf, "r", encoding="utf-8") as f:
                data = yaml.safe_load(f) or {}
        except Exception:
            data = {}

        if not isinstance(data, dict):
            data = {}

        metadata = data.get("metadata", {}) if isinstance(data, dict) else {}
        name = metadata.get("name", yf.stem)
        slide_count = metadata.get("slide_count", "?")
        saved_at = metadata.get("saved_at", "")

        # Format save date: truncate ISO timestamp to date only
        if saved_at and "T" in str(saved_at):
            saved_at = str(saved_at).split("T")[0]

        outline_entries.append({
            "name": name,
            "slide_count": slide_count,
            "saved_at": saved_at,
        })

    # Print table
    print("Available outlines:")
    for entry in outline_entries:
        name = entry["name"]
        sc = entry["slide_count"]
        date = entry["saved_at"] or "unknown"
        print(f"  \u2022 {name:<24} ({sc} slides, created {date})")

    return [e["name"] for e in outline_entries]


# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------


def _print_gather_summary(
    outline: ContentOutline,
    output_path: Path,
) -> None:
    """Print a human-readable summary of a completed gather operation.

    Layout distribution shows the count of each layout type across all
    slides in the generated outline.
    """
    title = outline.presentation_title or "Untitled"
    slide_count = len(outline.slides)
    section_count = len(outline.sections)

    # Count layout types
    layout_counts: dict[str, int] = {}
    for slide in outline.slides:
        lt = slide.layout_type.value
        layout_counts[lt] = layout_counts.get(lt, 0) + 1

    layout_dist = ", ".join(
        f"{k}: {v}" for k, v in sorted(layout_counts.items())
    )

    name = outline.metadata.get("name", output_path.stem)

    print(f'\u2713 Content outline "{title}" generated ({slide_count} slides, {section_count} sections)')
    print(f"  \u2192 Saved to {output_path}")
    if layout_dist:
        print(f"  Layout distribution: {layout_dist}")


# ---------------------------------------------------------------------------
# Module exports
# ---------------------------------------------------------------------------

__all__ = [
    "gather_content",
    "generate_outline_from_summary",
    "list_outlines",
]
