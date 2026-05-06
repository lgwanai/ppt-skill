"""CLI spec management commands — extract, list, select, and query active specs.

All functions use pathlib.Path for paths, write to stdout for user-facing
output, and write to stderr for errors. No interactive stdin prompts —
these are CLI commands designed for both human and programmatic consumption.

Functions are callable directly from Python (e.g., from Phase 3–4 code)
without requiring argparse. They will also be wired to a CLI entry point
in Phase 5 (packaging).
"""

from __future__ import annotations

import sys
from pathlib import Path

import yaml

from ppt_skill.spec.extractor import SpecExtractor

# Filename for the active spec marker
_ACTIVE_FILE = ".active"


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def extract_spec(name: str, pptx_path: str, specs_dir: str = "specs") -> Path:
    """Extract a design spec from a PPTX file and save as YAML.

    Args:
        name: Spec name (used as the YAML filename stem, e.g. "corporate-blue").
        pptx_path: Path to the source .pptx file.
        specs_dir: Directory where spec YAML files are stored (default: "specs").

    Returns:
        Path to the written YAML spec file.

    Example::

        path = extract_spec("corporate-blue", "deck.pptx")
        print(path)  # → specs/corporate-blue.yaml
    """
    extractor = SpecExtractor(pptx_path, name)
    spec = extractor.extract()
    output_path = extractor.save(specs_dir)

    slide_count = len(spec.slides)
    print(f'\u2713 Spec "{name}" extracted from {pptx_path} \u2192 {output_path} ({slide_count} slides)')
    return output_path


def list_specs(specs_dir: str = "specs") -> list[str]:
    """List all available design specs in the specs directory.

    Scans for *.yaml files, reads metadata from each, and prints a
    formatted table showing spec names with slide counts and extraction
    dates.

    Args:
        specs_dir: Directory where spec YAML files are stored (default: "specs").

    Returns:
        List of spec names (without .yaml extension). Empty list if none found.
    """
    spec_dir = Path(specs_dir)
    if not spec_dir.is_dir():
        print(f"No specs found in {specs_dir}/ directory. Use extract-spec to create one.")
        return []

    yaml_files = sorted(spec_dir.glob("*.yaml"))
    if not yaml_files:
        print(f"No specs found in {specs_dir}/ directory. Use extract-spec to create one.")
        return []

    # Read metadata from each spec
    spec_entries: list[dict] = []
    for yf in yaml_files:
        try:
            with open(yf, "r", encoding="utf-8") as f:
                data = yaml.safe_load(f) or {}
        except Exception:
            data = {}

        metadata = data.get("metadata", {}) if isinstance(data, dict) else {}
        name = metadata.get("name", yf.stem)
        slide_count = metadata.get("slide_count", "?")
        extracted_at = metadata.get("extracted_at", "")

        # Format extraction date: truncate ISO timestamp to date only
        if extracted_at and "T" in str(extracted_at):
            extracted_at = str(extracted_at).split("T")[0]

        spec_entries.append({
            "name": name,
            "slide_count": slide_count,
            "extracted_at": extracted_at,
        })

    # Determine active spec
    active_name = get_active_spec(specs_dir)

    # Print table
    print("Available specs:")
    for entry in spec_entries:
        name = entry["name"]
        sc = entry["slide_count"]
        date = entry["extracted_at"] or "unknown"
        marker = " * " if name == active_name else "   "
        print(f"  {marker}{name:<24} ({sc} slides, extracted {date})")

    if active_name:
        print(f"\nActive: {active_name}")

    return [e["name"] for e in spec_entries]


def select_spec(name: str, specs_dir: str = "specs") -> Path:
    """Set a spec as the active specification.

    Writes the spec name to a .active file inside specs_dir. This file
    is a project-local state file (not version-controlled).

    Args:
        name: Spec name to activate (must correspond to specs_dir/<name>.yaml).
        specs_dir: Directory where spec YAML files are stored (default: "specs").

    Returns:
        Path to the .active file.

    Raises:
        FileNotFoundError: If the spec file specs_dir/<name>.yaml doesn't exist.
    """
    spec_dir = Path(specs_dir)
    spec_path = spec_dir / f"{name}.yaml"

    if not spec_path.is_file():
        # List available specs for a helpful error
        available = sorted(
            [f.stem for f in spec_dir.glob("*.yaml") if f.is_file()]
        ) if spec_dir.is_dir() else []
        print(
            f"\u2717 Spec '{name}' not found.",
            f"Available: {', '.join(available)}" if available else "No specs available.",
            file=sys.stderr,
        )
        raise FileNotFoundError(f"Spec '{name}' not found in {specs_dir}/")

    active_file = spec_dir / _ACTIVE_FILE
    active_file.write_text(name + "\n", encoding="utf-8")
    print(f"\u2713 Active spec set to: {name}")
    return active_file


def get_active_spec(specs_dir: str = "specs") -> str | None:
    """Get the name of the currently active spec.

    Reads the .active file from the specs directory. Returns None
    silently if the file doesn't exist (not an error — no spec selected yet).

    Args:
        specs_dir: Directory where spec YAML files are stored (default: "specs").

    Returns:
        Active spec name as a string, or None if no .active file exists.
    """
    active_file = Path(specs_dir) / _ACTIVE_FILE
    if not active_file.is_file():
        return None

    try:
        name = active_file.read_text(encoding="utf-8").strip()
        return name if name else None
    except Exception:
        return None


__all__ = [
    "extract_spec",
    "get_active_spec",
    "list_specs",
    "select_spec",
]
