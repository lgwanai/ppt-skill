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

from ppt_skill.spec.enhanced_extractor import SpecExtractor

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
    extractor = SpecExtractor()
    spec = extractor.extract(Path(pptx_path))
    spec.metadata["name"] = name
    # Override the spec root to use the provided specs_dir
    spec.metadata["specs_dir"] = specs_dir
    extractor.save(spec, base_dir=specs_dir)
    return Path(specs_dir) / name

    slide_count = len(spec.slides)
    print(f'\u2713 Spec "{name}" extracted from {pptx_path} \u2192 {output_path} ({slide_count} slides)')
    return output_path


def list_specs(specs_dir: str = "specs") -> list[str]:
    """List all available design specs (supports both flat and directory formats).

    Returns spec names found from either:
      - Flat: specs/<name>.yaml (legacy format)
      - Directory: specs/<name>/spec.yaml (enhanced format)
    """
    spec_dir = Path(specs_dir)
    if not spec_dir.is_dir():
        print(f"No specs found in {specs_dir}/ directory. Use extract-spec to create one.")
        return []

    names: set[str] = set()

    # Legacy flat YAML format
    for yf in spec_dir.glob("*.yaml"):
        names.add(yf.stem)

    # Enhanced directory format
    for sub in spec_dir.iterdir():
        if sub.is_dir() and (sub / "spec.yaml").exists():
            names.add(sub.name)

    if not names:
        print(f"No specs found in {specs_dir}/ directory. Use extract-spec to create one.")
        return []

    spec_list = sorted(names)
    print(f"{'Name':<30} {'Slides':<8}")
    print("-" * 50)
    for name in spec_list:
        # Try to read metadata
        slide_count = "?"
        # Check directory format first
        spec_yaml = spec_dir / name / "spec.yaml"
        if spec_yaml.exists():
            try:
                data = yaml.safe_load(spec_yaml.read_text(encoding="utf-8"))
                meta = data.get("metadata", {}) if data else {}
                slide_count = str(meta.get("slide_count", "?"))
            except Exception:
                pass
        # Fall back to flat format
        flat_yaml = spec_dir / f"{name}.yaml"
        if slide_count == "?" and flat_yaml.exists():
            try:
                data = yaml.safe_load(flat_yaml.read_text(encoding="utf-8"))
                slide_count = str(data.get("metadata", {}).get("slide_count", "?")) if data else "?"
            except Exception:
                pass
        print(f"{name:<30} {slide_count:<8}")

    return spec_list


def select_spec(name: str, specs_dir: str = "specs") -> Path:
    """Set a spec as the active specification (supports both flat and directory formats)."""
    spec_dir = Path(specs_dir)
    spec_path = spec_dir / f"{name}.yaml"
    spec_dir_path = spec_dir / name / "spec.yaml"

    if not spec_path.is_file() and not spec_dir_path.is_file():
        # List available specs for helpful error
        available = []
        if spec_dir.is_dir():
            available = sorted(
                [f.stem for f in spec_dir.glob("*.yaml") if f.is_file()]
                + [d.name for d in spec_dir.iterdir() if d.is_dir() and (d / "spec.yaml").exists()]
            )
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
