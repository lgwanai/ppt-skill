"""CLI module — spec management commands for the ppt-skill tool.

Provides extract-spec, list-specs, select-spec commands and the
get_active_spec query. All functions are callable both from a CLI
entry point (Phase 5) and programmatically from Phase 3–4.
"""

from ppt_skill.cli.spec_commands import (
    extract_spec,
    get_active_spec,
    list_specs,
    select_spec,
)

__all__ = [
    "extract_spec",
    "get_active_spec",
    "list_specs",
    "select_spec",
]
