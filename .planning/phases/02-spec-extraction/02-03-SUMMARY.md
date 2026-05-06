---
phase: 02-spec-extraction
plan: 03
subsystem: spec
tags: [yaml, cli, orchestration, testing, pyyaml, dataclasses]

# Dependency graph
requires:
  - phase: 02-01
    provides: "spec_model dataclasses, theme.py color/font extraction, config.py defaults"
  - phase: 02-02
    provides: "slide_classifier, layout_analysis, font_analysis, density analysis modules"
provides:
  - "SpecExtractor orchestrator that ties all extraction modules into a single pipeline"
  - "CLI spec management commands: extract_spec, list_specs, select_spec, get_active_spec"
  - "End-to-end YAML spec serialization with dataclass→dict conversion"
  - "7 integration tests validating the full extraction pipeline"
affects: [phase-03, phase-04, phase-05]

# Tech tracking
tech-stack:
  added: [pyyaml]
  patterns: ["Orchestrator pattern: SpecExtractor ties theme + classifier + layout + density + font_analysis modules",
             "CLI function-based pattern: standalone stateless functions callable from Python or argparse",
             "Dataclass serialization: asdict() + Enum .value extraction for clean YAML output",
             "Fallback chain: theme extraction → config.py DESIGN_COLORS defaults"]

key-files:
  created:
    - src/ppt_skill/spec/extractor.py
    - src/ppt_skill/cli/__init__.py
    - src/ppt_skill/cli/spec_commands.py
    - tests/test_spec_extraction.py
  modified:
    - requirements.txt
    - pyproject.toml

key-decisions:
  - "Used dataclasses.asdict() + Enum .value walk for YAML serialization — minimal code, no external deps beyond PyYAML"
  - "CLI functions are stateless and callable programmatically (Phase 3-4) without argparse"
  - "Font size extraction runs post-loop as aggregate pass over all slides (fa_mod.extract_all_slide_fonts + compute_spec_typography_sizes)"
  - "Theme color fallback uses config.py DESIGN_COLORS — 12 Office 365 defaults ensure no None crashes"
  - "Source config canvas detection uses slide width/height ratio: 1.70-1.85 → ppt169, 1.28-1.38 → ppt43"

patterns-established:
  - "SpecExtractor orchestration: 8-step pipeline (metadata → colors → fonts → slides → density → sizes → rhythm → config)"
  - "CLI output format: ✓ and ✗ markers, formatted table with active spec indicator"
  - "Test fixture: programmatic PPTX construction via python-pptx — no binary files committed"

requirements-completed: [SPC-04, SPC-05]

# Metrics
duration: 6min
completed: 2026-05-06
---

# Phase 2 Plan 3: SpecExtractor Orchestrator + CLI Commands Summary

**SpecExtractor orchestrator tying 7 extraction modules into end-to-end pipeline with YAML serialization and 3 CLI spec management commands**

## Performance

- **Duration:** 6 min
- **Started:** 2026-05-06T16:45:52Z
- **Completed:** 2026-05-06T16:51:44Z
- **Tasks:** 3
- **Files created/modified:** 6 (4 created, 2 modified)

## Accomplishments

- SpecExtractor orchestrator integrates theme.py, slide_classifier, layout_analysis, density, and font_analysis into a single `extract()` → `DesignSpec` pipeline
- CLI commands: `extract_spec` (extract + save YAML), `list_specs` (scan specs/ with metadata table), `select_spec` (write .active state file), `get_active_spec` (query active)
- 7 integration tests pass — colors (12-field ColorPalette with HEX), fonts (families + heading_sizes/body_sizes), classification (title/content/divider), density (breathing/dense/anchor), YAML round-trip, spec management (list/select/query)

## Task Commits

Each task was committed atomically:

1. **Task 1: Add PyYAML dependency and create SpecExtractor orchestrator** - `fb0e11e` (feat)
2. **Task 2: Create CLI spec commands (extract-spec, list-specs, select-spec)** - `2fc660a` (feat)
3. **Task 3: Full integration test — extract a real spec end-to-end** - `472f4c9` (test)

**Plan metadata:** (pending final commit)

## Files Created/Modified

- `src/ppt_skill/spec/extractor.py` - SpecExtractor orchestrator (351 lines) — 8-step pipeline
- `src/ppt_skill/cli/__init__.py` - CLI package init with function exports
- `src/ppt_skill/cli/spec_commands.py` - 4 CLI functions for spec management
- `tests/test_spec_extraction.py` - 7 integration tests using programmatic PPTX fixture
- `requirements.txt` - Added PyYAML>=6.0
- `pyproject.toml` - Added pyyaml>=6.0 to dependencies

## Decisions Made

- **YAML serialization via asdict() + Enum .value walk** — Uses stdlib `dataclasses.asdict()` as base, then recursively extracts `.value` from Enum members. No additional serialization library needed beyond PyYAML.
- **CLI functions are stateless** — Each function is a standalone, callable entity. They can be invoked directly from Python code (Phase 3–4 integration) or wired to argparse (Phase 5 packaging).
- **Font size extraction as post-loop aggregate pass** — `extract_all_slide_fonts()` and `compute_spec_typography_sizes()` run after the per-slide classification/density loop. This produces aggregate heading_sizes/body_sizes dicts for the Typography model.
- **Color fallback chain: theme → config.py** — When theme1.xml is missing or colors are unresolvable, DESIGN_COLORS (12 Office 365 hex defaults) provide safe fallback values.

## Deviations from Plan

None - plan executed exactly as written.

## Issues Encountered

None — all imports resolved correctly, all 7 integration tests passed on first run.

## User Setup Required

None — no external service configuration required.

## Next Phase Readiness

Phase 2 (Spec Extraction) is complete — all 3 plans executed. Ready for:
- Phase 3: Content structuring
- Phase 4: PPT generation from extracted specs

---

*Phase: 02-spec-extraction*
## Self-Check: PASSED

- All 7 key files exist on disk
- All 3 task commits confirmed: fb0e11e, 2fc660a, 472f4c9
- All 7 integration tests pass

---

*Completed: 2026-05-06*
