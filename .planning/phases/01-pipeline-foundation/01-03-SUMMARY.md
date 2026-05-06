---
phase: 01-pipeline-foundation
plan: 03
subsystem: pipeline
tags: [pipeline, cli, testing, pytest, quality-checker, converter, e2e, pptx, native-shapes]

# Dependency graph
requires:
  - phase: 01-01
    provides: Forked converter + finalize modules with fixed imports
  - phase: 01-02
    provides: SVGQualityChecker, config, icon registry, chart templates
provides:
  - Unified pipeline CLI: convert_svg_to_pptx() + argparse --input/--output/--skip-check
  - 23 regression tests covering quality checker, converter imports, and e2e
  - 9 fixture SVGs: 4 clean + 5 banned features
  - Installable package via pyproject.toml (pip install -e .)
affects: [Phase 2 content-spec, Phase 3 composition, Phase 4 animation]

# Tech tracking
tech-stack:
  added: [pytest, argparse]
  patterns: [function-based pipeline with quality gate, temp_path fixtures for e2e, class-based test organization]

key-files:
  created:
    - src/ppt_skill/pipeline.py
    - pyproject.toml
    - tests/__init__.py
    - tests/conftest.py
    - tests/test_quality_checker.py
    - tests/test_converter.py
    - tests/test_e2e.py
    - tests/fixtures/sample_simple.svg
    - tests/fixtures/sample_text.svg
    - tests/fixtures/sample_gradient.svg
    - tests/fixtures/sample_icon.svg
    - tests/fixtures/banned_features_svg/mask.svg
    - tests/fixtures/banned_features_svg/rgba.svg
    - tests/fixtures/banned_features_svg/fontface.svg
    - tests/fixtures/banned_features_svg/html_entities.svg
    - tests/fixtures/banned_features_svg/style_tag.svg
  modified: []

key-decisions:
  - "Unified pipeline as single function (convert_svg_to_pptx) with quality gate and optional skip flag"
  - "CLI uses argparse with exactly 2 required args + 1 optional flag per Phase 1 scope"
  - "use_native_shapes=True and use_compat_mode=False as pipeline defaults (Phase 1 targets native shapes)"
  - "create_pptx_with_native_svg called with verbose=False in pipeline for clean CLI output"

patterns-established:
  - "Pipeline pattern: validate → quality-check → convert → output with clear error propagation"
  - "Test fixture pattern: minimal SVGs (4 clean, 5 banned) for targeted regression testing"

requirements-completed: [PIP-01, PIP-02, PIP-03]

# Metrics
duration: 5min
completed: 2026-05-06
---

# Phase 1 Plan 3: Pipeline Integration & Regression Tests Summary

**Unified SVG-to-PPTX pipeline with quality gate, CLI, and 23 passing regression tests across quality checker, converter imports, and end-to-end validation**

## Performance

- **Duration:** 5 min
- **Started:** 2026-05-06T14:55:39Z
- **Completed:** 2026-05-06T15:01:19Z
- **Tasks:** 3
- **Files modified:** 16 (all new)

## Accomplishments
- Unified pipeline CLI (`convert_svg_to_pptx`) orchestrating quality check → native-shape conversion
- 9 minimal fixture SVGs: 4 clean (simple, text, gradient, icon) + 5 banned features (mask, rgba, fontface, html_entities, style_tag)
- 23 regression tests passing: 8 quality checker + 10 converter imports/utils + 5 end-to-end
- Pipeline verified end-to-end: SVG → 28KB PPTX with native auto-shape on slide

## Task Commits

Each task was committed atomically:

1. **Task 1: Create unified pipeline entry point with CLI** - `5f2ba20` (feat)
2. **Task 2: Create test fixtures — banned SVGs + clean SVGs** - `78214da` (test)
3. **Task 3: Write regression tests — quality checker + converter + e2e** - `07f9f09` (test)

## Files Created/Modified
- `src/ppt_skill/pipeline.py` - Unified pipeline: convert_svg_to_pptx() + argparse CLI with --input/--output/--skip-check
- `pyproject.toml` - Package configuration for pip install -e . (blocking fix)
- `tests/conftest.py` - Shared pytest fixtures (fixtures_dir, banned_svg_dir, temp_output)
- `tests/test_quality_checker.py` - 8 tests: 5 banned feature rejections + 3 clean SVG acceptances
- `tests/test_converter.py` - 10 tests: converter + finalize imports + hex color util
- `tests/test_e2e.py` - 5 integration tests: single, multi-slide, text, rejection, skip-check
- `tests/fixtures/sample_*.svg` - 4 clean test SVGs (simple, text, gradient, icon)
- `tests/fixtures/banned_features_svg/*.svg` - 5 intentionally-invalid test SVGs

## Decisions Made
- `use_native_shapes=True` and `use_compat_mode=False` as pipeline defaults — Phase 1 targets native shapes exclusively
- `verbose=False` passed to builder for clean CLI output
- Import names adapted from plan to match actual module APIs (`parse_svg_path` not `parse_path_d`, `resolve_icon_path` not `resolve_data_icon`, `flatten_text_with_tspans` not `flatten_tspans`)
- Created pyproject.toml for package installability — required for `from ppt_skill.pipeline` imports in tests

## Deviations from Plan

### Auto-fixed Issues

**1. [Rule 3 - Blocking] Created pyproject.toml for package installability**
- **Found during:** Task 1 (pipeline.py creation)
- **Issue:** No setup.py or pyproject.toml existed — `pip install -e .` impossible, package imports failing
- **Fix:** Created minimal pyproject.toml with setuptools build-backend and src/ package discovery
- **Files modified:** pyproject.toml (new)
- **Verification:** `python -c "from ppt_skill.pipeline import convert_svg_to_pptx"` works
- **Committed in:** `5f2ba20` (Task 1 commit)

**2. [Rule 1 - Bug] Fixed test import names to match actual module APIs**
- **Found during:** Task 3 (test file creation)
- **Issue:** Plan specified imports that don't exist: `parse_path_d` (actual: `parse_svg_path`), `resolve_data_icon` (actual: `resolve_icon_path`), `flatten_tspans` (actual: `flatten_text_with_tspans`)
- **Fix:** Used actual module exports verified via runtime introspection
- **Files modified:** tests/test_converter.py
- **Verification:** All 10 converter tests pass
- **Committed in:** `07f9f09` (Task 3 commit)

**3. [Rule 1 - Bug] html_entities.svg intentionally fails XML parsing**
- **Found during:** Task 2 (fixture verification)
- **Issue:** Plan verification says "all 9 SVGs must parse as well-formed XML", but html_entities.svg contains `&mdash;` and `&nbsp;` which are HTML entities, not XML entities — intentionally designed to trigger parse errors detected by quality checker
- **Fix:** None needed — fixture content matches plan exactly, and test_quality_checker.py correctly asserts `len(result['errors']) > 0`
- **Files modified:** None (fixture preserved as-is per plan)
- **Verification:** `test_rejects_html_entities` passes

---

**Total deviations:** 3 auto-fixed (1 blocking, 2 bugs)
**Impact on plan:** All auto-fixes necessary for correctness. No scope creep.

## Issues Encountered
- `create_pptx_with_native_svg` signature has more parameters than plan assumed (16 total) — adapted pipeline call to pass `use_native_shapes=True`, `use_compat_mode=False`, `verbose=False`
- Plan verification statement about XML parsing contradicted the html_entities fixture design — documented as expected behavior

## User Setup Required
None — no external service configuration required.

## Next Phase Readiness
- Phase 1 complete: Pipeline foundation delivers working SVG→PPTX with quality validation and 23 regression tests
- Ready for Phase 2 (content spec): template extraction pipeline has a verified converter foundation to build upon
- No blockers or concerns

## Self-Check: PASSED

- 15/15 key files exist on disk
- 3/3 01-03 commits found in git log (5f2ba20, 78214da, 07f9f09)
- All 23 tests pass: `python -m pytest tests/ -v`
- CLI pipeline verified: SVG → valid PPTX with native shapes

---
*Phase: 01-pipeline-foundation*
*Completed: 2026-05-06*
