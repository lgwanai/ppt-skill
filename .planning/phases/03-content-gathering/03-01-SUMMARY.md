---
phase: 03-content-gathering
plan: 01
subsystem: content
tags: [python, dataclass, enum, yaml, content-outline, questioning, sufficiency]

# Dependency graph
requires:
  - phase: 02-spec-extraction
    provides: _dataclass_to_dict serialization helper from ppt_skill.spec.extractor
provides:
  - ContentOutline dataclass with validate/to_dict/from_dict — Phase 4 contract for slide-by-slide content
  - OutlineLayoutType enum (6 layout values) — layout recommendations per slide
  - SlideEntry dataclass — per-slide content (title, body, layout, notes, image_hint, section)
  - SufficiencyResult dataclass — scored input assessment with confidence and missing dimensions
  - Question and QuestionSession dataclasses — adaptive questioning budget (8 max) with gap tracking
affects:
  - 03-content-gathering
  - 04-ppt-generation

# Tech tracking
tech-stack:
  added: []
  patterns:
    - Dataclass-based data models with all-default fields following spec_model.py conventions
    - str-based Enum with .value extraction via _dataclass_to_dict for YAML serialization
    - @classmethod from_dict() factory pattern for YAML deserialization round-trip
    - validate() returning list[str] (empty = valid) — same pattern as Phase 2 validation

key-files:
  created:
    - src/ppt_skill/content/__init__.py - Package init with __all__ exports
    - src/ppt_skill/content/model.py - All 6 content data model types
  modified: []

key-decisions:
  - "Used dataclasses (NOT Pydantic) for content model — matching Phase 2 convention, zero new dependencies"
  - "OutlineLayoutType has 6 values — extends Phase 2's SlideType (5 values) with TWO_COLUMN"
  - "QuestionSession budget starts at 8 with mark_asked() decrement — enforces max questioning limit"
  - "_dataclass_to_dict reused from ppt_skill.spec.extractor — same serialization pattern for enums and nested dataclasses"
  - "Sufficiency threshold documented but NOT hardcoded in dataclass — enforced in sufficiency.py module"

patterns-established:
  - "Content contract pattern: to_dict() validates first, raises ValueError on invalid, then delegates to _dataclass_to_dict"
  - "Deserialization pattern: from_dict() filters to __dataclass_fields__, converts string enum values, handles nested SlideEntry"

requirements-completed: [GEN-01, GEN-02, GEN-03]

# Metrics
duration: 7 min
completed: 2026-05-06
---

# Phase 3 Plan 1: Content Data Model Summary

**6 dataclass content data model with OutlineLayoutType enum, Phase 4 YAML contract via validate/to_dict/from_dict, and adaptive questioning types with 8-question budget tracking**

## Performance

- **Duration:** ~7 min
- **Started:** 2026-05-06T22:18:10Z
- **Completed:** 2026-05-06T22:24:42Z
- **Tasks:** 2
- **Files modified:** 2

## Accomplishments
- OutlineLayoutType enum with 6 layout values (extends Phase 2 SlideType with TWO_COLUMN)
- ContentOutline with 8 fields, validate() (5 rules), to_dict(), to_yaml(), from_dict() round-trip
- SlideEntry with 6 fields capturing per-slide: title, body bullets, layout_type, notes, image_hint, section_name
- SufficiencyResult with 8 fields for scored input assessment with confidence and missing dimensions
- Question and QuestionSession with 8-question budget, can_ask(), mark_asked(), total_asked property
- Full YAML serialization/deserialization reusing _dataclass_to_dict from Phase 2

## Task Commits

Each task was committed atomically:

1. **Task 1: Create content package and Phase 4 contract dataclasses (OutlineLayoutType, SlideEntry, ContentOutline)** - `e2a8f00` (feat)
2. **Task 2: Create sufficiency and questioning dataclasses (SufficiencyResult, Question, QuestionSession)** - `13a0e81` (feat)

**Plan metadata:** (to be committed after SUMMARY, STATE, ROADMAP updates)

## Files Created/Modified
- `src/ppt_skill/content/__init__.py` - Package init exporting all 6 types via __all__
- `src/ppt_skill/content/model.py` - 353 lines: 6 dataclasses, 1 enum, serialization, validation

## Decisions Made
- Used dataclasses (NOT Pydantic) for content model — matching Phase 2 convention, zero new dependencies
- OutlineLayoutType has 6 values — extends Phase 2's SlideType (5 values) with TWO_COLUMN for flexible layouts
- QuestionSession budget starts at 8 with mark_asked() decrement — enforces max questioning limit
- _dataclass_to_dict reused from ppt_skill.spec.extractor — same serialization pattern for enums and nested dataclasses
- Sufficiency threshold (total_score >= 5 AND structure >= 1 AND detail >= 1) documented but NOT hardcoded in dataclass — enforced in sufficiency.py module

## Deviations from Plan

### Auto-fixed Issues

**1. [Rule 1 - Bug] Fixed Task 1 verification script importing Task 2 types**
- **Found during:** Task 1 verification
- **Issue:** Plan's Task 1 verification script imported SufficiencyResult, Question, and QuestionSession — types that don't exist until Task 2 extends the model file
- **Fix:** Modified Task 1 verification to only test Task 1 types (OutlineLayoutType, SlideEntry, ContentOutline). Full 6-type import tested in Task 2 verification and plan-level verification
- **Files modified:** (no code changes — verification script adjustment only)
- **Verification:** Task 1 tests pass with 3 types; Task 2 tests pass with all 6 types; plan-level verification confirms all imports work
- **Committed in:** Not reflected in code commits (verification script in plan, not committed code)

---

**Total deviations:** 1 auto-fixed (1 bug in plan verification script)
**Impact on plan:** All plan code delivered correctly. Verification script fixed to match actual task scope. No impact on output quality.

## Issues Encountered
None — implementation followed the plan and spec_model.py patterns exactly.

## User Setup Required
None — no external service configuration required.

## Next Phase Readiness
- Data model foundation complete — ready for Plan 03-02 (content parsing and sufficiency logic)
- All 6 types have defaults — constructable with partial data for incremental population
- ContentOutline.validate() provides 5-rule validation with descriptive messages
- ContentOutline.to_dict()/from_dict() provide full YAML serialization round-trip for Phase 4 consumption
- QuestionSession tracks 8-question budget with gap-per-section tracking for adaptive questioning flow

---
*Phase: 03-content-gathering*
*Completed: 2026-05-06*
