---
phase: 03-content-gathering
plan: 03
subsystem: content
tags: [content, outline, orchestration, cli, testing]
requires:
  - phase: 03-content-gathering
    plan: 01
    provides: ContentOutline, SlideEntry, OutlineLayoutType, SufficiencyResult, Question, QuestionSession dataclasses
  - phase: 03-content-gathering
    plan: 02
    provides: assess_sufficiency, build_sufficiency_result, questioning (generate_section_questions, generate_gap_questions, identify_content_gaps)
provides:
  - ContentGatherer orchestrator with 3-phase pipeline (assess → question → generate)
  - CLI content commands (gather_content, generate_outline_from_summary, list_outlines)
  - Integration tests (33 tests covering serialization, validation, sufficiency, budget, pipeline, YAML persistence, layout heuristics)
affects: [04-ppt-generation, 05-packaging]
tech-stack:
  added: []
  patterns:
    - "ContentGatherer follows Phase 2 SpecExtractor pattern: init → pipeline → save"
    - "CLI functions are stateless and callable from Python without argparse (matching spec_commands.py pattern)"
    - "LLM prompt templates are module-level constants (tuning surface)"
    - "YAML serialization reuses _dataclass_to_dict from ppt_skill.spec.extractor"
key-files:
  created:
    - src/ppt_skill/content/gatherer.py (ContentGatherer orchestrator, 848 lines)
    - src/ppt_skill/cli/content_commands.py (CLI content commands, 253 lines)
    - tests/test_content_gathering.py (33 integration tests, 655 lines)
  modified: []
key-decisions:
  - "ContentGatherer is a prompt-orchestrating state machine — LLM runtime processes prompts; Python code provides state tracking, validation, serialization"
  - "Layout type heuristics: comparison keywords → TWO_COLUMN, image keywords → IMAGE_TEXT, data/statistics keywords → DATA, default → CONTENT"
  - "Sections identified via regex (numbered: 1. 2. 3.) with fallback to short capitalized lines, then generic Section N labels"
  - "skip_questions mode fully bypasses Phase 2 (GEN-03 satisfied) — section identification still runs for structure extraction"
patterns-established:
  - "ContentGatherer.gather(mode) as pipeline entry — auto for full pipeline, skip_questions for programmatic content"
  - "ContentGatherer.save() writes YAML with Phase 2 serialization pattern (default_flow_style=False, sort_keys=False, allow_unicode=True)"
  - "Integration tests use mode=skip_questions to avoid LLM dependency in test suite"
  - "CLI commands auto-resolve spec_name via get_active_spec() when not explicitly provided"
requirements-completed: [GEN-02, GEN-03]
duration: 10 min
completed: 2026-05-07
---

# Phase 3 Plan 3: ContentGatherer Orchestrator + CLI + Integration Tests Summary

**ContentGatherer orchestrator tying the full 3-phase content pipeline together, CLI entry points for user interaction, and 33 integration tests validating end-to-end behavior**

## Performance

- **Duration:** 10 min
- **Started:** 2026-05-07T04:38:13Z
- **Completed:** 2026-05-07T04:48:40Z
- **Tasks:** 3
- **Files modified:** 3 (all new)

## Accomplishments

- ContentGatherer orchestrator with full 3-phase pipeline: sufficiency assessment → adaptive questioning → outline generation
- CLI commands for interactive and programmatic use: gather_content(), generate_outline_from_summary(), list_outlines()
- 33 integration tests validating serialization, validation, sufficiency, question budget, pipeline execution, YAML persistence, and layout heuristics

## Task Commits

1. **Task 1: ContentGatherer orchestrator with 3-phase pipeline** - `e8fffb9` (feat: create ContentGatherer orchestrator)
2. **Task 2: CLI content commands (gather-content, generate-outline, list-outlines)** - `c965ca4` (feat: create CLI content commands)
3. **Task 3: Integration tests for full content gathering pipeline** - `c64ae3c` (feat: add integration tests)

## Files Created

- `src/ppt_skill/content/gatherer.py` - ContentGatherer orchestrator: gather(), save(), load_outline(), prompt templates, layout type heuristics
- `src/ppt_skill/cli/content_commands.py` - CLI functions: gather_content(), generate_outline_from_summary(), list_outlines()
- `tests/test_content_gathering.py` - 33 integration tests: model serialization, validation, sufficiency, budget, pipeline, YAML persistence, layout heuristics

## Decisions Made

- Layout type heuristics use keyword matching (comparison/imagine/data keywords) — simple, deterministic, sufficient for outline generation
- Section extraction uses regex for numbered structures with fallback to short capitalized lines then generic labels — handles varied input formats
- skip_questions mode runs section identification for structure but bypasses all questioning — satisfies GEN-03
- Prompt templates are module-level constants — the primary tuning surface for output quality without code changes

## Deviations from Plan

### Auto-fixed Issues

**1. [Rule 1 - Bug] Fixed section identification in skip_questions mode**
- **Found during:** Task 3 (Integration tests)
- **Issue:** `ContentGatherer.gather()` in `skip_questions` mode created a `QuestionSession` but never set `sections_identified`, causing `_generate_outline()` to produce zero sections and a minimal outline.
- **Fix:** Added `self.session.sections_identified = self._extract_sections(user_input)` in the `skip_questions` and `self.sufficiency.sufficient` branches of `gather()`.
- **Files modified:** `src/ppt_skill/content/gatherer.py`
- **Verification:** `test_pipeline_sections_identified` now passes — 3 sections detected from detailed input.
- **Committed in:** `c64ae3c` (Task 3 commit)

**2. [Rule 1 - Bug] Fixed layout type test false positive**
- **Found during:** Task 3 (Integration tests)
- **Issue:** Test `test_default_layout_is_content` used "business metrics" which triggered the DATA layout heuristic (keyword "metric" matched). Changed test text to "business strategy and operations".
- **Files modified:** `tests/test_content_gathering.py`
- **Verification:** Test now correctly asserts CONTENT layout.
- **Committed in:** `c64ae3c` (Task 3 commit)

**3. [Rule 1 - Bug] Fixed YAML round-trip test body length**
- **Found during:** Task 3 (Integration tests)
- **Issue:** Test `test_outline_yaml_round_trip` used body entries "Point A" (7 chars) and "Point B" (7 chars) which are too short for the ContentOutline.validate() >10 char requirement. `to_dict()` calls `validate()` which rejects short body entries.
- **Fix:** Changed body entries to longer strings meeting the >10 char threshold.
- **Files modified:** `tests/test_content_gathering.py`
- **Verification:** `test_outline_yaml_round_trip` now passes.
- **Committed in:** `c64ae3c` (Task 3 commit)

---

**Total deviations:** 3 auto-fixed (all Rule 1 - Bug)
**Impact on plan:** All fixes essential for correctness. No scope creep. The section identification bug was a logical oversight in the skip_questions branch; the other two were test fixture issues.

## Issues Encountered

None — all issues resolved via deviation fixes above.

## Next Phase Readiness

- Phase 3 (Content Gathering) fully complete with all 3 plans executed
- ContentGatherer pipeline ready for Phase 4 (PPT Generation) — produces validated ContentOutline YAML files consumed by the generation pipeline
- CLI commands ready for Phase 5 (Packaging) wiring into main CLI entry point
- GEN-02 (sufficiency assessment) and GEN-03 (skip questions mode) satisfied

---
*Phase: 03-content-gathering*
*Completed: 2026-05-07*
