# Project State

## Project Reference

See: .planning/PROJECT.md (updated 2026-05-06)

**Core value:** Generate professional, editable PPTs that faithfully follow extracted design specifications — turning raw content into polished presentations without requiring design expertise.
**Current focus:** Phase 3 - Content Gathering

## Current Position

Phase: 3 of 5 (Content Gathering) — **IN PROGRESS**
Plan: 1 of 3 (03-01 executed — Content data model: 6 dataclasses + 1 enum, YAML contract with Phase 4)
Status: In Progress
Last activity: 2026-05-06 — Plan 03-01 executed: content data model (~7 min, 2 tasks, 2 files)

Progress: [████████░░░░░░░░░░░░] 33%

## Performance Metrics

**Velocity:**
- Total plans completed: 7
- Average duration: ~8 min
- Total execution time: ~56 min

**By Phase:**

| Phase | Plans | Duration | Avg/Plan |
|-------|-------|----------|----------|
| 01-pipeline-foundation | 3 | ~30 min | ~10 min |
| 02-spec-extraction | 3 | ~19 min | ~6.3 min |
| 03-content-gathering | 1 | ~7 min | ~7 min |

**Recent Trend:**
- Plan 01-01: Forked 17 converter + 8 finalize modules (~10 min)
- Plan 01-02: Quality checker + templates (~15 min)
- Plan 01-03: Pipeline CLI + 23 regression tests (~5 min)
- Plan 02-01: Spec data model + theme extraction + config placeholders (~7 min)
- Plan 02-02: Slide classifier + layout/font/density analyzers (~6 min)
- Plan 02-03: SpecExtractor orchestrator + CLI commands + 7 integration tests (~6 min)
- Plan 03-01: Content data model (OutlineLayoutType enum + 6 dataclasses) (~7 min)

*Updated after each plan completion*
| Phase 01-pipeline-foundation P01-01 | 15 min | 3 tasks | 26 files |
| Phase 01-pipeline-foundation P03 | 5 min | 3 tasks | 16 files |
| Phase 02-spec-extraction P01 | 7 min | 3 tasks | 5 files |
| Phase 02-spec-extraction P02 | 6 min | 4 tasks | 4 files |
| Phase 02-spec-extraction P03 | 6 min | 3 tasks | 6 files |
| Phase 03-content-gathering P01 | 7 min | 2 tasks | 2 files |

## Accumulated Context

### Decisions

Decisions are logged in PROJECT.md Key Decisions table.
Recent decisions affecting current work:

- Fork ppt-master v2.6.0 for SVG→DrawingML pipeline (reuse proven 17-module converter rather than rebuild)
- Spec extraction via python-pptx + raw lxml/XML (python-pptx's Slide.background API is broken — must use lxml for backgrounds)
- Spec files are project-local YAML artifacts (portable, git-diffable, no infrastructure needed)
- Smart hybrid questioning with 8-question cap (balances completeness vs. decision fatigue)
- Phase 1 (Pipeline) comes first — validates highest technical risk before building anything dependent
- **01-02:** Stripped template_mode, spec_lock drift, and sourced image checks from quality checker (Phase 1 only validates individual SVGs; template/spec concerns belong to Phases 2-4)
- **01-02:** Minimal config.py — only CANVAS_FORMATS (8 entries in EMU); DESIGN_COLORS/INDUSTRY_COLORS added in Phase 2
- **01-02:** Template assets at project root (templates/) not inside src/ package — mirrors ppt-master's runtime path resolution
- [Phase 01-pipeline-foundation]: Template resolution uses 4-parent traversal from deep package structure to reach project root — New structure src/ppt_skill/converter/ requires 4 levels of parent traversal vs original 3 levels
- [Phase 01-pipeline-foundation]: CANVAS_FORMATS fallback kept inline in dimensions.py until Plan 02 creates ppt_skill.config — Avoid pulling in 700+ lines of ppt-master configuration
- [Phase 01-pipeline-foundation]: Soft optional imports preserved for pptx_animations and cairosvg — Core native-shapes pipeline works without them; they fail gracefully
- [Phase 01-pipeline-foundation]: Cross-package imports use proper ppt_skill.finalize.* paths, sys.path.insert hacks removed — Clean package structure enables proper Python package imports
- **01-03:** Unified pipeline as single function (convert_svg_to_pptx) with quality gate and optional --skip-check flag
- **01-03:** Pipeline defaults: use_native_shapes=True, use_compat_mode=False — Phase 1 targets native shapes exclusively
- **01-03:** pyproject.toml created for pip install -e . — required for from ppt_skill.pipeline imports in tests/CLI
- [Phase 02-spec-extraction]: Used dataclasses (NOT Pydantic) — minimal dependencies, sufficient for Phase 2–4 contract
- [Phase 02-spec-extraction]: Two-pass color resolution for schemeClr references against clrScheme
- [Phase 02-spec-extraction]: 4-level background inheritance walk (slide→layout→master→theme) as workaround for python-pptx bug #1126
- [Phase 02-spec-extraction]: Config placeholders seeded with Office 365 defaults — provides safe fallback for Phase 4
- [Phase 02-spec-extraction]: Used content-based title vs body heuristic for font classification (placeholder type + position + size)
- [Phase 02-spec-extraction]: CLI functions are stateless and callable programmatically — designed for Phase 3-4 code reuse without argparse dependency
- [Phase 02-spec-extraction]: YAML serialization via dataclasses.asdict() + Enum .value walk — no additional serialization library beyond PyYAML
- [Phase 03-content-gathering]: Used dataclasses (NOT Pydantic) for content model — matching Phase 2 convention, zero new dependencies
- [Phase 03-content-gathering]: OutlineLayoutType has 6 values — extends Phase 2's SlideType (5 values) with TWO_COLUMN for flexible layouts
- [Phase 03-content-gathering]: QuestionSession budget starts at 8 with mark_asked() decrement — enforces max questioning limit
- [Phase 03-content-gathering]: _dataclass_to_dict reused from ppt_skill.spec.extractor — same serialization pattern for enums and nested dataclasses
- [Phase 03-content-gathering]: Sufficiency threshold documented but NOT hardcoded in dataclass — enforced in sufficiency.py module

### Pending Todos

None yet.

### Blockers/Concerns

- **Phase 1 risk:** ppt-master module entanglement — must map full dependency graph before stripping anything. Dual-consumer modules (`embed_icons.py`, `flatten_tspan.py`) can silently break if removed.
- **Phase 1 risk:** SVG banned-feature blacklist must be adopted wholesale from ppt-master — AI-generated SVGs routinely use masks, `rgba()`, `@font-face` that break DrawingML conversion.
- **Phase 2 risk:** Theme color resolution must produce concrete HEX values, not theme-aware OOXML references (`accent1` → `#4472C4`).
- **Phase 4 risk:** spec_lock anti-drift mechanism — must prevent color/font drift across long decks during generation.

## Session Continuity

Last session: 2026-05-06
Stopped at: Completed 03-01-PLAN.md — Content data model: OutlineLayoutType enum + 6 dataclasses. Phase 3 IN PROGRESS (1/3 plans).
Resume file: .planning/phases/03-content-gathering/03-01-SUMMARY.md
