# Project State

## Project Reference

See: .planning/PROJECT.md (updated 2026-05-06)

**Core value:** Generate professional, editable PPTs that faithfully follow extracted design specifications — turning raw content into polished presentations without requiring design expertise.
**Current focus:** Phase 1 - Pipeline Foundation

## Current Position

Phase: 1 of 5 (Pipeline Foundation)
Plan: 2 of 3 (01-01 converter fork complete, 01-02 quality + templates complete)
Status: In Progress
Last activity: 2026-05-06 — Plan 01-02 executed: quality checker + 11,631 icons + 70 chart templates

Progress: [██████░░░░░░░░░░░░░░] 30%

## Performance Metrics

**Velocity:**
- Total plans completed: 2
- Average duration: ~12 min
- Total execution time: ~25 min

**By Phase:**

| Phase | Plans | Duration | Avg/Plan |
|-------|-------|----------|----------|
| 01-pipeline-foundation | 2 | ~25 min | ~12 min |

**Recent Trend:**
- Plan 01-01: Forked 17 converter + 8 finalize modules (~10 min)
- Plan 01-02: Quality checker + templates (~15 min)

*Updated after each plan completion*

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

### Pending Todos

None yet.

### Blockers/Concerns

- **Phase 1 risk:** ppt-master module entanglement — must map full dependency graph before stripping anything. Dual-consumer modules (`embed_icons.py`, `flatten_tspan.py`) can silently break if removed.
- **Phase 1 risk:** SVG banned-feature blacklist must be adopted wholesale from ppt-master — AI-generated SVGs routinely use masks, `rgba()`, `@font-face` that break DrawingML conversion.
- **Phase 2 risk:** Theme color resolution must produce concrete HEX values, not theme-aware OOXML references (`accent1` → `#4472C4`).
- **Phase 4 risk:** spec_lock anti-drift mechanism — must prevent color/font drift across long decks during generation.

## Session Continuity

Last session: 2026-05-06
Stopped at: Completed 01-02-PLAN.md — quality checker + templates. Ready for Plan 01-03 (integration + tests).
Resume file: .planning/phases/01-pipeline-foundation/01-02-SUMMARY.md
