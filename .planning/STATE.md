# Project State

## Project Reference

See: .planning/PROJECT.md (updated 2026-05-06)

**Core value:** Generate professional, editable PPTs that faithfully follow extracted design specifications — turning raw content into polished presentations without requiring design expertise.
**Current focus:** Phase 1 - Pipeline Foundation

## Current Position

Phase: 1 of 5 (Pipeline Foundation)
Plan: 0 of 0 (not yet planned)
Status: Ready to plan
Last activity: 2026-05-06 — Roadmap created; 19 requirements mapped to 5 phases

Progress: [░░░░░░░░░░░░░░░░░░░░] 0%

## Performance Metrics

**Velocity:**
- Total plans completed: 0
- Average duration: N/A
- Total execution time: 0 hours

**By Phase:**

| Phase | Plans | Total | Avg/Plan |
|-------|-------|-------|----------|
| - | - | - | - |

**Recent Trend:**
- No plans executed yet.

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

### Pending Todos

None yet.

### Blockers/Concerns

- **Phase 1 risk:** ppt-master module entanglement — must map full dependency graph before stripping anything. Dual-consumer modules (`embed_icons.py`, `flatten_tspan.py`) can silently break if removed.
- **Phase 1 risk:** SVG banned-feature blacklist must be adopted wholesale from ppt-master — AI-generated SVGs routinely use masks, `rgba()`, `@font-face` that break DrawingML conversion.
- **Phase 2 risk:** Theme color resolution must produce concrete HEX values, not theme-aware OOXML references (`accent1` → `#4472C4`).
- **Phase 4 risk:** spec_lock anti-drift mechanism — must prevent color/font drift across long decks during generation.

## Session Continuity

Last session: 2026-05-06
Stopped at: Roadmap created and validated. 5 phases defined covering all 19 v1 requirements.
Resume file: None
