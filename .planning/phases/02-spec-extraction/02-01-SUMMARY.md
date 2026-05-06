---
phase: 02-spec-extraction
plan: 01
subsystem: spec
tags: [dataclass, lxml, python-pptx, theme1.xml, color-palette, ooxml, yaml-schema]

# Dependency graph
requires:
  - phase: 01-pipeline-foundation
    provides: "SVG→DrawingML pipeline, CANVAS_FORMATS config, package structure"
provides:
  - "DesignSpec dataclass schema (9 types) — Phase 2→4 contract for YAML serialization"
  - "Theme color extraction — 12-color HEX palette from theme1.xml via lxml"
  - "Theme font extraction — majorFont/minorFont typefaces from fontScheme"
  - "Slide background extraction — 4-level inheritance chain with bug #1126 workaround"
  - "DESIGN_COLORS, FONT_SIZES, LAYOUT_MARGINS config placeholders for Phase 4"
affects: [02-02-layout-classification, 02-03-spec-orchestrator, 04-spec-driven-generation]

# Tech tracking
tech-stack:
  added: []
  patterns:
    - "Hybrid python-pptx + lxml extraction (python-pptx for structure, lxml for theme details)"
    - "Two-pass color resolution (srgbClr/sysClr → schemeClr references)"
    - "4-level background inheritance chain (slide → layout → master → theme)"
    - "Dataclass schema-first design (contract before implementation)"
    - "Field defaults for incremental population (DesignSpec() constructible at any stage)"

key-files:
  created:
    - src/ppt_skill/spec/__init__.py
    - src/ppt_skill/spec/spec_model.py
    - src/ppt_skill/spec/theme.py
  modified:
    - src/ppt_skill/config.py
    - src/ppt_skill/__init__.py

key-decisions:
  - "Used dataclasses (NOT Pydantic) — minimal dependencies, sufficient for Phase 2–4 contract"
  - "Two-pass color resolution for schemeClr references against clrScheme"
  - "4-level background inheritance walk (slide→layout→master→theme) as workaround for python-pptx bug #1126"
  - "12-color palette naming follows OOXML scheme (dk1/lt1/accent1) with from_theme_scheme() for ColorPalette mapping"

patterns-established:
  - "Schema-first design: dataclass models defined before any extraction code"
  - "Hybrid extraction: python-pptx for slide hierarchy, lxml for theme/XML details"
  - "Incremental construction: all dataclass fields have defaults for partial population"
  - "Config placeholders: DESIGN_COLORS/FONT_SIZES/LAYOUT_MARGINS seeded with sensible Office defaults"

requirements-completed: [SPC-01]

# Metrics
duration: 7 min
completed: 2026-05-06
---

# Phase 2 Plan 1: Spec Data Model & Theme Extraction Summary

**Dataclass-based design spec schema (9 types) with lxml-driven theme color/font/background extraction from PPTX theme1.xml, plus config placeholders for Phase 4 generation**

## Performance

- **Duration:** 7 min
- **Started:** 2026-05-06T16:23:48Z
- **Completed:** 2026-05-06T16:30:51Z
- **Tasks:** 3
- **Files modified:** 5

## Accomplishments
- Defined 9 dataclass/enum types (SlideType, DensityLabel, ColorPalette, Typography, LayoutMargins, SlideLayoutSpec, SlideSpec, PresentationRhythm, DesignSpec) — all with field defaults for incremental population
- Implemented theme color extraction from theme1.xml: two-pass resolution handling srgbClr, sysClr (lastClr), and schemeClr references
- Implemented theme font extraction: majorFont/minorFont latin typefaces from fontScheme
- Implemented slide background extraction: 4-level inheritance chain (slide→layout→master→theme) with python-pptx bug #1126 workaround, including gradient detection
- Extended config.py with DESIGN_COLORS (12 keys), FONT_SIZES (7 tiers), LAYOUT_MARGINS (8 keys) — all with sensible Office defaults

## Task Commits

Each task was committed atomically:

1. **Task 1: Create spec data model** - `0a09049` (feat)
2. **Task 2: Implement theme extraction** - `762f0bd` (feat)
3. **Task 3: Extend config.py with placeholders** - `40ec682` (feat)

**Plan metadata:** (to be committed after SUMMARY creation)

## Files Created/Modified
- `src/ppt_skill/spec/__init__.py` - Package init exporting all 9 model types
- `src/ppt_skill/spec/spec_model.py` - 9 dataclass/enum types (201 lines), the Phase 2→4 contract
- `src/ppt_skill/spec/theme.py` - Color/font/background extraction via lxml (398 lines)
- `src/ppt_skill/config.py` - Extended with DESIGN_COLORS, FONT_SIZES, LAYOUT_MARGINS
- `src/ppt_skill/__init__.py` - Updated exports for new config constants

## Decisions Made
- Used dataclasses (NOT Pydantic) — minimal dependencies, sufficient for Phase 2–4 contract
- Two-pass color resolution for schemeClr references against clrScheme
- 4-level background inheritance walk (slide→layout→master→theme) as workaround for python-pptx bug #1126
- 12-color palette naming follows OOXML scheme (dk1/lt1/accent1) with `from_theme_scheme()` classmethod for semantic ColorPalette mapping
- Config placeholders seeded with Office 365 defaults — provides safe fallback for Phase 4

## Deviations from Plan

None - plan executed exactly as written.

## Issues Encountered

None.

## User Setup Required

None — PyYAML (>=6.0) was already installed in the environment. No additional configuration needed.

## Next Phase Readiness
- `spec_model.py` is complete — ready for Phase 3 and 4 modules to import
- `theme.py` is complete — ready for `SpecExtractor` orchestrator in Plan 02-03
- `config.py` placeholders are in place — Phase 4 has safe fallback values
- Ready for Plan 02-02: Slide classification, spatial analysis, and density/rhythm analysis

---
*Phase: 02-spec-extraction*
*Completed: 2026-05-06*
