---
phase: 01-pipeline-foundation
plan: 01
subsystem: pipeline
tags: [svg, drawingml, pptx, python, fork, converter, finalize]

# Dependency graph
requires:
  - phase: none
    provides: none (first plan in phase)
provides:
  - Forked 17-module SVG→DrawingML converter package under ppt_skill.converter
  - Forked 8-module SVG post-processing package under ppt_skill.finalize
  - Fixed all cross-package import chains (use_expander→embed_icons, tspan_flattener→flatten_tspan)
  - Template path resolution corrected for new 4-level package structure
affects: [02-quality-checker, 03-integration-test, 04-ppt-generation]

# Tech tracking
tech-stack:
  added: []
  patterns:
    - Package-relative imports: ppt_skill.converter.* and ppt_skill.finalize.*
    - Template resolution: 4-parent traversal from converter/ or finalize/ to project root
    - Lazy cross-package imports via _import_embed_icons() pattern
    - Soft optional dependency imports preserved (pptx_animations, cairosvg)

key-files:
  created:
    - src/ppt_skill/converter/__init__.py - Package marker with template path documentation
    - src/ppt_skill/converter/cli.py - CLI entry point (renamed pptx_cli.py)
    - src/ppt_skill/converter/builder.py - PPTX assembly (renamed pptx_builder.py)
    - src/ppt_skill/converter/converter.py - Core dispatcher (renamed drawingml_converter.py)
    - src/ppt_skill/converter/context.py - ConvertContext traversal (renamed drawingml_context.py)
    - src/ppt_skill/converter/elements.py - Element converters (renamed drawingml_elements.py)
    - src/ppt_skill/converter/paths.py - Path parsing (renamed drawingml_paths.py)
    - src/ppt_skill/converter/styles.py - Fill/stroke/effect (renamed drawingml_styles.py)
    - src/ppt_skill/converter/utils.py - Coords, colors, fonts (renamed drawingml_utils.py)
    - src/ppt_skill/converter/dimensions.py - Slide dimensions (renamed pptx_dimensions.py)
    - src/ppt_skill/converter/discovery.py - SVG file discovery (renamed pptx_discovery.py)
    - src/ppt_skill/converter/media.py - PNG rendering (renamed pptx_media.py)
    - src/ppt_skill/converter/narration.py - Narration audio (renamed pptx_narration.py)
    - src/ppt_skill/converter/notes.py - Speaker notes (renamed pptx_notes.py)
    - src/ppt_skill/converter/slide_xml.py - Legacy mode (renamed pptx_slide_xml.py)
    - src/ppt_skill/converter/use_expander.py - Icon expansion (name unchanged)
    - src/ppt_skill/converter/tspan_flattener.py - Text normalization (name unchanged)
    - src/ppt_skill/finalize/embed_icons.py - In-memory icon resolution
    - src/ppt_skill/finalize/flatten_tspan.py - Tspan flattening engine
    - src/ppt_skill/finalize/align_embed_images.py - Image alignment
    - src/ppt_skill/finalize/crop_images.py - Image cropping
    - src/ppt_skill/finalize/embed_images.py - Base64 image embedding
    - src/ppt_skill/finalize/fix_image_aspect.py - Aspect ratio correction
    - src/ppt_skill/finalize/svg_rect_to_path.py - Rounded rect conversion
  modified:
    - .gitignore - Added Python artifact patterns

key-decisions:
  - "Template resolution uses 4-parent traversal (Path(__file__).parent.parent.parent.parent) to reach project root from deep package structure"
  - "CANVAS_FORMATS fallback kept directly in dimensions.py until Plan 02 creates ppt_skill.config"
  - "Soft optional imports (pptx_animations, cairosvg) preserved as try/except — they fail gracefully without breaking core pipeline"
  - "sys.path.insert hacks removed from use_expander.py and tspan_flattener.py — cross-package imports use proper ppt_skill.finalize.* paths"

patterns-established:
  - "Lazy cross-package imports: _import_embed_icons() pattern in use_expander.py prevents hard dependency on finalize at import time"
  - "Path resolution pattern: 4-level parent traversal from any converter/ or finalize/ module to reach project-root templates/"
  - "Soft import pattern: try/except ImportError for optional dependencies (pptx_animations, cairosvg) with graceful fallback"

requirements-completed:
  - PIP-01
  - PIP-03

# Metrics
duration: 15 min
completed: 2026-05-06
---

# Phase 1 Plan 1: Pipeline Foundation — Fork & Import Fix Summary

**Forked ppt-master's 17-module SVG→DrawingML converter and 8-module post-processing pipeline into ppt-skill with all imports, paths, and cross-package references fixed for standalone operation.**

## Performance

- **Duration:** 15 min
- **Started:** 2026-05-06T14:35:23Z
- **Completed:** 2026-05-06T14:50:45Z
- **Tasks:** 3
- **Files modified:** 26 (25 created, 1 added)

## Accomplishments

- Forked 17 converter modules from ppt-master svg_to_pptx/ into src/ppt_skill/converter/ with systematic renaming (drawingml_* → short names, pptx_* → short names)
- Forked 8 finalize post-processing modules from svg_finalize/ into src/ppt_skill/finalize/ with original clean names
- Fixed all 25 internal imports across 11 converter files — zero `drawingml_` or `pptx_` (old module) references remain
- Fixed cross-package import chains: use_expander.py → ppt_skill.finalize.embed_icons and tspan_flattener.py → ppt_skill.finalize.flatten_tspan
- Corrected template path resolution from 3-parent to 4-parent traversal for the new deeper package structure
- Removed sys.path.insert hacks for svg_finalize imports; replaced with proper package-relative imports
- Stripped config.py import hack from dimensions.py, retaining CANVAS_FORMATS as fallback until Plan 02

## Task Commits

1. **Task 1: Fork 17 converter modules** - `81eda09` (feat)
2. **Task 2: Fork 8 finalize modules** - `a640199` (feat)
3. **Task 3: Fix imports, paths, cross-package refs** - `ae144c0` (feat)

## Files Created/Modified

### Created (25 files)
- `src/ppt_skill/converter/__init__.py` — Package marker with template path resolution docs
- `src/ppt_skill/converter/cli.py` — CLI entry point (from pptx_cli.py)
- `src/ppt_skill/converter/builder.py` — PPTX assembly with native shapes (from pptx_builder.py)
- `src/ppt_skill/converter/converter.py` — Core SVG→DrawingML dispatcher (from drawingml_converter.py)
- `src/ppt_skill/converter/context.py` — ConvertContext traversal state (from drawingml_context.py)
- `src/ppt_skill/converter/elements.py` — Element converters: rect, circle, path, text, image (from drawingml_elements.py)
- `src/ppt_skill/converter/paths.py` — SVG path command parsing (from drawingml_paths.py)
- `src/ppt_skill/converter/styles.py` — Fill, stroke, gradient, shadow (from drawingml_styles.py)
- `src/ppt_skill/converter/utils.py` — Coordinates, colors, fonts, matrices (from drawingml_utils.py)
- `src/ppt_skill/converter/dimensions.py` — Slide dimensions, canvas formats (from pptx_dimensions.py)
- `src/ppt_skill/converter/discovery.py` — SVG file finder (from pptx_discovery.py)
- `src/ppt_skill/converter/media.py` — SVG→PNG fallback rendering (from pptx_media.py)
- `src/ppt_skill/converter/narration.py` — Audio narration helpers (from pptx_narration.py)
- `src/ppt_skill/converter/notes.py` — Speaker notes processing (from pptx_notes.py)
- `src/ppt_skill/converter/slide_xml.py` — Legacy compatibility mode (from pptx_slide_xml.py)
- `src/ppt_skill/converter/use_expander.py` — In-memory icon expansion (name unchanged)
- `src/ppt_skill/converter/tspan_flattener.py` — In-memory text normalization (name unchanged)
- `src/ppt_skill/finalize/embed_icons.py` — Icon library resolution engine
- `src/ppt_skill/finalize/flatten_tspan.py` — Positional tspan→independent text
- `src/ppt_skill/finalize/align_embed_images.py` — Image alignment + embedding
- `src/ppt_skill/finalize/crop_images.py` — Image cropping
- `src/ppt_skill/finalize/embed_images.py` — Base64 image embedding
- `src/ppt_skill/finalize/fix_image_aspect.py` — Aspect ratio correction
- `src/ppt_skill/finalize/svg_rect_to_path.py` — Rounded rect→path conversion
- `.gitignore` — Python artifact patterns

## Decisions Made

- **4-parent path resolution**: The new package structure (`src/ppt_skill/converter/`) requires 4 levels of parent traversal to reach project root, vs. the original's 3 levels. Applied to `converter.py:308` and `finalize/embed_icons.py:52`.
- **CANVAS_FORMATS fallback**: Instead of importing the full ppt-master config.py (which has 700+ lines of unrelated configuration), we keep the essential CANVAS_FORMATS dict inline in dimensions.py. Plan 02 will create a proper `ppt_skill.config` module.
- **Soft imports preserved**: Optional dependencies (pptx_animations, cairosvg) keep their try/except ImportError blocks — the core native-shapes pipeline works without them.
- **Lazy cross-package imports**: `_import_embed_icons()` pattern keeps the converter from hard-requiring the finalize package at import time.

## Deviations from Plan

### Plan Discrepancies (Non-Blocking)

**1. [Plan Error] Wrong function name in verification: `resolve_data_icon`**
- **Found during:** Task 3 verification
- **Issue:** Plan's verification block references `resolve_data_icon` from `ppt_skill.finalize.embed_icons`, but the actual function is `resolve_icon_path`.
- **Resolution:** Used correct function name `resolve_icon_path` in verification. No code change needed — function was always named `resolve_icon_path`.
- **Impact:** Minimal — verification adjusted.

**2. [Plan Error] Wrong function name: `flatten_tspans`**
- **Found during:** Task 3 verification
- **Issue:** Plan's verification block references `flatten_tspans` from `ppt_skill.finalize.flatten_tspan`, but the actual function is `flatten_text_with_tspans`.
- **Resolution:** Used correct function name in verification.
- **Impact:** Minimal.

### Auto-fixed Issues

**1. [Rule 2 - Missing Critical] Added .gitignore for Python artifacts**
- **Found during:** Task 3
- **Issue:** No .gitignore existed, causing `__pycache__/` directories to show as untracked files.
- **Fix:** Created .gitignore with Python, IDE, and environment patterns.
- **Files modified:** `.gitignore` (created)
- **Verification:** `git status` no longer shows `__pycache__/` directories
- **Committed in:** `ae144c0` (Task 3 commit)

**2. [Rule 1 - Bug] First attempt at elements.py import fix used wrong function names**
- **Found during:** Task 3
- **Issue:** Initial edit replaced import function names with non-existent functions (`_coords`, `_hex`, `_scale`, etc.).
- **Fix:** Restored from git, re-applied only module name changes (`.drawingml_context` → `.context`, etc.), preserving all function names exactly.
- **Files modified:** `src/ppt_skill/converter/elements.py`
- **Verification:** `diff` against original confirms only module names changed, not function names; imports succeed.
- **Committed in:** `ae144c0` (Task 3 commit)

---

**Total deviations:** 3 (2 plan discrepancies, 1 auto-fix)
**Impact on plan:** All auto-fixes necessary for correctness. Plan discrepancies were naming errors in the verification block — implementation was correct.

## Issues Encountered

- `ModuleNotFoundError: No module named 'ppt_skill'` — resolved by adding `src/` to PYTHONPATH for verification. This will be addressed in Plan 02 with proper package setup (`__init__.py`, requirements.txt).
- First elements.py edit replaced function names instead of just module names — caught by ImportError during verification, fixed by restoring and re-editing with correct approach.

## Next Phase Readiness

- All 25 converter/finalize modules import successfully under the new package structure
- Cross-package dependencies (use_expander→finalize, tspan_flattener→finalize) resolve correctly
- Template path resolution corrected for 4-level traversal
- Ready for Plan 02: Quality checker forking and templates setup

---
*Phase: 01-pipeline-foundation*
*Completed: 2026-05-06*
