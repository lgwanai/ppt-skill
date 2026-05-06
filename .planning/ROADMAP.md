# Roadmap: PPT Spec Skill

## Overview

Five-phase journey building an AI-powered PPT generation skill. Start by forking and validating ppt-master's battle-tested SVG→DrawingML pipeline (the execution engine, highest technical risk). Then build the core differentiator — spec extraction from existing PPTX files. Add the interactive content gathering UX, integrate everything into spec-driven PPTX generation, and finally package as a dual-platform installable CLI skill.

## Phases

- [x] **Phase 1: Pipeline Foundation** - Fork and adapt ppt-master's SVG→DrawingML pipeline; validate standalone conversion
- [x] **Phase 2: Spec Extraction** - Extract design specifications from existing PPTX files; save as reusable spec artifacts (completed 2026-05-06)
- [ ] **Phase 3: Content Gathering** - Smart hybrid questioning and slide-by-slide outline generation
- [ ] **Phase 4: Spec-Driven PPT Generation** - Generate editable PPTX files that faithfully replicate extracted specs
- [ ] **Phase 5: Platform Packaging** - Package as dual-platform installable CLI skill with end-to-end verification

## Phase Details

### Phase 1: Pipeline Foundation
**Goal**: Forked SVG→DrawingML pipeline converts SVG files to natively editable PPTX shapes — the proven execution engine running standalone
**Depends on**: Nothing (first phase)
**Requirements**: PIP-01, PIP-02, PIP-03, PIP-04
**Success Criteria** (what must be TRUE):
  1. A standalone Python script converts a sample SVG to a .pptx file where every shape is a native DrawingML element (not a flattened image)
  2. The SVG quality checker correctly rejects SVGs containing banned features — masks, `rgba()` colors, `@font-face`, HTML entities, and `<style>` tags
  3. The output PPTX opens in Microsoft PowerPoint with all shapes remaining individually selectable, movable, and editable
   4. The forked pipeline passes a regression test suite against a known-good set of SVG→PPTX pairs covering the full converter module chain
**Plans**: 3 plans in 2 waves

Plans:
- [x] 01-01-PLAN.md — Fork and wire 17-module svg_to_pptx/ converter + 8-module svg_finalize/ pipeline (imports, paths, cross-refs)
- [x] 01-02-PLAN.md — Fork SVG quality checker + create minimal config + copy 11,631 icons and 73 chart templates
- [x] 01-03-PLAN.md — Create unified pipeline CLI + regression test suite (quality checker + converter integration + e2e)

### Phase 2: Spec Extraction
**Goal**: Users provide a reference PPTX and receive a structured, reusable design specification file capturing colors, fonts, layout patterns, and slide type classifications
**Depends on**: Phase 1
**Requirements**: SPC-01, SPC-02, SPC-03, SPC-04, SPC-05
**Success Criteria** (what must be TRUE):
  1. User provides a reference .pptx file and the tool outputs a structured spec file capturing color palette (HEX values, not theme references), font families/sizes/hierarchy, and spatial layout patterns
  2. The extracted spec classifies slides into distinct types (title, content, section divider, image+text, data) with each type's unique visual properties documented
  3. Spec captures presentation logic — slide sequencing patterns, content density rhythm, and storytelling structure — as structured metadata
  4. Spec files persist in a project-local `specs/` directory, survive across sessions, and are format-compatible with git diff/versioning
  5. User can list all available specs in the project and select one as the active design target for generation
**Plans**: 3 plans in 2 waves

Plans:
- [x] 02-01-PLAN.md — Spec data model (dataclass schemas) + theme extraction (colors, fonts, backgrounds via lxml) + config.py placeholders
- [x] 02-02-PLAN.md — Slide classification (5 types, dual-strategy) + spatial layout analysis (margins, positioning) + content density/rhythm analysis (percentile-based)
- [x] 02-03-PLAN.md — SpecExtractor orchestrator + CLI commands (extract-spec, list-specs, select-spec) + YAML serialization + integration tests

### Phase 3: Content Gathering
**Goal**: Tool intelligently gathers presentation content through adaptive questioning, producing a user-approved slide-by-slide outline before any generation begins
**Depends on**: Phase 2
**Requirements**: GEN-01, GEN-02, GEN-03
**Success Criteria** (what must be TRUE):
  1. When user provides a topic with insufficient detail, the tool asks section-level overview questions first to establish high-level structure before diving into specifics
  2. The tool asks targeted follow-up questions to fill content gaps (capped at 8 total questions) without causing decision fatigue
  3. User receives a detailed slide-by-slide content outline (title, body content, and suggested layout type per slide) for review and approval
  4. User can bypass all questioning when initial input is sufficiently detailed — the tool proceeds directly to outline generation
  5. The outline includes layout type recommendations per slide drawn from available template types (title, content, two-column, section header, image+text)
**Plans**: 3 plans in 3 waves

Plans:
- [ ] 03-01-PLAN.md — Content data model (ContentOutline, SlideEntry, SufficiencyResult, Question, QuestionSession dataclasses + OutlineLayoutType enum)
- [ ] 03-02-PLAN.md — Sufficiency assessment (rubric-based scoring) + adaptive questioning module (section-first overview, gap-fill follow-up, 8-question cap)
- [ ] 03-03-PLAN.md — ContentGatherer orchestrator (3-phase pipeline) + CLI commands (gather-content, generate-outline) + integration tests

### Phase 4: Spec-Driven PPT Generation
**Goal**: Generate natively editable PPTX files that faithfully replicate the selected spec's visual identity — colors, fonts, layouts, and content rhythm — from an approved outline
**Depends on**: Phase 1, Phase 2
**Requirements**: PPT-01, PPT-02, PPT-03, PPT-04
**Success Criteria** (what must be TRUE):
  1. Tool generates a .pptx file from an approved content outline combined with a selected design spec — all text, shapes, and layouts are natively editable in PowerPoint
  2. Generated PPTX strictly follows the spec's color palette, font hierarchy (families, sizes, weights), layout patterns, and content density rhythm
  3. No slide uses flattened images for elements that should be shapes — every text box, rectangle, line, and icon is a native DrawingML element
  4. SVG generation produces output compatible with ppt-master's banned-feature list — HEX colors only, no masks, no rgba(), no @font-face, no HTML entities
  5. Baseline templates exist for 5+ common slide types (title, content, two-column, section divider, image+text) with spec-applied styling applied automatically
**Plans**: TBD

### Phase 5: Platform Packaging
**Goal**: Skill is installable and fully functional as a CLI skill in both Opencode and Claude Code environments, with end-to-end workflow verified
**Depends on**: Phase 1, Phase 2, Phase 3, Phase 4
**Requirements**: PLT-01, PLT-02, PLT-03
**Success Criteria** (what must be TRUE):
  1. The skill installs via `~/.opencode/skills/` directory convention and runs correctly in an Opencode environment
  2. The skill installs via `~/.claude/skills/` directory convention and runs correctly in a Claude Code environment
  3. Core Python generation logic is platform-agnostic — platform-specific SKILL.md wrappers are thin adapters with no duplicated business logic
  4. End-to-end workflow completes successfully: reference PPTX input → spec extracted → content gathered and outlined → new PPTX generated matching the spec
  5. All dependencies install cleanly via `requirements.txt` with pinned versions, compatible with Python 3.10+ across macOS, Linux, and Windows
**Plans**: TBD

## Progress

**Execution Order:** Phases execute in numeric order.

| Phase | Plans Complete | Status | Completed |
|-------|----------------|--------|-----------|
| 1. Pipeline Foundation | 3/3 | Complete | 2026-05-06 |
| 2. Spec Extraction | 3/3 | Complete   | 2026-05-06 |
| 3. Content Gathering | 0/3 | Not started | - |
| 4. Spec-Driven PPT Generation | 0/— | Not started | - |
| 5. Platform Packaging | 0/— | Not started | - |
