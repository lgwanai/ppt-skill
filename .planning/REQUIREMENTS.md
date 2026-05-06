# Requirements: PPT Spec Skill

**Defined:** 2026-05-06
**Core Value:** Generate professional, editable PPTs that faithfully follow extracted design specifications — turning raw content into polished presentations without requiring design expertise.

## v1 Requirements

Requirements for initial release. Each maps to roadmap phases.

### Spec Extraction & Management (SPC)

- [ ] **SPC-01**: User can provide an existing .pptx file and the tool extracts visual style (color palette, backgrounds, gradients), typography (font families, sizes, weights, hierarchy), and spatial layout patterns (margins, spacing, element positioning)
- [ ] **SPC-02**: Extracted spec includes layout classification — identifying and categorizing slide types (title slides, content slides, section dividers, image+text layouts, data slides) with their distinct visual properties
- [ ] **SPC-03**: Extracted spec captures presentation logic — slide sequencing patterns, content density rhythm (anchor/dense/breathing), and storytelling structure
- [ ] **SPC-04**: Spec is saved as a structured project-local file that can be versioned, shared, and reused across generation sessions
- [ ] **SPC-05**: User can list available specs and select one as the target style for new PPT generation

### Content Gathering (GEN)

- [ ] **GEN-01**: Smart hybrid questioning — when user input lacks sufficient detail, the tool asks section-level overview questions first, then gap-fills per section with targeted follow-ups, capping at 8 total questions
- [ ] **GEN-02**: Tool generates a detailed content outline (title, body content, suggested layout type per slide) for user review and approval before any PPT generation begins
- [ ] **GEN-03**: User can skip questioning entirely when input is sufficiently detailed, going directly to content outline generation

### PPT Generation (PPT)

- [ ] **PPT-01**: Tool generates natively editable .pptx files using python-pptx and a forked/adapted SVG→DrawingML pipeline — all text, shapes, and layouts remain editable in PowerPoint (no flattened images)
- [ ] **PPT-02**: Generated PPTX strictly follows the selected spec's visual rules — colors, fonts, layout patterns, and content rhythm are faithfully reproduced
- [ ] **PPT-03**: Template layout system provides baseline templates for common slide types (title, content, two-column, section divider, image+text) with spec-applied styling
- [ ] **PPT-04**: SVG generation respects ppt-master's banned-feature list — no masks, no HTML entities, HEX colors only, PPT-safe font stack, proper opacity handling

### Pipeline Foundation (PIP)

- [ ] **PIP-01**: Fork and adapt ppt-master's core SVG→DrawingML converter modules (17-module pipeline) for spec-driven generation — every shape, text box, and gradient is a native PowerPoint element
- [ ] **PIP-02**: SVG quality checker validates generated SVGs against ppt-master compatibility rules before conversion
- [ ] **PIP-03**: Post-processing pipeline (icon embedding, tspan flattening, image alignment) produces PPTX-parsable SVG output
- [ ] **PIP-04**: Inherit ppt-master's icon library (11,600+ SVG icons) and chart templates (70+ chart types) for visual variety

### Platform & Packaging (PLT)

- [ ] **PLT-01**: Skill works as an installable CLI skill in both Opencode and Claude Code environments
- [ ] **PLT-02**: Core generation logic is platform-agnostic Python — platform-specific SKILL.md wrappers are thin adapters
- [ ] **PLT-03**: Dependency management via requirements.txt with pinned versions, compatible with Python 3.10+

## v2 Requirements

Deferred to future release. Tracked but not in current roadmap.

### Enhanced Spec Features

- **SPC-06**: Presentation logic model captures argument flow, narrative structure, and persuasion patterns as structured metadata
- **SPC-07**: Spec validation and repair — detect and fix common spec extraction issues automatically

### Extended Generation

- **GEN-04**: Multi-source content input — accept PDF, Word, URL, and Markdown documents as content sources
- **PPT-05**: Chart and data visualization generation from structured data (CSV, JSON, inline tables)

### Power Features

- **PLT-04**: Batch generation — apply one spec to multiple content sources producing consistent multi-deck output
- **PLT-05**: Multi-language content support for generated slides

## Out of Scope

Explicitly excluded. Documented to prevent scope creep.

| Feature | Reason |
|---------|--------|
| Web-based visual editor (WYSIWYG) | Doubles complexity; PPTX IS the editor. Skill is terminal-native |
| Real-time collaboration | Requires server infrastructure; violates local-only constraint |
| Cloud-based spec sharing | Requires backend/auth; specs are portable files for Git/file sharing |
| TTS narration / video export | Entirely different pipeline; user records in PowerPoint |
| Multi-format output (HTML, docs, social) | Focus on one format done well; PPTX only |
| AI image generation in slides | Requires image API, adds complexity, reduces professional feel |
| Live data-connected presentations | Huge infrastructure dependency; user exports data manually |
| Auto-branding from URL | Fragile and unreliable; spec extraction from PPTX is more precise |

## Traceability

Which phases cover which requirements. Updated during roadmap creation.

| Requirement | Phase | Status |
|-------------|-------|--------|
| SPC-01 | — | Pending |
| SPC-02 | — | Pending |
| SPC-03 | — | Pending |
| SPC-04 | — | Pending |
| SPC-05 | — | Pending |
| GEN-01 | — | Pending |
| GEN-02 | — | Pending |
| GEN-03 | — | Pending |
| PPT-01 | — | Pending |
| PPT-02 | — | Pending |
| PPT-03 | — | Pending |
| PPT-04 | — | Pending |
| PIP-01 | — | Pending |
| PIP-02 | — | Pending |
| PIP-03 | — | Pending |
| PIP-04 | — | Pending |
| PLT-01 | — | Pending |
| PLT-02 | — | Pending |
| PLT-03 | — | Pending |

**Coverage:**
- v1 requirements: 19 total
- Mapped to phases: 0
- Unmapped: 19 ⚠️

---
*Requirements defined: 2026-05-06*
*Last updated: 2026-05-06 after research synthesis*
