# PPT Spec Skill

## What This Is

An AI-powered PPT generation skill that analyzes existing PPTs to extract design specifications (colors, fonts, layouts, presentation logic), then uses those specs to guide interactive content-to-PPT generation. Users provide reference PPTs to define styles, input content, and receive professional, natively editable PPTX files. Built on a forked adaptation of ppt-master's SVG→DrawingML pipeline, supporting both Opencode and Claude Code platforms.

## Core Value

Generate professional, editable PPTs that faithfully follow extracted design specifications — turning raw content into polished presentations without requiring design expertise.

## Requirements

### Validated

(None yet — ship to validate)

### Active

- [ ] **SPEC-01**: Analyze existing .pptx files to extract visual style (colors, backgrounds), typography (font families, sizes, weights), layout patterns, and presentation logic (argument flow, storytelling, data presentation)
- [ ] **SPEC-02**: Save extracted specs as project-local structured file (spec format captures visual theme, layout patterns, typography, page rhythms, and logic model)
- [ ] **SPEC-03**: User can select from available specs when initiating new PPT generation
- [ ] **GEN-01**: Smart hybrid questioning — when content input is insufficient, ask section-level overview first, then gap-fill per section with targeted questions
- [ ] **GEN-02**: Generate detailed content outline (title, content, layout description per slide) for user approval before PPT generation
- [ ] **GEN-03**: Generate natively editable .pptx files that strictly adhere to the selected spec's visual and structural rules
- [ ] **PLAT-01**: Dual-platform compatible — works as installable skill in both Opencode and Claude Code environments
- [ ] **PIPE-01**: Fork and adapt ppt-master's SVG→DrawingML conversion pipeline for spec-driven generation workflow

### Out of Scope

- Real-time collaborative editing — single-user generation workflow
- Cloud-based spec sharing — project-local spec management
- Video/audio generation (TTS narration) — PPTX generation only
- Web-based visual editor (ppt-master's visual-edit workflow) — text-driven generation only

## Context

**Reference project:** [ppt-master](https://github.com/hugohe3/ppt-master) (v2.6.0) — an open-source AI presentation system that converts source documents to natively editable PPTX via SVG→DrawingML pipeline. Key capabilities to leverage:
- Multi-step pipeline: Source → Markdown → Design Spec → SVG → PPTX
- SVG→DrawingML converter (17 modules) for native PowerPoint shapes
- 22 layout templates, 70+ chart templates, 11,600+ SVG icons
- python-pptx for PPTX construction with real DrawingML shapes

**Key differentiation from ppt-master:**
- Spec extraction from existing PPTs (not design-from-scratch)
- Spec-driven generation (reproducible style application)
- Interactive content refinement (not single-pass content injection)
- Skill format (installable, platform-agnostic) rather than standalone project

**Target platforms:** Opencode CLI, Claude Code, VS Code Copilot

## Constraints

- **Tech stack**: Python 3.10+ (python-pptx, PyMuPDF for PPT analysis), skill definition in Markdown
- **Output format**: Native .pptx with editable DrawingML shapes (no flattened images)
- **Compatibility**: Must work on macOS, Linux, and Windows
- **Skill format**: Must follow opencode/claude-code skill conventions for discoverability

## Key Decisions

| Decision | Rationale | Outcome |
|----------|-----------|---------|
| Fork ppt-master | Reuse proven SVG→DrawingML pipeline rather than rebuild | — Pending |
| Spec as project-local files | Simplicity, no infrastructure needed | — Pending |
| Smart hybrid questioning | Balances efficiency and completeness for content gathering | — Pending |
| Full logic model spec capture | User wants visual + structural + presentation logic captured | — Pending |

---
*Last updated: 2026-05-06 after initialization*
