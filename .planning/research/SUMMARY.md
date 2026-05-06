# Project Research Summary

**Project:** PPT Spec Skill (AI-powered PPT generation)
**Domain:** CLI skill for AI-assisted presentation generation with design spec extraction
**Researched:** 2026-05-06
**Confidence:** HIGH

## Executive Summary

This product is an **AI-powered PPT generation skill** that works as an installable CLI tool in Opencode and Claude Code environments. Unlike web-based competitors (Gamma, Beautiful.ai, Decktopus), it operates locally in the terminal, produces natively editable PowerPoint shapes (not flattened images), and introduces a unique workflow: **extract design specs from an existing reference PPTX**, then generate new presentations that replicate the extracted visual identity, layout patterns, and presentation logic. This "bottom-up" approach (analyze existing → replicate style) is a market gap — every competitor works top-down (choose template → add content).

The recommended approach is to **fork ppt-master's battle-tested SVG→DrawingML pipeline** (17 converter modules, production-proven in v2.6.0) as the execution engine, and build a net-new spec extraction system using python-pptx + lxml on top of it. The architecture follows a role-switching pattern where the AI agent loads specialized instruction files per phase (Spec Extractor → Interviewer → Outline Builder → PPTX Generator), uses a `spec_lock.md` anti-drift contract for generation fidelity, and enforces a single blocking gate (outline approval) between interactive and automated phases.

**Key risks**: (1) python-pptx's `Slide.background` API is broken — background extraction must use raw lxml/XML; (2) the spec extraction→generation mapping is lossy unless theme references are fully resolved to concrete HEX/font values; (3) forking ppt-master without understanding its module entanglement (dual-consumer modules like `embed_icons.py` and `flatten_tspan.py`) silently breaks the pipeline. All three are well-documented with prevention strategies and must be addressed in Phase 1–2 planning.

## Key Findings

### Recommended Stack

Python 3.12+ with python-pptx (OOXML manipulation) and lxml (deep XML traversal). The stack is entirely Python — there is no viable Node.js equivalent for SVG→DrawingML conversion or PPTX construction. Fork ppt-master's 17-module SVG→DrawingML pipeline as the core execution engine rather than building from scratch. Use uv for package management (2025–2026 standard), pytest for testing, and ruff for linting.

**Core technologies:**
- **python-pptx 1.0.2**: De facto standard for PPTX read/write in Python — reads layouts, themes, shapes; writes native DrawingML. No viable alternative exists.
- **lxml 6.1.0**: Deep OOXML traversal for spec extraction (theme XML, font schemes, color resolution) beyond python-pptx's high-level API. Also used by ppt-master's SVG parser.
- **PyYAML 6.0+**: Spec file serialization — human-readable, git-diffable, LLM-friendly structured format for design specs.
- **Pillow 12.2.0**: Extract embedded images from PPTX during spec analysis; PNG fallback rendering for Office compatibility mode.
- **cairosvg 2.7+ or svglib 1.5.0**: SVG→PNG fallback for Office versions that don't natively support SVG (LTSC 2021 and earlier). cairosvg preferred for fidelity but requires system Cairo.

**What to avoid**: Image-flattening PPTX tools (Gamma, Beautiful.ai), HTML-to-PPTX converters (lose native shapes), LibreOffice UNO API (500MB dependency, fragile), AppleScript/VBA automation (platform-locked), Node.js stack (no python-pptx equivalent), Python <3.10 (dependency incompatibility).

### Expected Features

The competitive landscape (Gamma, Beautiful.ai, Decktopus, Presentations.ai, Tome) has commoditized prompt-to-slide generation but nobody offers reverse-spec extraction or truly editable native-shape exports.

**Must have (table stakes):**
- **Editable PPTX export with native shapes** — dealbreaker if broken. The SVG→DrawingML pipeline produces real PowerPoint shapes, not flattened images. This is what competitors consistently fail at.
- **Spec extraction from existing PPTX** — core differentiator. No competitor offers this. Extracts colors, fonts, layouts, and presentation logic from reference files.
- **Content outline generation** — users must see and approve slide structure before generation. All competitors have this.
- **Template-based layout system** — baseline templates for common slide types (title, content, two-column, chart, image+text). Fork ppt-master's 22 layouts.
- **Spec file save/load** — extracted specs persist as portable, reusable artifacts (not platform-locked templates).
- **Prompt-to-slide content generation** — table stakes for AI PPT tools; LLM generates slide content from natural language input.

**Should have (differentiators):**
- **Presentation logic model capture** — extracting argument flow and storytelling patterns, not just colors/fonts. Second-order differentiator on top of spec extraction.
- **Smart hybrid questioning** — section-level overview + targeted gap-filling. Balances completeness vs. decision fatigue. Falls back to pure generation when input is sufficient.
- **Spec-driven reproducible generation** — same spec applied to different content produces consistent, branded PPTs. The spec is a complete generative blueprint.
- **Dual-platform skill packaging** — installable in both Opencode and Claude Code. Terminal-native, no browser, no account.

**Defer (v2+):**
- Chart/data visualization (use ppt-master's 70+ chart templates when needed)
- Multi-source document import (PDF, Word, URL — ppt-master handles this)
- GitHub-based spec sharing (no cloud backend)
- Multi-language content support

### Architecture Approach

Three-layer architecture: **Skill Definition Layer** (SKILL.md workflow + role reference files + Python scripts), **Orchestration Layer** (AI agent follows workflow, switches roles, manages state), and **Data Storage Layer** (project-local `specs/` directory, per-generation state). Uses four proven patterns from ppt-master: (1) role-specialized reference loading to prevent prompt contamination across phases, (2) `spec_lock.md` as anti-drift mechanism for color/font fidelity across long decks, (3) single blocking gate (outline approval) before automated pipeline execution, (4) Python scripts for deterministic work (PPTX parsing, SVG conversion, quality checking) with AI handling creative decisions (design interpretation, SVG composition).

**Major components:**
1. **Spec Extraction** — Reads PPTX via python-pptx + lxml; extracts colors, fonts, layouts, logic model. Net-new capability, not from ppt-master. Outputs `design_spec.md`.
2. **Spec Management** — Save/load/list/validate spec files in project-local `specs/` directory. Filesystem-native, no database.
3. **Content Gathering** — AI-agent-driven smart hybrid questioning (section overview → gap analysis → targeted filling). Pure conversational flow, no Python scripts.
4. **Outline Generation** — AI-agent-driven: converts gathered content + loaded spec into slide-by-slide outline with layout assignments. Blocking gate for user approval.
5. **PPTX Generation** — Forked from ppt-master: SVG generation per page (AI), spec_lock anti-drift re-read, SVG quality check, SVG finalize (icon embedding, text flattening), and SVG→DrawingML conversion (~10–17 modules).

### Critical Pitfalls

1. **PPTX Background API Is Broken** — python-pptx's `Slide.background` returns `<p:cSld>` instead of `<p:bg>`, corrupting shape trees. **Avoid by**: using raw lxml/XML for background extraction; placing full-slide rects in SVG for backgrounds in generation.

2. **SVG `<tspan>` Fragmentation** — `x`/`y`/`dy` attributes on tspans within the same logical line cause text to split into independent text frames in PPTX, destroying editability. **Avoid by**: enforcing strict tspan discipline — inline formatting tspans must not carry positioning attributes; run `tspan_flattener` validation before export.

3. **Spec Extraction → Authoring Gap** — Extracted specs use theme-aware OOXML references (e.g., `accent1`) instead of resolved HEX values. Font inheritance from slide masters isn't flattened. Layout patterns are implicit spatial relationships. **Avoid by**: two-pass extraction with resolution/normalization pass; define the spec schema before writing extraction code.

4. **SVG→DrawingML Feature Gap Assumptions** — AI-generated SVGs routinely use `<mask>`, `<style>` tags, `@font-face`, `rgba()` color syntax, and nested `<use>` chains — all banned in DrawingML. **Avoid by**: adopting ppt-master's banned-feature blacklist wholesale; running `svg_quality_checker.py` before conversion, not after.

5. **Forking Without Understanding Pipeline Entanglement** — ppt-master's modules (`embed_icons.py`, `flatten_tspan.py`, `svg_rect_to_path.py`) are called from both the disk pipeline AND the in-memory converter for different purposes. Deleting a "seemingly unnecessary" module silently breaks icons, text, or shape rendering. **Avoid by**: reading ppt-master's full technical design docs; mapping the full module dependency graph before stripping anything.

## Implications for Roadmap

Based on research, suggested phase structure:

### Phase 1: Pipeline Foundation (Fork & Adapt ppt-master)
**Rationale:** The SVG→DrawingML pipeline is the execution engine. It has no dependencies on spec extraction or content gathering and represents the highest technical risk. Validating it early de-risks everything downstream.
**Delivers:** Forked and adapted `svg_finalize/` + `svg_to_pptx/` modules; standalone pipeline that converts sample SVGs to native-shape PPTX; python-pptx + lxml + cairosvg dependencies installed and verified.
**Addresses:** PIPE-01 (SVG→DrawingML pipeline)
**Avoids:** Pitfall #6 (pipeline entanglement) — read full technical-design.md, map module dependency graph before stripping. Pitfall #4 (SVG feature gaps) — adopt banned-feature blacklist from day one.
**Research needed:** YES — ppt-master internals, module dependency mapping, which modules are dual-consumer.

### Phase 2: Spec Extraction System
**Rationale:** Spec extraction is the core differentiator and enables everything in the generation pipeline. It has no dependencies on Phase 1 (read-only, no SVG→DrawingML needed) and can theoretically run in parallel, but front-loading it after Phase 1 ensures the spec format is designed with generation requirements in mind.
**Delivers:** `analyze_pptx.py` (extract colors, fonts, layouts, logic from PPTX); `spec_manager.py` (save/load/list/validate specs); defined spec format schema (spec-format.md); project-local `specs/` directory convention.
**Addresses:** SPEC-01 (spec extraction), SPEC-02 (spec save/load)
**Avoids:** Pitfall #1 (background API trap) — use raw lxml from the start. Pitfall #3 (spec extraction gap) — two-pass extraction, schema-first design.
**Research needed:** YES — OOXML namespace traversal, theme resolution algorithms, layout pattern classification.

### Phase 3: Spec-Driven Generation
**Rationale:** Connects Phase 1's pipeline to Phase 2's specs. This is the integration that proves the core value proposition: extract a spec from PPTX A, generate PPTX B that matches its design. Must come before content gathering because the generation path needs to work with specs before adding the content layer.
**Delivers:** `spec_to_lock.py` (convert spec → spec_lock contract); pptx-generator role file (spec-locked Executor variant); spec_lock anti-drift mechanism; verified round-trip (extract → generate → compare).
**Addresses:** GEN-03 (spec-driven generation), PIPE-01 (SVG→DrawingML integration)
**Avoids:** Pitfall #8 (empty spec_lock defaults) — explicit layout binding per page in spec schema.
**Research needed:** YES — spec-to-layout binding design, anti-drift mechanism adaptation.

### Phase 4: Content Gathering & Outline
**Rationale:** The interactive UX layer. Can be built in parallel with Phase 2–3 (no pipeline dependency), but placed after because the content gathering strategy should be informed by what the spec-driven generation can handle. Users need to see the full flow before optimizing the questioning UX.
**Delivers:** interviewer.md role (smart hybrid questioning strategy); outline-builder.md role (outline generation from gathered content + spec); defined outline format; approval gate UX.
**Addresses:** GEN-01 (smart hybrid questioning), GEN-02 (content outline generation)
**Avoids:** Pitfall #7 (interactive UX fatigue) — bundle confirmations into one gate, pre-fill with recommendations, cap at 8 questions total.
**Research needed:** Light — well-documented UX patterns from ppt-master, but questioning strategy needs tuning.

### Phase 5: Integration & Platform Packaging
**Rationale:** Wires all components together into a working skill, packages for dual-platform distribution, and runs end-to-end verification. Must come last — depends on everything else working.
**Delivers:** Complete SKILL.md (workflow, role switching, serial pipeline rules); dual-platform installation support (Opencode + Claude Code); end-to-end verification: PPTX in → spec extracted → content gathered → PPTX out; documentation (spec-format.md, canvas-formats.md, shared-standards.md).
**Addresses:** PLAT-01 (dual-platform skill format), all reference files (shared-standards.md, canvas-formats.md)
**Avoids:** Pitfall #5 (multi-platform divergence) — thin platform wrappers, `${SKILL_DIR}` resolution, tested on both platforms.
**Research needed:** YES — Claude Code skill marketplace format, Opencode skill discovery mechanics, dual bootstrapping.

### Phase Ordering Rationale

- **Phase 1 first** because it's the highest technical risk (entangled modules, banned SVG features) and has zero dependencies. If the pipeline doesn't work standalone, nothing else matters.
- **Phase 2 before Phase 3** because spec generation needs the spec format to exist. While they could be designed together, the extraction system benefits from being designed with the generation contract in mind.
- **Phase 3 after Phase 1+2** because it's the integration point between the pipeline and the spec system. This is where the core value proposition is validated.
- **Phase 4 can float** — it's independent of the pipeline and spec extraction. Placing it after Phase 3 lets the content gathering UX be informed by what the generation system actually produces. In a team setting, Phase 4 could run in parallel with Phase 2–3.
- **Phase 5 last** because it's integration and packaging — it needs everything working.

### Research Flags

Phases likely needing deeper research during planning:
- **Phase 1 (Pipeline):** ppt-master module dependency graph, dual-consumer module identification, which modules can safely be removed. High risk of breaking the pipeline without full understanding.
- **Phase 2 (Spec Extraction):** OOXML theme resolution algorithms, font inheritance flattening, layout pattern classification (not just spatial clustering — semantic classification). Complex domain with sparse documentation.
- **Phase 3 (Spec-Driven Generation):** spec_lock anti-drift mechanism adaptation, layout binding design, getting the spec→SVG parameter mapping right. Integration risk between new spec system and forked pipeline.
- **Phase 5 (Platform):** Claude Code skill marketplace format, Opencode skill discovery mechanics. Both platforms are evolving — documentation may be out of date.

Phases with standard patterns (skip research-phase):
- **Phase 4 (Content Gathering):** Smart questioning UX is well-documented from ppt-master's Eight Confirmations pattern. Outline generation is a straightforward LLM call with role guidance. Minimal research needed — mostly UX design.

## Confidence Assessment

| Area | Confidence | Notes |
|------|------------|-------|
| Stack | HIGH | All packages verified against PyPI with exact versions. ppt-master dependency list confirmed from requirements.txt. Alternatives considered systematically. |
| Features | MEDIUM | Competitor features from live web scraping (Decktopus, Presentations.ai). Gamma/Beautiful.ai/Tome from biased competitor comparison pages. Market gap analysis is strong but competitor feature completeness may have blind spots. |
| Architecture | HIGH | ppt-master v2.6.0 architecture verified against actual source tree, technical design docs, SKILL.md workflow. Build order derived from real dependency graph. |
| Pitfalls | HIGH | Sourced from python-pptx GitHub issues (confirmed open bugs), ppt-master technical design docs (hard-won lessons), and ppt-master GitHub issues (real failure modes). Prevention strategies tested in production by ppt-master. |

**Overall confidence: HIGH** — the adapted components (ppt-master pipeline) are production-proven in v2.6.0, and the new components (spec extraction, content gathering) follow the same architectural patterns. The stack is mature, the pitfalls are well-documented with fixes, and the competitive landscape has clear gaps.

### Gaps to Address

- **PPTX export fidelity across PowerPoint versions**: Research confirmed the SVG→DrawingML pipeline works but did not exhaustively test across PowerPoint 2016/2019/2021/LTSC/365. Plan to test export on at least 3 versions during Phase 1 verification.
- **Claude Code skill marketplace registration**: Research confirmed Claude Code supports `~/.claude/skills/` installation but did not verify marketplace publishing process or `marketplace.json` format. Resolve during Phase 5 planning.
- **Opencode skill discovery specifics**: Research confirmed Opencode uses `~/.opencode/skills/` with frontmatter-described skills but exact discovery mechanics (directory scan depth, supported fields beyond `description`/`color`/`tools`) need validation during Phase 5.
- **Large PPTX file parsing performance**: lxml.iterparse is the recommended approach for >50-slide decks but was not benchmarked. Test during Phase 2 implementation.
- **Spec format design for presentation logic model**: The spec schema sections for "presentation logic" (argument flow, storytelling patterns) are conceptually scoped but the exact schema fields and classification taxonomy need design during Phase 2 planning.

## Sources

### Primary (HIGH confidence)
- [python-pptx 1.0.2 on PyPI](https://pypi.org/project/python-pptx/) — verified latest version, release date, Python version compatibility
- [lxml 6.1.0 on PyPI](https://pypi.org/project/lxml/) — verified latest version, wheel support through Python 3.14
- [PyMuPDF 1.27.2.3 on PyPI](https://pypi.org/project/PyMuPDF/) — verified latest version, license (AGPL)
- [Pillow 12.2.0 on PyPI](https://pypi.org/project/pillow/) — verified latest version, Python 3.10+ requirement
- [ppt-master v2.6.0](https://github.com/hugohe3/ppt-master) — architecture, pipeline, dependency list, bans, build manifest
- [ppt-master skills/requirements.txt](https://raw.githubusercontent.com/hugohe3/ppt-master/main/skills/ppt-master/requirements.txt) — verified dependency declarations
- [ppt-master docs/technical-design.md](https://github.com/hugohe3/ppt-master) — pipeline architecture, SVG constraints, module entanglement, anti-drift mechanism
- [ppt-master skills/ppt-master/SKILL.md](https://github.com/hugohe3/ppt-master/tree/main/skills/ppt-master) — workflow definition, role switching protocol, execution discipline
- [python-pptx docs](https://python-pptx.readthedocs.io/en/latest/) — official documentation, known limitations
- [python-pptx GitHub issue #1126](https://github.com/scanny/python-pptx/issues/1126) — Slide.background._element bug confirmation

### Secondary (MEDIUM confidence)
- **Decktopus.com** — live scraping 2026-05-06. Features: AI prompt→deck, brand upload, PDF import, team workspaces.
- **Presentations.ai** — live scraping of home page + comparison pages. Features: document import, Brand Sync, 10,000+ templates, SOC 2 Type II. Competitor comparison pages contain biased but detailed feature critique.
- **Gamma.app / Beautiful.ai / Tome.app** — features characterized via Presentations.ai comparison pages (single biased source).
- **Opencode official site** (opencode.ai) — skill ecosystem conventions.
- [python-pptx GitHub issue #1111](https://github.com/scanny/python-pptx/issues/1111) — Font.color getter mutates XML
- [python-pptx GitHub issue #1112](https://github.com/scanny/python-pptx/issues/1112) — SVG image not supported
- [ppt-master GitHub issues](https://github.com/hugohe3/ppt-master/issues) — #90 (format error), #61 (layout aesthetics), #60 (markdown issues), #62 (HTML entities in SVG), #74 (skill capability), #67 (OpenCode support)

### Tertiary (LOW confidence)
- ppt-master FAQ.md — model recommendations, layout troubleshooting, template derivation guidance. Community-maintained, may not reflect v2.6.0 accurately.

---
*Research completed: 2026-05-06*
*Ready for roadmap: yes*
