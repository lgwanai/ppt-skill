# Feature Research

**Domain:** AI-powered PPT generation skill (Opencode/Claude Code)
**Researched:** 2026-05-06
**Confidence:** MEDIUM

> Sources: Live web scraping of Decktopus.com, Presentations.ai (including their Gamma/Beautiful.ai/Tome comparison pages). ppt-master known from PROJECT.md context. Gamma/Beautiful.ai/Tome features corroborated across Presentations.ai's direct competitor comparison content. No Brave Search available — MEDIUM confidence on web-only findings, adjusted for corroboration across multiple competitor sources.

## Feature Landscape

### Table Stakes (Users Expect These)

Features users assume exist. Missing these = product feels incomplete.

| Feature | Why Expected | Complexity | Notes |
|---------|--------------|------------|-------|
| **Prompt-to-slide content generation** | Every AI PPT tool (Gamma, Beautiful.ai DesignerBot, Decktopus, Presentations.ai) generates slide content from natural language prompts. Users expect "write a presentation about X" to produce slide-level output | MEDIUM | LLM call for content + layout mapping. Core competency of all competitors |
| **Multi-source input** | Presentations.ai imports PDF, Word, URL, text. Decktopus imports PDF. Gamma imports text. Users expect to feed existing content, not retype | MEDIUM | Need PDF/Word text extraction + URL content fetcher. ppt-master already handles document import |
| **Content outline generation** | Decktopus, Presentations.ai both generate slide outlines for review before full generation. Tome generates page outlines. Users expect to see and approve structure before design | LOW | LLM call for outline + simple CLI display. Already scoped in GEN-02 |
| **Template-based layout** | Every competitor has template libraries: Presentations.ai has 10,000+, ppt-master has 22 layouts + 70 chart types, Gamma has preset themes. Users expect visual variety | HIGH | Requires layout template system with layout-to-SVG mapping. Inherit from ppt-master's 22 templates |
| **Brand/theme application** | Decktopus auto-brands from upload, Presentations.ai brands from URL, Beautiful.ai has smart templates. Users expect their colors/fonts to be applied consistently | MEDIUM | Theme variables applied to templates. Our spec system is more powerful — extract from PPTX, not just config |
| **Editable PPTX export** | Presentations.ai exports to .pptx and Google Slides. Decktopus exports. Beautiful.ai exports (with issues). Gamma exports are broken. Users expect native, editable PowerPoint output. **Dealbreaker if broken** | HIGH | python-pptx + DrawingML pipeline from ppt-master fork. This is the hardest and most critical feature — corrupt exports kill all credibility |
| **Text + image slides** | All tools produce slides with headings, body text, and images. Baseline expectation | LOW | Core SVG template rendering |
| **Chart/data visualization support** | Presentations.ai connects to live data. Gamma generates charts from data. ppt-master has 70+ chart templates. Users presenting data expect charts | HIGH | Requires chart generation pipeline (matplotlib → SVG or direct python-pptx charts). ppt-master Chart2SVG modules |
| **Icon/graphic libraries** | ppt-master has 11,600+ SVG icons. Presentations.ai has vast libraries. Users expect visual variety without manual asset sourcing | MEDIUM | Inherit ppt-master's icon library. No new work needed |
| **Editable shapes (not flattened images)** | ppt-master's key differentiator: SVG→DrawingML creates native PowerPoint shapes. All competitors that export images-only are inferior. Users expect real text boxes they can edit | HIGH | 17-module SVG→DrawingML converter. This is why we fork ppt-master |

### Differentiators (Competitive Advantage)

Features that set the product apart. Not required, but valuable.

| Feature | Value Proposition | Complexity | Notes |
|---------|-------------------|------------|-------|
| **Spec extraction from existing PPTX files** | NO competitor does this. Users provide a reference PPTX and the tool extracts colors, fonts, layouts, and presentation logic for reuse. This is the one-way bridge from "I have a good-looking deck" to "I need more decks like this." Gamma/Beautiful.ai/Decktopus all work top-down (choose template, add content). We work bottom-up (analyze existing, replicate style) | HIGH | SPEC-01. Requires python-pptx or PyMuPDF to parse PPTX XML, extract theme elements, identify layout patterns, classify page types. Most complex differentiator |
| **Presentation logic model capture** | Extracting not just colors/fonts but the *argumentation model* — title slides, section dividers, content density rhythms, storytelling patterns. No tool captures presentation logic as a reusable artifact | HIGH | Goes beyond visual specs. Requires AI analysis of slide sequencing, content patterns, narrative structure detection. Builds on spec extraction but adds semantic layer |
| **Spec-as-file portability** | Saved specs are project-local structured files that can be versioned, shared, reused. Gamma/Decktopus lock templates to their platform. Our specs are portable artifacts you own | LOW | SPEC-02. Simple file format (YAML/JSON) with visual + layout + logic sections |
| **Smart hybrid questioning** | When user input is insufficient (no topic, no outline), the tool asks section-level questions first, then gap-fills per section. Competitors either force you through a rigid wizard (Decktopus) or generate blind (Gamma). Our approach balances efficiency and completeness | MEDIUM | GEN-01. LLM-driven dialog with structured questioning strategy. Falls back to pure generation when input is sufficient — no forced wizard |
| **Skill-format delivery (not web app)** | Works as an installable skill in Opencode and Claude Code. No account, no browser, no subscription. Users generate PPTs in their terminal alongside their code workflow. Only Copilot has a PowerPoint integration, but it's a plugin, not a standalone skill | LOW | PLAT-01. Markdown skill definition files + Python backend. Platform-agnostic design |
| **Spec-driven reproducible generation** | Apply the same extracted spec to different content, producing consistent PPTs. Presentations.ai has brand sync but it's manual config, not extracted. Gamma/Beautiful.ai require choosing templates per-deck. Our spec is a complete generative blueprint | MEDIUM | GEN-03. Spec → Template selection → SVG rendering → DrawingML pipeline. The spec drives all decisions deterministically |
| **SVG→DrawingML native shapes** | ppt-master's 17-module converter produces real PowerPoint shapes (not images). Text remains editable in PowerPoint. Beautiful.ai/Decktopus export images. This is the "editable shapes" differentiator, inherited from ppt-master fork | HIGH | PIPE-01. Already built in ppt-master. Adaptation work needed to make spec-driven (vs ppt-master's design-from-scratch approach) |
| **No internet dependency for core generation** | Since it's a local skill using python-pptx, the generation pipeline works offline. Gamma/Tome/Presentations.ai are all cloud-dependent. For sensitive business presentations, this is a privacy win | LOW | Architecture decision: local Python runtime for generation, LLM for content/AI tasks |

### Anti-Features (Commonly Requested, Often Problematic)

Features that seem good but create problems.

| Feature | Why Requested | Why Problematic | Alternative |
|---------|---------------|-----------------|-------------|
| **Web-based visual editor (WYSIWYG)** | Users expect to see and tweak slides visually before export. Gamma, Decktopus, Beautiful.ai all have visual editors | Doubles complexity (maintain visual editor + generation engine). Creates multi-format confusion (Gamma's editor was their failed 2020 product). The skill format means CLI/terminal, not a browser. ppt-master dropped its visual-edit workflow for good reason | Content outline approval + instant PPTX export. User opens in PowerPoint to verify. The PPTX IS the editor |
| **Real-time collaboration** | Presentations.ai has real-time sync, Decktopus has team workspaces | Server infrastructure, authentication, state sync. Violates the "project-local, no infrastructure" constraint. Adds months of engineering | Share spec files + generated PPTXs via existing tools (Git, Slack). The skill is single-user |
| **AI image generation in slides** | Decktopus generates AI images for slides. Gamma/Beautiful.ai include image gen | Requires image API (DALL-E, Midjourney), adds cost/complexity, generated images often look AI-generated and reduce professional feel | Use stock image placeholders from ppt-master's icon library. User can replace with their own images in PowerPoint |
| **Web-first/HTML output** | Gamma outputs interactive web presentations ("gammas"), not just PPTX | Competing format to PPTX. Doubles output pipeline. Gamma's web-first approach killed their PPTX export quality | PPTX only. Users who need web can upload to Google Slides or export from PowerPoint |
| **Cloud-based spec/template sharing** | Team users want shared template libraries like Decktopus's organization panel | Requires backend, auth, cloud storage. Violates project-local constraint. Scope creep for a CLI skill | Spec files are portable — share via Git or file sharing. Could add later as v2 if validated |
| **TTS narration / video export** | Decktopus has "Vocal Decks" with audio narration. Some users want auto-narrated video exports | Entirely different pipeline (TTS, audio syncing, video encoding). Massively increases complexity. Already listed as out-of-scope | Share PPTX, user records narration in PowerPoint |
| **Multi-format output (websites, docs, social posts)** | Gamma does presentations + websites + documents. Simplified does presentations + social posts + ads | Total loss of focus. Gamma's scatter-shot approach is why their presentations are generic and their exports break. "When you can't win at one thing, try everything." Tome pivoted to sales/marketing-only | PPTX only. Be the best at one format |
| **Live data-connected presentations** | Presentations.ai refreshes charts from live data sources | Requires data connectors, scheduling, API integration. Huge infrastructure dependency. Not MVP | User exports data to CSV/Excel, feeds into spec. Manual refresh is fine for MVP |
| **Auto-branding from URL** | Presentations.ai extracts brand from company URL. Users expect "just apply my brand" | URL scraping for brand is fragile and unreliable. Spec extraction from PPTX is more precise and user-controlled | Spec extraction from PPTX. User provides the reference PPTX — that IS their brand |
| **AI icebreakers/speaker notes** | Decktopus generates hooks, icebreakers, speaker scripts | Blurs line between content generator and speech coach. Adds LLM complexity without core PPT value. Slide content already serves as speaker notes | Optional speaker notes section in spec format. Only if trivial to add |

## Feature Dependencies

```
Spec Extraction (SPEC-01)
    └──requires──> PPTX parsing (python-pptx/PyMuPDF)
    └──requires──> Layout pattern recognition
    └──requires──> Color/font extraction
    └──enables──> Spec file generation (SPEC-02)

Spec File (SPEC-02)
    └──requires──> Spec extraction (SPEC-01)
    └──enables──> Spec-driven generation (GEN-03)
    └──enables──> Spec selection UI (SPEC-03)

Content Outline Generation (GEN-02)
    └──requires──> LLM integration
    └──requires──> Template/layout awareness (to suggest layout types)
    └──enables──> PPTX generation (GEN-03)

PPTX Generation (GEN-03)
    └──requires──> Spec file (SPEC-02)
    └──requires──> Content outline (GEN-02)
    └──requires──> SVG→DrawingML pipeline (PIPE-01)
    └──requires──> Template layouts

SVG→DrawingML Pipeline (PIPE-01)
    └──requires──> ppt-master fork baseline
    └──requires──> python-pptx
    └──enables──> PPTX generation (GEN-03)

Smart Questioning (GEN-01)
    └──requires──> LLM integration
    └──enhances──> Content outline generation (GEN-02)
    └──independent of──> Spec pipeline (can work without spec)

Skill Format (PLAT-01)
    └──independent of──> All generation features (packaging only)
    └──requires──> Markdown skill definition files
    └──requires──> Python runtime + dependencies

Spec Selection (SPEC-03)
    └──requires──> Spec files exist (SPEC-02)
    └──enables──> Spec-driven generation (GEN-03)
```

### Dependency Notes

- **SPEC-01 (Spec Extraction) is the critical path.** It enables SPEC-02, which enables everything downstream. Without spec extraction, we're just another prompt-to-PPT tool. This is phase 1.
- **PIPE-01 (SVG→DrawingML) is the execution path.** Without it, we generate crude python-pptx slides indistinguishable from basic scripts. The 17-module converter is what makes output professional.
- **GEN-01 (Smart Questioning) is independent** — it enhances content gathering but doesn't block the core pipeline. Can be built in parallel with spec extraction.
- **Content outline + PPTX generation are the visible result.** Everything else is infrastructure. Users judge based on output quality.
- **Spec selection (SPEC-03) is trivial once spec files exist.** A file listing/scanning feature.

## MVP Definition

### Launch With (v1)

Must ship these to validate the concept.

- [ ] **Spec extraction from PPTX (SPEC-01)** — Core differentiator. If we can't extract specs, we're just another prompt-to-PPT tool. This IS the value proposition.
- [ ] **Spec file save/load (SPEC-02)** — Extracted specs must be persisted as reusable artifacts. Without this, extraction has no purpose.
- [ ] **Content outline generation (GEN-02)** — Users must see and approve slide structure before generation. All competitors have this.
- [ ] **PPTX generation with native shapes (GEN-03 + PIPE-01)** — The output. Must be editable, professional, and spec-compliant. The SVG→DrawingML pipeline from ppt-master is non-negotiable.
- [ ] **Template layout system** — Baseline templates for common slide types (title, content, two-column, chart, image+text). Without templates, there's nothing to apply specs to.

### Add After Validation (v1.x)

Features to add once core extraction + generation pipeline works.

- [ ] **Smart hybrid questioning (GEN-01)** — Add once the basic prompt→outline→PPT flow is solid. Improves UX for vague inputs but doesn't block core pipeline.
- [ ] **Spec selection UI (SPEC-03)** — List available specs, select one for new generation. Trivial once specs exist.
- [ ] **Presentation logic model in spec** — Enhance spec extraction to capture argument flow, not just colors. Phase 2 of spec extraction.
- [ ] **Multi-source input (PDF, Word, URL)** — Expand beyond text prompts to document imports. ppt-master can handle this.
- [ ] **Dual-platform skill packaging (PLAT-01)** — Opencode format + Claude Code format. Packaging work, not engine work.

### Future Consideration (v2+)

Features to defer until product-market fit is established.

- [ ] Chart/data visualization (use ppt-master charts)
- [ ] Advanced layout templates beyond ppt-master's 22
- [ ] Spec validation/repair tools
- [ ] Batch generation from multiple content sources with same spec
- [ ] GitHub-based spec sharing (no cloud backend)
- [ ] Multi-language content support

## Feature Prioritization Matrix

| Feature | User Value | Implementation Cost | Priority |
|---------|------------|---------------------|----------|
| Spec extraction from PPTX | HIGH | HIGH | P1 |
| Spec file save/load | HIGH | LOW | P1 |
| Content outline generation | HIGH | LOW | P1 |
| PPTX generation (native shapes) | HIGH | HIGH | P1 |
| Template layout system | HIGH | MEDIUM | P1 |
| SVG→DrawingML pipeline fork | HIGH | HIGH | P1 |
| Smart hybrid questioning | MEDIUM | MEDIUM | P2 |
| Spec selection UI | MEDIUM | LOW | P2 |
| Logic model in spec | HIGH | HIGH | P2 |
| Multi-source input (PDF/Word/URL) | MEDIUM | MEDIUM | P2 |
| Dual-platform skill packaging | MEDIUM | LOW | P2 |
| Chart generation | MEDIUM | HIGH | P3 |
| Batch generation | LOW | MEDIUM | P3 |
| Multi-language support | LOW | MEDIUM | P3 |

**Priority key:**
- P1: Must have for launch
- P2: Should have, add when possible
- P3: Nice to have, future consideration

## Competitor Feature Analysis

| Feature | Gamma.app | Beautiful.ai | Tome.app | Decktopus | Presentations.ai | ppt-master | Our Approach |
|---------|-----------|-------------|----------|-----------|-----------------|------------|-------------|
| AI content generation | Prompt → slides | DesignerBot prompt → slides | Prompt → tiles (paid only) | Prompt → full deck | Prompt/document → deck | Source → Markdown → SVG | Prompt → outline → PPTX |
| Document import | Text paste only | Text input only | Limited | PDF import | PDF, Word, URL, text | Source docs | Text initially; PDF/Word in v1.x |
| Design spec management | Manual themes | Manual brand setup | Fonts/colors only | Brand upload | Brand sync from URL | Design spec stage | **Spec extraction from PPTX (unique)** |
| PowerPoint export | Broken (overlap, missing fonts) | Issues reported | Not supported | Yes (export) | High-fidelity | Native PPTX | **Native DrawingML shapes (best-in-class)** |
| Template library | Preset themes | Smart templates (limited) | Basic tile blocks | Templates available | 10,000+ | 22 layouts, 70+ charts | Fork ppt-master's 22 layouts |
| Editable output | Web-first, PPTX broken | Partially editable | Not editable | Editable export | Editable PPTX | **Fully editable shapes** | **Fully editable shapes (inherit from ppt-master)** |
| Collaboration | Sharing | Basic sharing | Real-time (limited) | Team workspaces | Real-time sync | None | **None (by design — single-user skill)** |
| Enterprise security | No SOC 2 | Not certified | N/A | Not certified | SOC 2 Type II | N/A | **N/A (local execution, no server)** |
| Platform | Web app | Web app | Web app (pivoting) | Web app | Web app + API | Python project | **CLI skill (Opencode + Claude Code)** |
| Pricing | Free tier + $10/mo | $40/user/mo | $20-60/user/mo (no free AI) | Paid tiers | Free tier + $8-16/mo | Open source | **Free (skill + local Python)** |

## Market Gap Analysis

### What Everyone Has (commoditized)
- Prompt-to-slide content generation
- Some form of template/theme system
- Export to PPTX (with varying quality)
- Basic image/text placeholder slides

### What Some Have (emerging standard)
- Document import (PDF, Word, URL)
- Brand auto-application
- Content outline generation
- Team collaboration

### What Nobody Has (our wedge)
- **Reverse-spec from existing PPTX:** The ability to look at "that one great deck Legal made" and say "make more like this." Every competitor works top-down (choose template → add content). We alone work bottom-up (analyze existing → replicate style).
- **Presentation logic extraction:** Capturing not just colors/fonts but the storytelling pattern and argument structure as a reusable model.
- **Spec-as-portable-artifact:** Competitors lock templates to their platform. Our specs are files you own, version, share.
- **Terminal-native generation:** A skill that lives in your CLI, generates PPTXs alongside your code, no browser context switch.
- **Truly editable shapes:** ppt-master's SVG→DrawingML pipeline produces native PowerPoint shapes. Most competitors flatten slides to images or produce broken exports.

### Competitive Positioning

This is NOT a Gamma/Decktopus competitor. Those are web-first SaaS products with visual editors, collaboration, and cloud hosting. This is a developer tool — a CLI skill that bridges "I have a reference PPT" to "generate more PPTs like this" using AI content generation and a shape-level PPTX pipeline.

**Target user:** A professional who has a good-looking PPTX and needs to produce consistent, branded presentations from content outlines — all from their terminal.

**Unfair advantage:** Local execution (privacy), editable shapes (quality), spec extraction (unique workflow), open-source pipeline (no subscription).

## Sources

- **Decktopus.com** — Live scraping. Features: AI prompt→deck, brand upload, drag-drop editor, PDF import, Zapier integration, AI image generation, team workspaces, auto-branded slide library. (2026-05-06)
- **Presentations.ai** — Live scraping of home page + comparison pages. Features: document import (PDF/Word/URL), Brand Sync, 10,000+ templates, anti-fragile templates, SOC 2 Type II, API, live data, analytics, multilingual support. Comparison pages contain direct feature critique of Gamma, Beautiful.ai, and Tome. (2026-05-06)
- **Gamma.app** — Features characterized via Presentations.ai's Gamma comparison page (biased but detailed). Prompt generation, multi-format output, broken PPTX export, 50M users, no SOC 2, retrofitted AI architecture. (2026-05-06, MEDIUM confidence — single biased source)
- **Beautiful.ai** — Features characterized via Presentations.ai's Beautiful.ai comparison page. Smart templates, DesignerBot, manual brand setup, PPTX export issues, $40/user/mo. (2026-05-06, MEDIUM confidence — single biased source)
- **Tome.app** — Features characterized via Presentations.ai's Tome comparison page. Tile system, pivot from presentations, no PPTX export, 1.5/5 rating. (2026-05-06, MEDIUM confidence — single biased source)
- **ppt-master** — Features from PROJECT.md context: Multi-step pipeline (Source→Markdown→Design Spec→SVG→PPTX), 17-module SVG→DrawingML converter, 22 layouts, 70+ charts, 11,600+ icons, python-pptx. GitHub repo: github.com/hugohe3/ppt-master. (HIGH confidence from project documentation)

---
*Feature research for: PPT Spec Skill (AI-powered PPT generation)*
*Researched: 2026-05-06*
