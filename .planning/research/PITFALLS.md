# Pitfalls Research

**Domain:** AI PPT Generation Skill (PPTX + SVG→DrawingML + Spec Extraction)
**Researched:** 2026-05-06
**Confidence:** HIGH

## Critical Pitfalls

### Pitfall 1: PPTX Backgrounds Are Not Shapes — python-pptx Background API Is a Trap

**What goes wrong:**
Attempting to read or reconstruct slide backgrounds via python-pptx's `Slide.background` property silently corrupts the slide's shape tree (spTree). The `_element` getter returns `<p:cSld>` instead of `<p:bg>`, so background-related code operates on the wrong XML node and breaks shape inheritance. This is python-pptx issue #1126 (open since 2024). When extracting design specs from reference PPTX files, background color/fill extraction will fail silently or corrupt the source file.

**Why it happens:**
python-pptx's internal element resolution for `Slide.background` returns the Common Slide Data node (`<p:cSld>`) rather than the actual background element (`<p:bg>`). The library's OOXML abstraction layer has known gaps around slide masters, layouts, and background inheritance — areas where the OOXML spec is complex and the library's coverage is incomplete (1.0.0 as of 2024).

**How to avoid:**
Use direct lxml/XML parsing for background extraction instead of python-pptx's high-level API. Parse `<p:bg>` from the slide XML directly using `Slide._element.find()` with the correct OOXML namespace. For the spec extraction phase, handle backgrounds at the raw XML level. For generation, place a full-slide `<rect>` as the top-back element in SVG (as ppt-master does) rather than attempting to set a PowerPoint slide background programmatically.

**Warning signs:**
- Background color appears transparent/white when slide's layout clearly has a colored background
- `Slide.background.fill` returns `None` for slides with visible backgrounds
- Shape tree corruption errors after background reads on macOS PowerPoint 16.79+
- Slide-copy operations lose background formatting

**Phase to address:**
Spec Extraction Phase (SPEC-01). Must establish XML-level background parsing early; attempting python-pptx background API mid-stream will cause rewrite.

---

### Pitfall 2: SVG `<tspan>` Handling Splits Lines When It Shouldn't

**What goes wrong:**
In the SVG→DrawingML converter, `<tspan>` elements with `x`/`y`/`dy` attributes on the same logical line cause text to be split into independent text frames in the output PPTX. A single paragraph with inline formatting gets fragmented into separate un-selectable text boxes, destroying editability. Issue specifically confirmed in ppt-master's own documentation (shared-standards.md) and the tspan_flattener.py module exists solely to prevent this problem.

**Why it happens:**
DrawingML text runs (`<a:r>`) cannot reposition mid-paragraph. When a `<tspan>` carries `x` or `dy`, the converter interprets it as a new line/paragraph boundary and creates a separate `<a:p>` (paragraph) element. `x`-anchored tspans cause column jumps where the columns lose alignment. The AI Executor generating SVGs routinely produces multi-tspan text without understanding the converter's structural constraints.

**How to avoid:**
Enforce strict `<tspan>` discipline in the AI executor instructions:
- Same-line inline formatting tspans MUST NOT carry `x`/`y`/`dy` attributes — only `fill`, `font-weight`, `font-size`, `text-decoration`
- Only set `x`/`y`/`dy` on tspans that genuinely start a new line
- Multi-column layouts MUST use separate `<text>` elements, not `x`-jump tspans
- Run tspan_flattener validation before export; flag violations at quality-checker stage

**Warning signs:**
- Text boxes in output PPTX are fragmented — selecting one word doesn't select the whole line
- Alignment drift between adjacent text runs after PPTX export
- Multi-column text loses layout structure entirely in exported PPTX

**Phase to address:**
SVG→DrawingML Pipeline Phase (PIPE-01). The tspan discipline rules must be embedded in executor SVGs from the start.

---

### Pitfall 3: Design Spec Extraction → Authoring Gap

**What goes wrong:**
Extracting visual specs from an existing PPTX (colors, fonts, layouts) produces data that is structurally incompatible with the AI generation pipeline. Colors are extracted as theme-aware OOXML references (e.g., `accent1` → depends on slide master theme), not as resolved HEX values. Fonts may be inherited from slide layouts rather than explicitly set on text runs. Layout patterns (multi-column, grid, image placement) are implicit spatial relationships with no programmatic representation. The extracted spec looks complete but produces garbled output when fed into generation.

**Why it happens:**
The spec-to-generation mapping is a lossy translation between two domains: (1) the PowerPoint object model (themes, masters, layouts, inheritance) and (2) absolute-coordinate SVG with explicit HEX/font-face on every element. The intermediate spec format must resolve all inheritance chains, calculate absolute positions, and parameterize layout patterns — but most implementations skip this and produce "specs" that are just property dumps without semantic structure.

**How to avoid:**
Build the spec extraction as a two-pass process: (1) raw property extraction from OOXML XML (not via python-pptx high-level API), and (2) resolution/normalization pass that converts theme references to concrete values, calculates absolute coordinates from inherited layout positions, and classifies layout patterns into reusable templates. The spec format must be the *input contract* for the generation phase, not a dump of extraction artifacts. Design the spec schema before implementing extraction.

**Warning signs:**
- Extracted colors are theme index names ("accent1", "dk2") instead of HEX values
- Font sizes differ between extraction and re-generation because layout-level inheritance wasn't resolved
- Extracted layout patterns can't be parameterized into reusable templates — each page spec is a unique snowflake
- Spec is too large for AI context windows because it contains raw XML dumps

**Phase to address:**
Spec Extraction Phase (SPEC-01). The spec schema MUST be defined before extraction code is written.

---

### Pitfall 4: SVG→DrawingML Feature Gap Assumptions

**What goes wrong:**
Assuming the SVG→DrawingML converter handles arbitrary SVG features that are trivial in SVG but impossible in DrawingML. Specific gaps: `<mask>` (no per-pixel alpha in DrawingML — ppt-master explicitly bans it), `<style>` tags and class attributes (DrawingML uses inline attributes), `@font-face` (PowerPoint can't load web fonts at export time), `<foreignObject>` (DrawingML has no HTML embedding), `<symbol>` + `<use>` chains (only direct `<use data-icon>` is supported), `rgba()` color syntax and `<g opacity>` (DrawingML needs fill-opacity/stroke-opacity per element).

**Why it happens:**
SVG is the "easy" part of the pipeline — AI models generate it fluently. Developers test the conversion on simple rects and text and assume feature coverage is complete. The actual banned-feature list in ppt-master grew empirically from PPT export failures, not from reading the OOXML spec. Features that "should" work (like `<mask>`) are theoretically expressible in DrawingML but practically unreliable across PowerPoint versions.

**How to avoid:**
1. Adopt ppt-master's banned-feature blacklist wholesale as the starting constraint set
2. Run `svg_quality_checker.py` (or equivalent) on every generated SVG BEFORE conversion, not after
3. Treat the blacklist as a living document — add new entries whenever a conversion failure occurs
4. Never implement "auto-fix" for banned features — require the AI to re-author the page with the correct substitute (ppt-master's hard-won lesson: auto-fix silently loses design intent)
5. Document substitute-effect routing for each banned feature (e.g., `<mask>` image overlay → stacked `<rect>` with gradient)

**Warning signs:**
- SVG renders fine in browser preview but produces blank shapes in exported PPTX
- Conversion errors at page N+5 when pages 1-5 worked — banned feature crept in mid-deck
- Silent drops: elements in SVG simply don't appear in PPTX with no error message
- `<style>` or `class` in AI-generated SVG (AI defaults to this pattern heavily)

**Phase to address:**
SVG→DrawingML Pipeline Phase (PIPE-01). The blacklist is the contract for AI SVG generation.

---

### Pitfall 5: Multi-Platform Skill Format Divergence (Opencode vs Claude Code)

**What goes wrong:**
Writing a "universal" skill that works identically on both Opencode and Claude Code fails because the two platforms have different tooling models, execution environments, and discoverability mechanisms. Opencode uses `.opencode/skills/` with frontmatter-described skills and has its own tool suite; Claude Code uses `.claude/skills/` with `.claude-plugin/marketplace.json` and the `/plugin install` command. Script paths, file system access patterns, and execution discipline rules differ. A skill built for one platform breaks silently on the other.

**Why it happens:**
Both platforms are evolving rapidly and independently. ppt-master's own history shows this explicitly — issue #67 ("Add OpenCode support") was added as a separate integration path, not as a "it just works" scenario. Key divergences: Claude Code expects `SKILL.md` at `skills/<name>/SKILL.md` with specific frontmatter format; Opencode's skill loader expects skills under `~/.opencode/skills/` with potentially different discovery conventions. Python venv management, dependency installation expectations, and command execution models also differ.

**How to avoid:**
1. Maintain the skill as a Python package with an installable CLI entry point — let the platform-specific SKILL.md files be thin wrappers
2. Abstract all file system paths through a config layer that detects the host environment
3. Implement dual bootstrapping: a single `install.sh` that handles both Opencode and Claude Code conventions
4. Test the full pipeline on both platforms before declaring compatibility
5. Keep the skill's core logic in standalone Python modules (not embedded in SKILL.md) so platform wrappers stay thin

**Warning signs:**
- Hardcoded paths referencing `~/.claude/` or `~/.opencode/` that fail on the other platform
- Skill works in one platform but the other shows "skill not found" even after installation
- Python dependency issues (pip install paths differ between platform environments)
- Platform-specific skill frontmatter that's incompatible (e.g., Claude Code `tools:` field format vs Opencode format)

**Phase to address:**
Platform Compatibility Phase (PLAT-01). Must be designed for from the start, not retrofitted.

---

### Pitfall 6: Forking ppt-master Without Understanding Pipeline Entanglement

**What goes wrong:**
Forking ppt-master and attempting to "strip down" or "repurpose" it fails because the pipeline's components are deeply entangled: the SVG quality checker validates against constraints that the post-processing pipeline depends on; `finalize_svg.py` transforms SVGs in ways that `svg_to_pptx.py`'s memory consumer relies on at conversion time; the icon embedding system expects a specific `data-icon` attribute format; `tspan_flattener` is called both from the disk pipeline AND from the in-memory converter but for different purposes. Changing one module breaks two others silently.

**Why it happens:**
ppt-master's technical design document explicitly warns about this: "Why each artifact and module exists in the engineering conversion stage, and which workflows would break if you delete it." The pipeline grew organically over 2.6+ versions with each module solving a discovered failure mode. The entanglement is documented but easy to overlook when forking. Developers typically assume they can strip the "unnecessary" parts (image generation, narration, templates) and end up with a broken pipeline.

**How to avoid:**
1. Read ALL of ppt-master's `docs/technical-design.md` before writing any code — the "Post-Processing Pipeline" section and "Native PPTX Conversion Internals" section are critical
2. Map the full module dependency graph: which module calls which, from which consumer (disk vs memory), and what would break
3. Fork the FULL codebase first and run end-to-end tests; only THEN begin stripping/adapting
4. Pay special attention to the dual-consumer modules: `embed_icons.py` (disk: `finalize_svg`; memory: `use_expander.py`) and `flatten_tspan.py` (disk: `finalize_svg`; memory: `tspan_flattener.py`)
5. Never remove `svg_finalize/` modules without understanding their in-memory use during PPTX export

**Warning signs:**
- "Simplified" pipeline produces blank icons in exported PPTX (removed `embed_icons.py` without understanding dual use)
- Multi-line text collapses to single line (removed `flatten_tspan.py` — this is the most common breakage)
- Rounded rectangles become sharp-cornered in PowerPoint's internal SVG parser (removed `svg_rect_to_path.py`)
- Preview PPTX has broken images but native PPTX works (or vice versa) — the two export paths use different sources

**Phase to address:**
Pipeline Adaptation Phase (PIPE-01). The module dependency mapping must be done before adaptation begins.

---

### Pitfall 7: Interactive Questioning → Content Exhaustion / Decision Fatigue

**What goes wrong:**
Interactive content gathering ("smart hybrid questioning") either asks too few questions (producing thin, generic slides) or asks too many (the user abandons the workflow). The middle ground is extremely narrow. ppt-master's solution — the "Eight Confirmations" bundled into a single blocking gate — is a hard-won design: each individual confirmation phase proved to cause contradictory decisions and backtracking. Spreading questions across phases invites the user to change their mind mid-generation, requiring full regeneration.

**Why it happens:**
AI coding assistants tend toward "let me ask first" patterns as safety defaults. For PPT generation, most AI tools default to asking 15-20 incremental questions distributed across the workflow. Each question appears reasonable in isolation but the cumulative effect is decision fatigue. The user doesn't know up-front how many questions will be asked, and each question may invalidate prior answers because design choices are correlated (color → icon library → typography → image style).

**How to avoid:**
1. Bundle ALL design confirmations into ONE blocking gate (ppt-master's model: Eight Confirmations)
2. Pre-fill every confirmation with a data-driven recommendation (not blank); user only needs to approve or adjust
3. Never re-ask design questions mid-generation — the design spec is the contract; changes happen in explicit edit workflow afterward
4. For content gap-filling: ask section-level overview first (what sections, what key messages), then gap-fill with targeted questions per section — but cap total questions at 5-8
5. Provide a "skip and generate" escape hatch on every question group

**Warning signs:**
- User stops responding to questions (indicates fatigue threshold crossed)
- User changes earlier answers when asked later questions (indicates bundling is needed)
- More than 3 rounds of Q&A before generation begins
- Generation output doesn't match user's last answer because they're confused about which answer was the "final" one

**Phase to address:**
Content Generation Phase (GEN-01, GEN-02). The questioning UX design must be locked before implementation.

---

### Pitfall 8: Empty/Default spec_lock.md Silently Produces Uniform Card-Grid Layouts

**What goes wrong:**
When the `spec_lock.md` has empty entries for `page_rhythm`, `page_layouts`, or `page_charts`, the AI Executor defaults to producing uniform card-grid layouts for every slide — regardless of content needs. The result is a visually monotonous deck where every page looks the same. This is a documented anti-drift mechanism failure mode in ppt-master's executor-base.md §2.1: empty entries are intentional Strategist signals meaning "design freely for this page," not "use the default template." The AI ignores this signal and reaches for the safest default.

**Why it happens:**
AI models under context-compression pressure revert to the lowest-risk pattern. A card grid is a safe, predictable layout that rarely produces visual errors. The Strategist's signal of "empty = design freely" demands more creative reasoning per page, which the Executor avoids when context is stretched (especially on decks >10 slides). The spec_lock re-read discipline partially mitigates but doesn't eliminate the problem.

**How to avoid:**
1. Never have truly empty `page_layouts` — provide at least 2-3 distinct layout templates for the Executor to choose from, even in "free design" mode
2. In the spec extraction pipeline, classify reference slides into layout archetypes and parameterize them, making them available as "spec-derived layouts" in the output spec
3. Add a quality check that detects card-grid uniformity: if >60% of pages share the same layout structure, flag for diversity review
4. Make the spec-to-layout binding explicit: every page in `spec_lock.md` should reference a specific layout template or an explicit "free-design with these constraints" instruction

**Warning signs:**
- Every slide after slide 3 has an image-left/text-right or image-right/text-left structure
- No page uses any layout variant mentioned in the design spec
- Decks >8 slides show progressive layout uniformity (early slides vary, later slides converge)
- `page_rhythm` field is empty or "default" on every page

**Phase to address:**
Content Generation Phase (GEN-02). The spec_lock schema must have explicit layout binding per page.

---

## Technical Debt Patterns

| Shortcut | Immediate Benefit | Long-term Cost | When Acceptable |
|----------|-------------------|----------------|-----------------|
| Skip OOXML-level parsing, use python-pptx high-level API for all extraction | Faster initial implementation | Background/font/theme extraction silently wrong; re-extraction needed with raw lxml | Only for prototyping, never in production spec extraction |
| Hard-code HEX color palette instead of resolving theme references | No theme resolution logic needed | Specs break on any reference PPTX not using default Office theme | Never — theme resolution is table-stakes |
| Use `<mask>` in AI-generated SVGs (works in browser preview) | AI generates visually correct previews | Silent element drops in exported PPTX; manual re-authoring per page | Never |
| Skip `svg_quality_checker.py` during development | Faster iteration | Banned features creep in; discovered only at export time when they're hard to find | Only in the very first draft of a single page |
| Hard-code `/Users/wuliang/.claude/skills/` paths | Works on developer's machine | Fails on Opencode installations, Linux file systems, any non-macOS setup | Never — path resolution must be environment-aware from day 1 |
| Delete `svg_finalize/embed_icons.py` because it "seems unnecessary" | Smaller codebase | All icons silently dropped from native PPTX output | Never |
| Merge Eight Confirmations into sequential questions | Simpler to implement | Decision fatigue, contradictory answers, restart loops | Never |

## Integration Gotchas

| Integration | Common Mistake | Correct Approach |
|-------------|----------------|------------------|
| python-pptx for PPTX analysis | Using high-level API for everything | Use `lxml` with OOXML namespaces for background, theme, font resolution; python-pptx only for shape tree iteration |
| `Slide._element` for XML access | Assuming it returns the expected node type | Always inspect the actual tag name (`element.tag`); `background._element` returns `<p:cSld>` not `<p:bg>` |
| SVG quality checker pre-processing | Running after `finalize_svg.py` | Must run on raw `svg_output/` BEFORE finalize; finalize rewrites SVG and masks source violations |
| `finalize_svg.py` script | Substituting with `cp` or manual copying | `finalize_svg.py` performs icon embedding, image crop, text flattening, rounded rect conversion — none of which `cp` does |
| Icon embedding in SVGs | Using bare `<path>` elements or missing `data-icon` attribute | Must use `<use data-icon="library/name" ...>` with library prefix; post-processing replaces these with inline paths |
| PPTX template import via `pptx_template_import.py` | Expecting it to handle all PPTX files equally | The importer reads OOXML directly; heavily customized/decorative files need `fidelity` mode with literal visual fidelity |
| Opencode skill installation | Copying Claude Code skill format verbatim | Use Opencode's `~/.opencode/skills/` convention with OpenCode-compatible frontmatter; test skill discovery on both platforms |

## Performance Traps

| Trap | Symptoms | Prevention | When It Breaks |
|------|----------|------------|----------------|
| AI SVG generation batching | Visual style drift across slides — colors and fonts mutate mid-deck | One page at a time, sequential generation with per-page `spec_lock.md` re-read | Decks > 8 slides |
| SVG file size bloat from Base64-embedded images | 50MB+ SVG per page; preview PPTX import timeout | External image references in `svg_output/`; Base64 inline only in `svg_final/` for preview | Decks with > 5 full-bleed images |
| XML entity parsing on large OOXML files | Spec extraction hangs or OOM on 100+ slide reference files | Stream parsing with `lxml.iterparse`; process slides one at a time | Reference PPTX > 50 slides |
| Context window exhaustion from nested template reads | AI loses earlier pages' design context when reading too many template/chart files | Batch-read ALL layout/chart templates ONCE up front before any SVG generation | Decks with > 5 distinct layout templates |
| python-pptx `add_picture()` on large images | 5+ second delays per image; file size balloons | Resize images to slide-appropriate resolution before embedding; use `add_picture()` with width/height constraints | Photos > 5MB or > 4K resolution |

## Security Mistakes

| Mistake | Risk | Prevention |
|---------|------|------------|
| Embedding user-provided PPTX content into `.env` or log output | Leaking proprietary presentation content to console logs | Strip file content from logging; only log metadata (slide count, dimensions) |
| Hardcoding API keys for image generation in SKILL.md | Keys committed to public repo | Use `.env` with `.env.example` template; detect environment variable and prompt user to set if missing |
| Extracting and storing all text from reference PPTX without sanitization | Sensitive corporate data persisted in spec files in plain text | Treat spec files as containing confidential data; `.gitignore` the `specs/` directory if storing user-specific specs |
| Using `exec()` or `eval()` for dynamic layout parameterization | Arbitrary code execution from spec file content | Use structured data formats (JSON/YAML) for layout parameters; never evaluate spec content as code |

## UX Pitfalls

| Pitfall | User Impact | Better Approach |
|---------|-------------|-----------------|
| Asking the user to choose between 22 layout templates | Choice paralysis; user picks random template, gets poor results | AI recommends 1-3 top matches based on content analysis, presents with rationale, user confirms or cycles |
| Eight Confirmation values presented as blank fields for user to fill | User doesn't know what reasonable values look like | Every confirmation has a data-driven pre-filled recommendation; user only needs to approve or tweak |
| "Smart hybrid questioning" asks everything up front | Long question sequence feels like an interrogation before any value is delivered | Section-level overview first (2-3 questions), then gap-fill per section (1-2 targeted questions each), capped at 8 total |
| Spec extraction requires the user to understand PPTX internals | Non-technical users can't use the spec extraction feature | Automatically detect and extract all spec dimensions; only ask user to confirm which aspects to carry forward to generation |
| Generated PPTX has ugly slides with no explanation of why | User assumes the tool is broken | Surface the "design draft" philosophy clearly: generated PPTX is a 90% starting point, not a finished product; minor manual editing is expected and the UI should say so |
| No preview before final export | User waits minutes for generation only to discover layout issues | Provide intermediate SVG preview (open in browser) after 2-3 pages are generated; user can stop and fix early |

## "Looks Done But Isn't" Checklist

- [ ] **Spec extraction:** Often missing font weight resolution — verify bold/normal classification works on inherited styles, not just explicit `<a:rPr b="1">` attributes
- [ ] **Spec extraction:** Often missing gradient stops — verify multi-color gradients (not just single-color) are extracted with correct stop positions
- [ ] **SVG→DrawingML:** Often missing `stroke-dasharray` conversion — verify dashed/dotted borders render correctly in exported PPTX
- [ ] **SVG→DrawingML:** Often missing `text-decoration` (underline/strikethrough) — verify decorated text passes through conversion intact
- [ ] **SVG→DrawingML:** Often missing `transform="rotate(...)"` conversion — verify rotated text/elements preserve orientation in PPTX
- [ ] **Multi-platform:** Often missing Python venv activation differences between platforms — verify skill works when Python is in PATH but venv is not activated
- [ ] **Multi-platform:** Often missing file permission handling on macOS (quarantine xattrs) — verify scripts execute after `pip install` without `xattr -d com.apple.quarantine`
- [ ] **Skill format:** Often missing the XML character encoding trap — verify `&mdash;` styled named entities are caught before SVG export (they're XML-illegal in standalone SVG)
- [ ] **Content generation:** Often missing pagination for long content — verify 15+ content sections don't produce a 55-slide deck (needs merge/split logic)
- [ ] **Forking:** Often missing the `preserveAspectRatio` → `srcRect` conversion that ppt-master does — verify cropped images aren't distorted in exported PPTX

## Recovery Strategies

| Pitfall | Recovery Cost | Recovery Steps |
|---------|---------------|----------------|
| Background property gap (#1) | LOW | Switch to lxml direct parsing; 1-2 files changed; no pipeline impact |
| tspan fragmentation (#2) | MEDIUM | Fix requires re-authoring SVGs with corrected tspan discipline; 1-2 hours for a 15-page deck |
| Spec extraction gap (#3) | HIGH | Rewrite entire extraction pipeline with resolved-theme approach; 3-5 days if caught late |
| SVG feature gap (#4) | MEDIUM | Add to blacklist, regenerate affected pages; 30 min per instance |
| Platform divergence (#5) | MEDIUM | Rewrite path resolution layer, add bootstrapping script; 1-2 days |
| Pipeline entanglement (#6) | HIGH | Restore deleted modules, retest full pipeline; 2-5 days if caught late |
| Interactive UX fatigue (#7) | LOW | Restructure questioning flow; design change only, no code rewrite |
| Spec_lock empty defaults (#8) | LOW | Add explicit layout binding per page; spec schema change |

## Pitfall-to-Phase Mapping

| Pitfall | Prevention Phase | Verification |
|---------|------------------|--------------|
| PPTX backgrounds API trap (#1) | SPEC-01 (Spec Extraction) | Test extraction on PPTX with non-white, gradient, and picture backgrounds. Verify HEX resolved not theme index. |
| tspan fragmentation (#2) | PIPE-01 (SVG→DrawingML Pipeline) | Run tspan_flattener on generated SVGs; verify single-line text is NOT split into multiple frames in exported PPTX |
| Spec extraction gap (#3) | SPEC-01 (Spec Extraction) | Re-extract specs from the generated PPTX; compare to original reference PPTX; colors and fonts must match |
| SVG feature gap (#4) | PIPE-01 (SVG→DrawingML Pipeline) | svg_quality_checker.py must pass with 0 errors on every generated SVG |
| Multi-platform divergence (#5) | PLAT-01 (Platform Compatibility) | Run full pipeline on both Opencode and Claude Code; skill must be discoverable and executable on both |
| Forking entanglement (#6) | PIPE-01 (Pipeline Adaptation) | Delete any "unnecessary" module and verify full pipeline still works; if it breaks, module was necessary |
| Interactive UX fatigue (#7) | GEN-01 (Content Generation) | User test: count total questions before generation begins; must be ≤ 8; count re-ask occurrences; must be 0 |
| Empty spec_lock defaults (#8) | GEN-02 (Content Generation) | Audit spec_lock for empty page_layouts entries; every page must have explicit layout reference or constrained free-design instruction |

## Sources

- **python-pptx official docs:** https://python-pptx.readthedocs.io/ — shape types, chart API, known limitations (HIGH confidence)
- **python-pptx GitHub issues:** #1126 (Slide.background._element bug), #1111 (Font.color getter mutates XML), #1112 (SVG image not supported) — background gap, font mutation, SVG support gaps (HIGH confidence)
- **ppt-master GitHub issues:** #90 (generated PPT format error/can't open), #61 (layout aesthetics), #60 (fine-detailed ppt markdown), #62 (HTML entities in SVG), #74 (claude skill capability incomplete), #67 (OpenCode support added) — real-world failure modes (HIGH confidence)
- **ppt-master Technical Design:** https://github.com/hugohe3/ppt-master — pipeline architecture, SVG constraints, module entanglement, anti-drift mechanism, dual-consumer modules, banned feature blacklist (HIGH confidence)
- **ppt-master shared-standards.md:** SVG banned features, tspan discipline, line-end marker constraints, clip-path rules, post-processing pipeline rules (HIGH confidence)
- **ppt-master SKILL.md:** Execution discipline rules, role switching protocol, spec_lock re-read requirement, Eight Confirmations bundling (HIGH confidence)
- **ppt-master FAQ.md:** Model recommendations, layout issue troubleshooting, template derivation, image acquisition paths (MEDIUM confidence)
- **Opencode official site:** https://opencode.ai — skill ecosystem conventions (MEDIUM confidence)

---
*Pitfalls research for: AI PPT Generation Skill*
*Researched: 2026-05-06*
