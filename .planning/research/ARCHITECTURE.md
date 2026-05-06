# Architecture Research

**Domain:** AI-powered PPT generation skill with spec extraction
**Researched:** 2026-05-06
**Confidence:** HIGH

> Sources: ppt-master v2.6.0 technical design docs, SKILL.md, source tree analysis, PROJECT.md requirements. Architecture verified against actual working pipeline in production use.

## System Overview

```
┌──────────────────────────────────────────────────────────────────────────┐
│                         SKILL DEFINITION LAYER                            │
│  ┌───────────────┐  ┌──────────────────┐  ┌─────────────────────────┐   │
│  │   SKILL.md    │  │  references/     │  │  scripts/ (Python)      │   │
│  │  (workflow)   │  │  (role definitions │  │  (pipeline executables) │   │
│  │               │  │   + constraints)  │  │                         │   │
│  └───────┬───────┘  └────────┬─────────┘  └───────────┬─────────────┘   │
│          │                   │                        │                  │
├──────────┴───────────────────┴────────────────────────┴──────────────────┤
│                        ORCHESTRATION LAYER                                │
│  The AI agent follows SKILL.md workflow, dispatches to scripts, manages   │
│  state transitions between phases. Roles: Spec Extractor → Interviewer →  │
│  Outline Builder → PPTX Generator.                                        │
├──────────────────────────────────────────────────────────────────────────┐
│ ┌─────────────────────┐    ┌──────────────────┐    ┌──────────────────┐ │
│ │  SPEC EXTRACTION    │    │  CONTENT GATHERING│    │  PPTX GENERATION │ │
│ │                     │    │                   │    │                  │ │
│ │ pptx → design_spec  │    │ Smart Hybrid Q&A  │    │ Outline→SVG→PPTX │ │
│ │ extract_colors()    │    │ section_overview()│    │ generate_slides()│ │
│ │ extract_fonts()     │    │ gap_fill_section()│    │ svg_to_drawingml │ │
│ │ extract_layouts()   │    │                   │    │                  │ │
│ │ extract_logic()     │    │                   │    │                  │ │
│ └─────────┬───────────┘    └────────┬─────────┘    └────────┬─────────┘ │
│           │                         │                       │           │
├───────────┴─────────────────────────┴───────────────────────┴───────────┤
│                           DATA STORAGE LAYER                              │
│                                                                           │
│  ┌──────────────────┐ ┌──────────────────┐ ┌─────────────────────────┐   │
│  │ specs/           │ │ project/         │ │ ~/.ppt-skill/           │   │
│  │ (project-local   │ │ (per-generation  │ │ (global config)         │   │
│  │  spec files)     │ │  state)          │ │                         │   │
│  └──────────────────┘ └──────────────────┘ └─────────────────────────┘   │
└──────────────────────────────────────────────────────────────────────────┘
```

## Component Boundaries

### 1. Skill Definition Layer

| Component | Responsibility | Communicates With |
|-----------|---------------|-------------------|
| `SKILL.md` | Entry point, workflow definition, role-switching rules, serial pipeline enforcement | Orchestration Layer (AI agent reads and follows it) |
| `references/` | Role-specific instructions (spec-extractor, interviewer, outline-builder, pptx-generator), shared constraints (SVG limits, format specs) | AI agent loads per role |
| `scripts/` | Python executables for deterministic pipeline steps (PPTX analysis, SVG→DrawingML conversion, quality checking) | Orchestration Layer (invoked via bash) |

**Key design decision — one SKILL.md with role switching, not separate skill files.** ppt-master's model of "one agent, multiple roles, loaded on demand" is proven. Each role gets its own reference file loaded at the start of its phase. This avoids prompt contamination (Strategist's conversational mode vs Executor's strict-XML mode) while keeping a single agent with full context continuity.

**Skill directory structure:**

```
ppt-skill/
├── SKILL.md                          # Workflow entry point + global rules
├── requirements.txt                  # Python dependencies (python-pptx, lxml, PyMuPDF, etc.)
├── references/
│   ├── spec-extractor.md             # Role: analyze PPTX, extract design spec
│   ├── interviewer.md                # Role: smart hybrid questioning flow
│   ├── outline-builder.md            # Role: generate content outline from gathered info
│   ├── pptx-generator.md             # Role: execute SVG→DrawingML pipeline
│   ├── shared-standards.md           # SVG/PPT technical constraints (adapted from ppt-master)
│   ├── canvas-formats.md             # Supported output formats
│   └── spec-format.md                # Spec file schema documentation
├── scripts/
│   ├── analyze_pptx.py               # PPTX analysis: extract colors, fonts, layouts, logic
│   ├── spec_manager.py               # Save/load/list/validate spec files
│   ├── spec_to_lock.py               # Convert design_spec.md → spec_lock.md
│   ├── svg_to_pptx.py                # SVG → DrawingML PPTX converter (adapted from ppt-master)
│   ├── svg_quality_checker.py        # SVG constraint validator (adapted from ppt-master)
│   ├── finalize_svg.py               # SVG post-processing (adapted from ppt-master)
│   ├── total_md_split.py             # Speaker notes splitting (adapted from ppt-master)
│   ├── update_spec.py                # Propagate spec changes across SVGs (adapted from ppt-master)
│   ├── svg_finalize/                 # SVG post-processing modules (adapted from ppt-master)
│   └── svg_to_pptx/                  # DrawingML conversion modules (adapted from ppt-master)
└── templates/
    └── spec_template.md              # Template for extracted design_spec.md
```

### 2. Spec Extraction Component

| Sub-component | Responsibility | Implementation |
|---------------|----------------|----------------|
| `extract_colors` | Parse theme colors, accent colors, background/fill colors from PPTX XML | python-pptx + lxml OOXML traversal |
| `extract_fonts` | Extract font families, sizes, weights, styles from slide masters and layouts | python-pptx font analysis + OOXML fallback |
| `extract_layouts` | Identify layout patterns (title slides, content slides, section dividers), placeholder positions, element hierarchies | Slide layout analysis, spatial clustering |
| `extract_logic` | Infer presentation logic: argument flow, section structure, storytelling patterns from slide sequence and content | LLM-assisted analysis of slide content flow |
| `build_spec` | Assemble all extracted data into a structured design_spec.md | Merge extraction results into canonical format |

**Architecture:** This component reads PPTX files directly via python-pptx and lxml. It does NOT use ppt-master's pipeline — it's a net-new capability. The output is a `design_spec.md` file that follows the same schema ppt-master uses, enabling downstream reuse of its generation pipeline.

**Why python-pptx + lxml (not python-pptx alone):** python-pptx abstracts the XML layer for common operations but doesn't expose theme XML, font scheme details, or layout geometry at the fidelity needed for complete spec extraction. lxml provides direct XPath access to the underlying OOXML where needed.

### 3. Spec Management Component

| Operation | Responsibility |
|-----------|---------------|
| `spec_manager.py save` | Save extracted spec to `specs/<name>.md` with versioning |
| `spec_manager.py load` | Load a spec by name, validate structure |
| `spec_manager.py list` | List available specs with metadata (source, date, slide count) |
| `spec_manager.py validate` | Validate spec file conforms to schema |

**Storage:** Project-local `specs/` directory. Each spec is a Markdown file with YAML frontmatter (name, source file, extraction date, slide count) followed by the design_spec.md content. This is filesystem-native — no database, no config file. The AI agent can list them with `ls`, read them with `read_file`, and the user can browse them in their IDE.

**Why project-local, not global:** Aligns with the constraint that this is a single-user generation workflow. Specs are tied to the project they were extracted for. Sharing requires copying files, which is simple and transparent. A global spec library could be added later but adds complexity without proven demand.

### 4. Content Gathering Component (Smart Hybrid Questioning)

| Phase | Responsibility | Trigger |
|-------|---------------|---------|
| **Section overview** | Ask the user: what are the main sections of this presentation? Get title + brief description per section. | Always — this is the starting point |
| **Gap analysis** | Compare user's input against spec requirements and presentation logic model. Identify missing information. | After section overview |
| **Targeted filling** | For each gap, ask exactly one targeted question per round. Never ask "what else?" — always be specific. | Per-section, iteratively |
| **Completion check** | Determine when enough information exists to build a quality outline. | After each gap-fill round |

**Architectural fit:** This is an AI-agent-driven conversational flow, not a script. The `references/interviewer.md` role file defines the questioning strategy, pacing, and completion criteria. The AI agent executes it in conversation with the user. No Python script runs this phase — it's pure agent behavior guided by role instructions.

**Why hybrid (not pure agent, not pure script):** The AI agent handles the conversational nuance (understanding user intent, adapting questions), while the role file provides structure (don't skip phases, don't ask vague questions, enforce completion criteria). A rigid script-based form would miss context; a pure agent without guidance would drift.

### 5. Outline Generation Component

| Sub-component | Responsibility |
|---------------|----------------|
| `build_outline` | From gathered content + loaded spec, produce a slide-by-slide outline with title, content summary, and layout assignment per slide |
| `validate_outline` | Check outline against spec constraints (slide count, layout availability, content fitting) |
| `approval_gate` | Present outline to user for approval/revision before generation |

**Architectural fit:** Like Content Gathering, this is AI-agent-driven with role guidance from `references/outline-builder.md`. The output is a structured outline (Markdown table or JSON) that the PPTX Generator component consumes.

### 6. PPTX Generation Component

| Sub-component | Responsibility | Source |
|---------------|----------------|--------|
| `spec_lock_generation` | Convert approved outline + design spec into machine-readable `spec_lock.md` | New: `spec_to_lock.py` |
| `SVG generation` | AI generates one SVG per slide following spec_lock constraints | AI agent (Executor role) |
| `SVG quality check` | Validate SVGs for banned features, spec compliance, viewBox correctness | Adapted: `svg_quality_checker.py` |
| `SVG finalize` | Post-processing: icon embedding, image inlining, text flattening, rounded rect fix | Adapted: `finalize_svg.py` |
| `PPTX export` | SVG → DrawingML conversion, speaker notes embedding, animation support | Adapted: `svg_to_pptx.py` |

**Integration with ppt-master:** We fork and adapt (not wrap, not reference). The SVG→DrawingML pipeline (`svg_finalize/` + `svg_to_pptx/` modules) is the high-value technical asset. We vendor it into our skill's `scripts/` directory as adapted copies with these modifications:

1. **Remove ppt-master-specific assumptions:** Source-to-MD converters, project_manager.py, template system, image acquisition (out of scope)
2. **Keep:** SVG→DrawingML conversion (17 modules), quality checker, SVG finalize, slide splitting, PPTX export
3. **Add:** Spec-driven generation path — instead of the AI inventing design from scratch, it reads `spec_lock.md` (built from extracted spec + user content) and generates SVGs that conform

**Why fork, not wrap:** Wrapping would require ppt-master as a dependency, forcing users to install its full dependency tree and follow its project structure conventions. Forking lets us strip down to the essential pipeline (~10 modules) and adapt the interfaces to our spec-driven workflow. The adaptation surface is narrow — primarily the Executor's design parameter source changes from "AI-invented" to "spec_lock.md-extracted".

## Data Flow

### Primary Flow: Spec Extraction → Storage

```
User provides PPTX file
         ↓
[analyze_pptx.py]  ← reads PPTX via python-pptx + lxml
         ↓
    ┌────┴────┬─────────┬──────────┐
    ↓         ↓         ↓          ↓
 Colors    Fonts    Layouts    Logic Model
    └────┬────┴─────────┴──────────┘
         ↓
[spec_manager.py save]  ← assembles into design_spec.md + frontmatter
         ↓
specs/<spec_name>.md  ← stored in project-local specs/ directory
```

### Primary Flow: Content → Generation

```
User initiates PPT generation ("create PPT from spec X")
         ↓
[spec_manager.py load <spec_name>]  ← AI agent loads spec into context
         ↓
[Interviewer Role]  ← smart hybrid questioning
    ├── Section overview (user describes sections)
    ├── Gap analysis (AI identifies missing info per section)
    └── Targeted filling (AI asks specific gap-filling questions)
         ↓
[Outline Builder Role]  ← AI generates slide-by-slide outline
         ↓
⛔ BLOCKING: User approves/modifies outline
         ↓
[spec_to_lock.py]  ← converts outline + spec → spec_lock.md
         ↓
[PPTX Generator Role]  ← AI generates SVGs pag eby page
    ├── Per page: read_file spec_lock.md (anti-drift)
    ├── Per page: write SVG to svg_output/
    └── All pages: run svg_quality_checker.py
         ↓
[Post-processing]  ← scripts run sequentially
    ├── total_md_split.py
    ├── finalize_svg.py
    └── svg_to_pptx.py
         ↓
exports/<name>_<timestamp>.pptx  ← native editable PPTX
```

### State Transitions

```
IDLE → SPEC_EXTRACTING → SPEC_STORED
                                    ↓
                              WAITING_USER (user selects spec)
                                    ↓
                              CONTENT_GATHERING
                                    ↓
                              OUTLINE_BUILDING
                                    ↓
                              WAITING_APPROVAL (⛔ BLOCKING)
                                    ↓
                              SVG_GENERATING
                                    ↓
                              POST_PROCESSING
                                    ↓
                              DONE (PPTX exported)
```

## Patterns to Follow

### Pattern 1: Role-Specialized Reference Loading

**What:** Each phase loads only its own role reference file, never all references at once. The AI agent switches modes by reading the new role file explicitly.

**When:** At every phase boundary (Spec Extraction → Content Gathering → Outline Building → PPTX Generation).

**Trade-offs:** Increases total file reads but prevents prompt contamination between modes. The conversation transcript shows a clear audit trail (`[Role Switch: Spec Extractor]`).

**Why this pattern for our skill:** ppt-master validated this approach. Spec extraction ("analyst mode") vs interviewing ("conversational mode") vs generation ("strict execution mode") are incompatible disciplines for a single prompt. Each role loads exactly the instructions it needs.

### Pattern 2: spec_lock.md as Anti-Drift Mechanism

**What:** A machine-readable execution contract (spec_lock.md) that the AI re-reads before generating each slide. Prevents color/font drift across long decks.

**When:** Every slide in the PPTX generation phase.

**Trade-offs:** Adds a file read per slide (low cost). Requires disciplined enforcement via SKILL.md rules.

**Why this pattern for our skill:** Directly adopted from ppt-master. The spec-driven nature of our system makes this even more critical — we're reproducing an extracted design, so drift from the spec is a first-class failure mode.

### Pattern 3: Blocking Gate → Auto-Proceed Pipeline

**What:** The pipeline has exactly one blocking gate (outline approval). Before the gate, the AI must wait for user confirmation and cannot make decisions on the user's behalf. After the gate, the pipeline runs to completion without further interrupts.

**When:** Interactive phases (content gathering, outline review) require user confirmation. Generation (SVG production, post-processing) is auto-proceeding.

**Trade-offs:** Fewer interrupts means faster generation and fewer "what should I do next" prompts, but places more responsibility on the blocking gate to be the right decision point. If the outline is wrong, the generated PPTX is wrong.

**Why this pattern for our skill:** The Eight Confirmations in ppt-master's Strategist phase is the equivalent gate. We move it to outline approval because our content gathering phase already handles spec selection and design decisions — the outline is the natural "this is what you'll get" confirmation point.

### Pattern 4: Scripts for Deterministic Work, AI for Creative Work

**What:** Everything that can be deterministic (PPTX parsing, SVG conversion, quality checking, PPTX export) is implemented as Python scripts. Everything that requires creativity or user interaction (design spec interpretation, question formulation, SVG page design) stays with the AI agent.

**When:** Apply at component design time. Ask: "Can this be done deterministically?" If yes → script. If no → AI agent with role guidance.

**Trade-offs:** Scripts are faster, more reliable, and cheaper than AI model calls. But some tasks (interpreting a user's vague content description, deciding which layout fits a specific slide's content type) genuinely need AI reasoning.

## Anti-Patterns to Avoid

### Anti-Pattern 1: Flat Image PPTX

**What:** Generating slides as full-page images and embedding them as pictures in PPTX.

**Why bad:** Destroys editability — text can't be selected, colors can't be changed, shapes can't be repositioned. The spec specifically requires "natively editable."

**Instead:** Use the SVG→DrawingML pipeline. Every shape, text box, and chart must be a real PowerPoint DrawingML object. This is the core technical differentiator.

### Anti-Pattern 2: Single-Pass Generation

**What:** One API call to generate the entire PPTX in one shot.

**Why bad:** LLM context compression causes progressive drift in colors, fonts, and layout quality across pages. Long decks become visually inconsistent.

**Instead:** Serial page-by-page generation with spec_lock.md re-read before each page. Quality checker catches issues immediately rather than after all pages are generated.

### Anti-Pattern 3: Mega-Prompt Monolith

**What:** Loading all role instructions into one giant system prompt.

**Why bad:** Mode contamination — the analytical discipline needed for spec extraction conflicts with the conversational mode needed for content gathering, which conflicts with the strict-XML mode needed for SVG generation.

**Instead:** Per-role reference files loaded only when that role activates. The AI agent switches modes explicitly.

### Anti-Pattern 4: Spec by Inference

**What:** Having the AI "look at" or "describe" a reference PPTX rather than programmatically extracting its design properties.

**Why bad:** AI visual analysis of PPTX files is unreliable for exact color values, font names, and layout measurements. Subjective descriptions ("warm colors", "clean layout") don't provide the precise HEX codes and pixel positions the generator needs.

**Instead:** `analyze_pptx.py` uses python-pptx + lxml to extract exact OOXML values. The AI only interprets the presentation logic (argument flow, storytelling patterns), which genuinely requires reasoning.

## How Dual Platform Support Affects Architecture

Opencode and Claude Code have different skill conventions but share fundamentals. The architecture handles this through:

### Common Ground
- **Markdown SKILL.md:** Both platforms support installing skills from a Markdown file with YAML frontmatter
- **Python runtime:** Both platforms can execute Python scripts via bash commands
- **POSIX file operations:** Both platforms use the same file read/write/bash tools

### Platform Differences Handled

| Concern | Opencode | Claude Code | Architecture Choice |
|---------|----------|-------------|---------------------|
| Skill install path | `~/.opencode/skills/` | `~/.claude/skills/` | Use skill marketplace format; detect install path at runtime via `${SKILL_DIR}` |
| Frontmatter format | `name:`, `description:` | `name:`, `description:`, `type: skill` | Include both conventions; unused fields are ignored |
| Shell command style | `bash` tool | `bash` tool | Same — no adaptation needed |
| File operation tools | `Read`, `Write`, `Edit`, `Grep` | `Read`, `Write`, `Edit`, `Grep` | Same — no adaptation needed |
| Skill invocation | `/skill-name` command | `/skill-name` command or `@skill-name` | Same naming convention |

### Architecture Impact

**None on core components.** The Python scripts, role reference files, and data flow are platform-agnostic — they run in any environment with Python 3.10+ and a file system. The AI agent (whether Opencode or Claude Code) follows the same SKILL.md workflow.

**One adaption needed:** `${SKILL_DIR}` resolution. ppt-master uses `${SKILL_DIR}` to reference its own scripts regardless of install location. We adopt the same convention — the SKILL.md frontmatter or the AI agent's platform resolves `${SKILL_DIR}` to the skill's install path at runtime.

## Project Structure

```
ppt-skill/
├── SKILL.md                          # Workflow entry point (platform-agnostic)
├── requirements.txt                  # python-pptx, lxml, Pillow, PyMuPDF
│
├── references/                       # Role definitions (loaded per phase)
│   ├── spec-extractor.md             # PPTX analysis role
│   ├── interviewer.md                # Content gathering role
│   ├── outline-builder.md            # Outline generation role
│   ├── pptx-generator.md             # SVG→PPTX execution role (adapted from executor-base.md)
│   ├── shared-standards.md           # SVG/PPT technical constraints (adapted)
│   ├── canvas-formats.md             # Output format specs
│   └── spec-format.md                # Spec file schema
│
├── scripts/                          # Deterministic Python executables
│   ├── analyze_pptx.py               # NEW: PPTX → design spec extraction
│   ├── spec_manager.py               # NEW: spec save/load/list/validate
│   ├── spec_to_lock.py               # NEW: design_spec.md → spec_lock.md
│   ├── svg_to_pptx.py                # ADAPTED: SVG → DrawingML converter
│   ├── svg_quality_checker.py        # ADAPTED: SVG constraint validator
│   ├── finalize_svg.py               # ADAPTED: SVG post-processing
│   ├── total_md_split.py             # ADAPTED: speaker notes splitting
│   ├── update_spec.py                # ADAPTED: spec propagation across SVGs
│   ├── svg_finalize/                 # ADAPTED: post-processing modules
│   │   ├── embed_icons.py
│   │   ├── flatten_tspan.py
│   │   ├── align_embed_images.py
│   │   ├── crop_images.py
│   │   ├── embed_images.py
│   │   ├── fix_image_aspect.py
│   │   └── svg_rect_to_path.py
│   └── svg_to_pptx/                  # ADAPTED: DrawingML conversion modules
│       ├── use_expander.py
│       ├── tspan_flattener.py
│       └── ... (other converter modules)
│
└── templates/                        # Templates (none from ppt-master)
    └── spec_template.md              # Template for extracted design_spec.md
```

### Structure Rationale

- **`references/` is flat (not nested):** Each file is loaded independently by role. Nesting would create false hierarchies — these are peer role definitions, not sub-aspects of a parent concept.
- **`scripts/` mirrors ppt-master's structure where adapted:** Makes diffs against upstream easier. New scripts (`analyze_pptx.py`, `spec_manager.py`, `spec_to_lock.py`) live alongside adapted scripts at the top level.
- **No `projects/`, `exports/`, `backup/` directories in the skill repo:** These are runtime artifacts created in the user's working directory, not in the skill's install directory. The skill operates on the user's project.
- **No image acquisition scripts:** Out of scope per requirements. Generation uses existing images or text-only slides.
- **No template layout SVGs from ppt-master:** Our templates come from extracted specs, not predefined layout libraries.

## Build Order Implications

The component dependencies drive a natural build order:

```
Phase 1: Fork + Adaptation
  ├── Fork ppt-master, extract svg_finalize/ + svg_to_pptx/ modules
  ├── Adapt scripts to remove ppt-master project structure assumptions
  └── Verify: the pipeline converts sample SVGs to PPTX standalone

Phase 2: Spec Extraction (depends on: nothing)
  ├── Build analyze_pptx.py (python-pptx + lxml)
  ├── Build spec_manager.py (save/load/list/validate)
  ├── Define spec format schema (spec-format.md)
  └── Verify: extracting a sample PPTX produces a valid, complete spec

Phase 3: Spec-Driven Generation (depends on: Phase 1 + Phase 2)
  ├── Build spec_to_lock.py (converts design_spec → spec_lock)
  ├── Write pptx-generator.md role (spec-locked Executor variant)
  ├── Integrate: spec locked values feed into SVG generation
  └── Verify: generating from a spec produces a PPTX matching the spec's design

Phase 4: Content Gathering (depends on: nothing, parallel with Phase 2-3)
  ├── Write interviewer.md role (smart hybrid questioning strategy)
  ├── Write outline-builder.md role (outline generation rules)
  ├── Define outline format
  └── Verify: Q&A flow produces a complete, spec-compatible outline

Phase 5: Full Integration + SKILL.md (depends on: Phase 2 + Phase 3 + Phase 4)
  ├── Write SKILL.md (full workflow, role switching, serial pipeline rules)
  ├── Wire all phases together
  ├── Dual-platform testing (Opencode + Claude Code)
  └── End-to-end verification: PPTX in → spec extracted → content gathered → PPTX out
```

**Dependency graph:**

```
Phase 2 (Spec Extraction) ──┐
                             ├──→ Phase 5 (Integration + SKILL.md)
Phase 1 (Fork Pipeline) ──┬──┘
                          │
Phase 3 (Spec Gen) ───────┘
                          │
Phase 4 (Content) ────────┘
```

Phase 1 and Phase 4 have no dependencies and can run in parallel. Phase 2 can also parallel with Phase 4. Phase 3 depends on Phase 1 (needs the pipeline) and Phase 2 (needs spec format). Phase 5 integrates everything.

## Scaling Considerations

This is a single-user CLI skill, not a web service. Scaling concerns are minimal:

| Scale | Architecture Adjustment |
|-------|------------------------|
| 1 user, <50 slide decks | Monolithic SKILL.md + scripts. No splitting needed. |
| Multiple specs stored | `specs/` directory scales to hundreds of files without issue. `spec_manager.py list` may need pagination if >1000 specs. |
| Very long decks (>100 slides) | spec_lock.md re-read per page keeps context fresh. Quality checker catches issues page-by-page. No architecture change needed. |

**First bottleneck:** Context window during SVG generation for long decks (>60 pages). Mitigation: spec_lock.md re-read already handles color/font drift. Page-by-page generation keeps per-slide context bounded. If deeper issues emerge, split generation across multiple AI sessions with spec_lock.md as the continuity anchor.

## Integration Points

### Internal Boundaries

| Boundary | Communication | Notes |
|----------|---------------|-------|
| Spec Extraction ↔ Spec Storage | File write (design_spec.md) | Extraction script writes, manager loads on demand |
| Spec Storage ↔ Content Gathering | AI agent reads spec file | The AI loads the spec into context for the interviewer |
| Content Gathering ↔ Outline Building | Conversation context | Outline builder uses gathered info from conversation |
| Outline Building ↔ PPTX Generation | File write (outline + spec → spec_lock.md) | spec_to_lock.py is the bridge |
| SVG Generation ↔ Post-processing | File system (svg_output/ → scripts) | Scripts read SVG files, write processed SVG files, export PPTX |

### External Dependencies

| Dependency | Integration Pattern | Notes |
|------------|---------------------|-------|
| ppt-master SVG→DrawingML pipeline | Vendored fork in `scripts/svg_finalize/` and `scripts/svg_to_pptx/` | ~10-17 Python modules, adapted for spec-driven interface |
| python-pptx | pip dependency | PPTX reading (extraction) and PPTX writing (export). Dual use in two different phases. |
| lxml | pip dependency | Raw OOXML access for spec extraction (theme XML, font schemes) |
| Pillow | pip dependency | Image dimension analysis during spec extraction |

## Sources

- ppt-master v2.6.0 README.md (GitHub: hugohe3/ppt-master) — overall architecture and pipeline description
- ppt-master v2.6.0 docs/technical-design.md — detailed technical architecture, design philosophy, SVG rationale, post-processing internals, animation model
- ppt-master v2.6.0 skills/ppt-master/SKILL.md — workflow definition, role switching protocol, execution discipline rules
- ppt-master v2.6.0 skills/ppt-master/scripts/README.md — script inventory and pipeline structure
- PROJECT.md — project requirements, constraints, key decisions for ppt-skill

All sources accessed and verified 2026-05-06. Architecture confidence HIGH — the adapted components (ppt-master pipeline) are production-proven in v2.6.0, and the new components (spec extraction, content gathering) follow the same architectural patterns.

---
*Architecture research for: PPT Generation Skill*
*Researched: 2026-05-06*
