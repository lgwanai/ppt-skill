---
name: ppt-skill
description: >
  AI-powered PPT generation skill. Extracts design specs from existing PPTX files
  (colors, fonts, layouts, presentation logic) and generates natively editable .pptx
  files with real DrawingML shapes — not flattened images. Use the slash commands
  /ppt-spec to analyze a reference PPTX and extract a design spec, or /ppt to
  generate a presentation from content. Also use when the user asks to create slides,
  make a presentation, generate pptx, analyze pptx style, extract ppt design,
  convert svg to pptx, or needs a slide deck. Supports cover/toc/content/end_page
  recognition, multi-layout content pages, VL model analysis, agent-loop style
  verification with 90%+ fidelity, and multi-threaded generation.
---

# PPT Skill

Spec-driven PPT generation: extract design DNA from reference PPTX → gather
content → generate natively editable slides with agent-loop style verification.

## Slash Commands

| Command | Purpose |
|---------|---------|
| `/ppt-spec <file.pptx>` | Extract design spec from reference PPTX (cover/toc/content/end_page recognition + layout analysis) |
| `/ppt <content or file>` | Generate PPT: gather content → confirm outline → generate PPTX with style verification |

## Quick Start

```bash
# Extract design spec from a reference PPTX
python scripts/ppt_cli.py extract-spec reference.pptx

# Generate PPT from content (with agent-loop style verification)
python scripts/ppt_cli.py generate-pptx --spec specs/my_spec/ --outline outline.yaml -o output.pptx --workers 4

# Direct SVG to PPTX conversion
python scripts/ppt_cli.py convert slide1.svg slide2.svg -o output.pptx
```

## Commands

| Command | Description |
|---------|-------------|
| `extract-spec <pptx>` | Extract design spec → `specs/<name>/` directory (pages by type, assets, logic.yaml) |
| `generate-pptx` | Generate PPTX from spec + outline with agent-loop eval (--spec --outline -o --workers) |
| `convert <svg>... -o <pptx>` | Direct SVG→native-shape PPTX conversion |
| `list-specs` | List available design specs |
| `select-spec <name>` | Set active design spec |
| `gather-content <text\|file>` | Content gathering → slide-by-slide outline (adaptive questioning, 8-question cap) |
| `list-outlines` | List saved content outlines |

All scripts run from the skill root directory.

## Workflow

### `/ppt-spec` — Extract Design Spec

1. Read the reference PPTX with `python-pptx` + `lxml`
2. Extract color palette (12 HEX values from theme1.xml), fonts, backgrounds
3. Classify each page: cover / toc / transition / content / end_page
4. For content pages, identify layout sub-type: left_right, top_bottom, etc.
5. If `VL_ENABLED=true` in `config.txt`, use VL model for enhanced layout description
6. Extract reusable assets (backgrounds, images)
7. Analyze presentation logic (narrative, density rhythm, sections)
8. Save as directory: `specs/<name>/` with `spec.yaml`, `pages/`, `assets/`, `logic.yaml`

See [references/spec-format.md](references/spec-format.md).

### `/ppt` — Generate Presentation

1. **Assess sufficiency**: evaluate if user input has enough detail
2. **If insufficient**: ask section-level overview questions, then gap-fill (8-question cap)
3. **Generate outline**: slide-by-slide ContentOutline with titles, body, layout recommendations
4. **User confirms** the outline before generation
5. **Agent-loop generation** (per slide, in parallel):
   - Match slide to correct spec page type (cover→cover spec, content→matching layout spec, etc.)
   - Generate SVG with strict style constraints from spec
   - Evaluate: color match, font match, layout IoU, background, density
   - If score < 90% → add fix instructions, regenerate (max 5 iterations)
   - All slides generated in parallel via ThreadPoolExecutor
6. **Convert**: all SVGs → native-shape PPTX via SVG→DrawingML pipeline
7. Output: editable `.pptx` with real PowerPoint shapes

### Content Gathering Mode

When the user just wants to prepare content without generating PPTX:

1. Assess content sufficiency (4-dimension rubric: structure, detail, audience, scope)
2. If sufficient → directly generate outline
3. If insufficient → adaptive questioning: section-level overviews first, then gap-fill
4. Output: `outlines/<name>.yaml` — ready for `/ppt` generation

## Configuration

Copy `config.example.txt` to `config.txt` for VL model analysis in spec extraction:

```ini
VL_ENABLED=true
VL_PROVIDER=openai        # openai | anthropic | gemini | ollama
VL_MODEL=gpt-4o
VL_API_KEY=sk-...
```

## Resources

| Directory | Contents |
|-----------|----------|
| `scripts/ppt_cli.py` | CLI entry point |
| `scripts/ppt_skill/` | Python package (converter, spec, content, finalize, evaluator) |
| `references/spec-format.md` | Design spec schema |
| `assets/templates/icons/` | 11,600+ SVG icons (5 libraries) |
| `assets/templates/charts/` | 70+ chart/infographic SVGs |
| `config.example.txt` | Model configuration template |

## Requirements

Python 3.10+ with:

```
python-pptx>=0.6.21
Pillow>=9.0.0
PyYAML>=6.0
```

Optional for VL analysis: `openai`, `anthropic`, or `google-generativeai`.
