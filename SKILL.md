---
name: ppt-skill
description: >
  AI-powered PPT generation skill that analyzes existing PPTs to extract design
  specifications (colors, fonts, layouts, presentation logic) and generates new
  natively editable .pptx files from content outlines. Use when the user needs to
  create professional presentations, analyze reference PPTX files for style extraction,
  generate slide-by-slide content outlines, or convert SVG to native PowerPoint shapes.
  Triggers: "create a PPT", "make slides", "generate a presentation", "pptx",
  "extract design from pptx", "analyze pptx style", "ppt spec", "presentation outline",
  "convert svg to pptx".
---

# PPT Skill

Generate professional, natively editable PowerPoint files from content with
spec-driven design. Analyze existing PPTs to extract styles, gather content
through adaptive questioning, and produce PPTX with real DrawingML shapes.

## Quick Start

```bash
# Extract design spec from an existing PPTX
python scripts/ppt_cli.py extract-spec reference.pptx

# List and select active spec
python scripts/ppt_cli.py list-specs
python scripts/ppt_cli.py select-spec my_spec

# Gather content and generate slide outline
python scripts/ppt_cli.py gather-content "Write a presentation about Q3 results"

# Convert SVGs to PPTX (Phase 1 pipeline)
python scripts/ppt_cli.py convert slide1.svg slide2.svg -o output.pptx
```

## Core Workflow

### 1. Spec Extraction (analyze existing PPT)

Load a reference .pptx file and extract its design DNA:

```
python scripts/ppt_cli.py extract-spec reference.pptx
```

Output: `specs/<name>.yaml` — captures color palette (12 HEX values), font
families/sizes/hierarchy, spatial layout patterns, slide type classifications
(title, content, section_divider, image_text, data), and presentation rhythm
(breathing/dense/anchor density classification).

**See:** [references/spec-format.md](references/spec-format.md) for full spec schema.

### 2. Content Gathering (adaptive questioning)

Generate a slide-by-slide content outline from user input. When content is
insufficient, asks section-level overview questions first, then gap-fills
(8-question cap).

```bash
python scripts/ppt_cli.py gather-content "Topic or outline"
python scripts/ppt_cli.py gather-content content_file.md

# Skip questioning when input is detailed enough
python scripts/ppt_cli.py gather-content detailed_outline.md --mode skip_questions
```

Output: YAML ContentOutline with per-slide title, body, layout type recommendation,
and speaker notes — ready for PPT generation.

### 3. PPT Generation (spec-driven)

Generate a .pptx file from content outline + design spec. Every shape is a native
DrawingML element (text, rect, circle, path, gradient) — NOT flattened images.
Output is fully editable in PowerPoint.

**[Phase 4 — in development]**

### 4. Direct SVG → PPTX Conversion

Convert SVG files directly to native PowerPoint shapes using the forked
SVG→DrawingML pipeline:

```bash
python scripts/ppt_cli.py convert slide1.svg slide2.svg slide3.svg -o deck.pptx
python scripts/ppt_cli.py convert svg/*.svg -o deck.pptx --skip-check
```

## CLI Commands

| Command | Description |
|---------|-------------|
| `convert <svg>... -o <pptx>` | Convert SVGs to native-shape PPTX |
| `extract-spec <pptx>` | Extract design spec from reference PPTX |
| `list-specs` | List all available design specs |
| `select-spec <name>` | Set active design spec |
| `gather-content <text\|file>` | Generate content outline (adaptive questioning) |
| `list-outlines` | List saved content outlines |

All commands output structured results. Scripts must be run from the skill root
directory for template path resolution.

## Architecture

```
User Content
    │
    ├─→ Phase 3: ContentGatherer (adaptive questioning → ContentOutline)
    │       │
Spec File ←── Phase 2: SpecExtractor (PPTX → colors/fonts/layouts/logic)
    │       │
    └───→ Phase 4: PPT Generator (outline + spec → SVG → DrawingML PPTX)
                │
                └── Phase 1: SVG→DrawingML Converter (native shapes)
```

## Resources

### scripts/
- `ppt_cli.py` — Unified CLI entry point
- `ppt_skill/` — Python package (converter, spec, content, finalize, cli)

### references/
- `spec-format.md` — Design specification file schema

### assets/
- `templates/icons/` — 11,631 SVG icons (5 libraries)
- `templates/charts/` — 70+ chart/infographic SVGs

## Requirements

- Python 3.10+
- `python-pptx>=0.6.21`
- `Pillow>=9.0.0`
- `PyYAML>=6.0`

```bash
pip install -r requirements.txt
```
