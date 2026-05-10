---
name: ppt-skill
description: AI-powered PPT generation skill — extracts design specs from existing PPTX files and generates natively editable .pptx files with agent-loop style verification. Use /ppt-spec to analyze a reference PPTX and extract design spec, or /ppt to generate a presentation. Also use when the user asks to create slides, make a presentation, generate pptx, analyze pptx style, or convert content to slides. Supports cover/toc/content/end_page recognition, multi-layout content, VL model analysis, and 90%+ style fidelity.
---

# PPT Skill

Spec-driven PPT generation: extract design DNA from reference PPTX → gather
content → generate natively editable slides with agent-loop style verification.

## Slash Commands

| Command | Purpose |
|---------|---------|
| `/ppt-spec <file.pptx>` | Extract design spec from reference PPTX (cover/toc/content/end_page recognition + layout analysis) |
| `/ppt-outline <content>` | Generate content outline using WPS model from prompt-ppt-content.md (titles + body only) |
| `/ppt-diagram <type> <content>` | Generate diagram (architecture/flowchart/sequence) from description |
| `/ppt <content or file>` | Generate PPT: gather content → confirm outline → generate PPTX → apply layout design |

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

### `/ppt-outline` — Generate Content Outline

1. **Analyze input**: Topic, theme, article, or existing outline
2. **Agent loop**: Use `prompt-ppt-content.md` principles
   - If insufficient input → adaptive questioning (8-question cap)
   - Extract sections and key points
3. **Apply WPS model**: Each slide has:
   - **W (What)**: Navigation label
   - **P (Point)**: Core conclusion (bold, prominent)
   - **S (Support)**: Evidence (data, cases, quotes)
4. **Output**: Markdown outline ready for `/ppt` generation

### `/ppt-diagram` — Generate Diagram

1. **Classify diagram type**: architecture, data-flow, flowchart, sequence, class, er-diagram, network-topology
2. **Extract structure**: identify nodes, edges, containers from description
3. **Apply layout rules**:
   - Architecture: horizontal layers (Client → Gateway → Services → Data)
   - Flowchart: top-to-bottom, diamond for decisions
   - Sequence: vertical lifelines with horizontal messages
4. **Generate SVG** using 7 visual styles:
   - Style 1: Flat Icon (default, white bg)
   - Style 2: Dark Terminal (monospace, dark bg)
   - Style 3: Blueprint (cyan on navy)
   - Style 4: Notion Clean (minimal)
   - Style 5: Glassmorphism (translucent)
   - Style 6: Claude Official (warm cream)
   - Style 7: OpenAI (green accent)
5. **Export**: SVG + PNG via rsvg-convert

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
5. **Phase 1: Spec Matching** (agent-loop, per slide, in parallel):
   - Match slide to correct spec page type (cover→cover spec, content→matching layout spec, etc.)
   - Generate SVG with strict style constraints from spec
   - Evaluate: color match, font match, layout IoU, background, density
   - If score < 90% → add fix instructions, regenerate (max 5 iterations)
   - All slides generated in parallel via ThreadPoolExecutor
6. **Phase 2: Layout Design** (agent-loop, using `prompt-ppt-layout.md`):
   - Apply WPS hierarchy: W (navigation) small, P (point) large/bold, S (support) regular
   - Typography: sans-serif fonts, 1.5x line spacing, justified text
   - Color rules: single primary color, dark text (#333-#444)
   - Whitespace: 8-12% margins, breathing room
   - Visuals: Bootstrap Icons from `assets/bootstrap-icons-1.13.1/`
   - Validate layout checklist (max 3 iterations)
7. **Convert**: all SVGs → native-shape PPTX via SVG→DrawingML pipeline
8. Output: editable `.pptx` with real PowerPoint shapes

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

| Directory/File | Contents |
|----------------|----------|
| `scripts/ppt_cli.py` | CLI entry point |
| `scripts/ppt_skill/` | Python package (converter, spec, content, layout, diagram, finalize, evaluator) |
| `references/spec-format.md` | Design spec schema |
| `references/prompt-ppt-content.md` | Content generation principles (WPS model) |
| `references/prompt-ppt-layout.md` | Layout design principles |
| `references/diagram-styles/` | Diagram style references (7 styles + icons) |
| `templates/diagram/` | Diagram SVG templates |
| `assets/Illustration/` | **400 vector illustrations** (unDraw, MIT License) |
| `assets/bootstrap-icons-1.13.1/` | 2000+ SVG icons for slides |
| `assets/templates/icons/` | 11,600+ SVG icons (5 libraries) |
| `assets/templates/charts/` | 70+ chart/infographic SVGs |
| `config.example.txt` | Model configuration template |

## Illustration Library

`assets/Illustration/` contains **283** open-source vector illustrations from unDraw:

### Categories

| Category | Count | Examples |
|----------|-------|----------|
| **AI & Technology** | 10+ | `ai-chat`, `ai-code-assistant`, `coding-assistant` |
| **Data Analytics** | 15+ | `analyze`, `data-table`, `analytics-setup` |
| **Team & Collaboration** | 10+ | `collaboration`, `followers`, `eating-together` |
| **Success & Achievement** | 10+ | `accomplishments`, `completed`, `goals` |
| **Business & Marketing** | 15+ | `business-decisions`, `customer-survey`, `contract` |
| **Learning & Education** | 10+ | `continuous-learning`, `exam-prep`, `five-year-plan` |
| **Documents & Files** | 10+ | `document-ready`, `file-analysis`, `add-files` |

### Search Illustrations

```bash
# Search for AI-related illustrations
python3 scripts/ppt_skill/illustration_finder.py search "ai"

# Search by category
python3 scripts/ppt_skill/illustration_finder.py list --category data

# Get random illustrations
python3 scripts/ppt_skill/illustration_finder.py random --count 5

# List all categories
python3 scripts/ppt_skill/illustration_finder.py list --categories
```

### Use in PPT

```python
from pptx import Presentation
from pptx.util import Inches

# Add illustration to slide
prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])
slide.shapes.add_picture(
    "assets/Illustration/undraw_ai-chat_ljb9.svg",
    Inches(1), Inches(1), Inches(4)
)
```

### Customize Colors

Edit SVG `fill` attributes to match your brand:

```xml
<!-- Original orange -->
<path fill="#f9a826" d="..." />

<!-- Change to brand blue -->
<path fill="#0070C0" d="..." />
```

## Requirements

Python 3.10+ with:

```
python-pptx>=0.6.21
Pillow>=9.0.0
PyYAML>=6.0
```

Optional for VL analysis: `openai`, `anthropic`, or `google-generativeai`.
