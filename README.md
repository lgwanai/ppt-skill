# ppt-skill

AI-powered PPT generation skill — extracts design specs from existing PPTX and generates natively editable .pptx with agent-loop style verification.

## Quick Start

```bash
# Extract design spec from reference PPTX
python scripts/ppt_cli.py extract-spec reference.pptx

# Generate PPT from content
python scripts/ppt_cli.py generate-pptx --spec specs/my_spec/ --outline outline.yaml -o output.pptx
```

## Features

- **Spec extraction**: Analyze existing PPTX to extract colors, fonts, layouts, page types (cover/toc/content/end)
- **Agent-loop generation**: Generate → screenshot → VL compare → fix → repeat until 90%+ style match
- **Native PPTX output**: Editable shapes via python-pptx, no SVG intermediate layer
- **VL model support**: Doubao Seed Dream / GPT-4o / Claude for layout analysis
- **Multi-layout**: Cover and content pages use distinct layout images (logos, backgrounds)
- **Slide validation**: Pre-flight checks for overflow, zero-width, overlaps before agent loop

## Architecture

```
PPTX → SpecExtractor → specs/<name>/
                          ├── spec.yaml (palette, fonts)
                          ├── pages/ (per-page blueprints)
                          ├── assets/ (extracted images)
                          └── logic.yaml
         ↓
Content → PPTX Generator → .pptx → QuickLook screenshot
         ↓                        ↓
    Agent Loop ←── VL Compare ←──┘
    (max 5 iter, 85% threshold)
```

## Requirements

- Python 3.10+
- python-pptx, Pillow, PyYAML, lxml
- Optional: openai (for VL model)

## License

MIT
