# Design Specification File Format
#
# Stored as specs/<name>.yaml — machine-readable design contract
# extracted from reference PPTX files by SpecExtractor.
#
# This file is consumed by Phase 4 (PPT Generation) as the
# source-of-truth for visual styling.

## Top-Level Structure

```yaml
# DesignSpec — the complete extracted design specification
metadata:
  name: "string"           # Spec identifier (from PPTX filename)
  source_pptx: "string"    # Original reference file path
  extracted_at: "string"   # ISO timestamp of extraction
  slide_count: 0           # Total slides in reference PPTX

color_palette:
  # 12 HEX color entries resolved from theme1.xml
  # Keys: accent1..6, dark1, dark2, light1, light2, hlink, folHlink
  accent1: "#HEX"
  accent2: "#HEX"
  background: "#HEX"
  text_primary: "#HEX"
  # ...

typography:
  heading_family: "string"   # e.g. "Calibri"
  body_family: "string"      # e.g. "Calibri"
  heading_sizes:             # min/max/median from text runs
    title: 0
    subtitle: 0
  body_sizes:
    body: 0
    small: 0

layout_margins:
  # Spatial measurements in inches
  top: 0.0
  bottom: 0.0
  left: 0.0
  right: 0.0
  title_y: 0.0
  content_y: 0.0

slides:
  # Per-slide classification and layout data
  - slide_number: 1
    slide_type: "title"     # One of: title, content, section_divider, image_text, data
    layout_name: "string"   # Original PowerPoint layout name
    title: "string"         # Extracted title text
    content_density: 0.0    # Character count proxy

presentation_rhythm:
  # Sequencing and density analysis
  density_sequence: ["breathing", "dense", "anchor", ...]
  story_arc: "string"       # Heuristic: "ascending", "pyramid", "oscillating"
  total_slides: 0
```

## ContentOutline Format

The companion format produced by Phase 3 (Content Gathering), consumed by Phase 4:

```yaml
presentation_title: "string"
presentation_subtitle: "string"
target_audience: "string"
presentation_purpose: "string"
sections: ["Section 1", "Section 2"]
spec_name: "string"       # Active spec to apply
slides:
  - slide_number: 1
    title: "Slide Title"
    body:
      - "Bullet point 1"
      - "Bullet point 2"
    layout_type: "content"  # title, content, two_column, section_divider, image_text, data
    notes: "Speaker notes"
    image_hint: ""          # Keyword for image generation/search
    section_name: "Section 1"
```
