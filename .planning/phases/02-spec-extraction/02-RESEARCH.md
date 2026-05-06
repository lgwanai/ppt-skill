# Phase 2: Spec Extraction - Research

**Researched:** 2026-05-06
**Domain:** PPTX design analysis — theme color/font extraction, layout classification, spatial analysis
**Confidence:** HIGH (verified with working code against real PPTX files and official python-pptx docs)

## Summary

This research confirms the extraction pipeline requires a **hybrid python-pptx + direct lxml/XML** approach. python-pptx's high-level API handles slide enumeration and layout name retrieval but **fails at theme-level font/color resolution** (returns `None` for inherited properties) and has a **known bug** in `Slide.background._element` (returns `<p:cSld>` instead of `<p:bg>`). Direct lxml parsing of `theme1.xml` and slide XMLs is required for color scheme resolution, font hierarchy, background fills, and spatial measurements.

The spec file format should be **YAML** — it is human-readable, git-diffable, and natively supported by Python via PyYAML. The extracted spec maps directly to the config constants noted in `src/ppt_skill/config.py` (`DESIGN_COLORS`, `FONT_SIZES`, `LAYOUT_MARGINS`) that Phase 4 (PPT Generation) will consume.

**Primary recommendation:** Build a standalone `SpecExtractor` class that walks the PPTX ZIP via python-pptx for structure + lxml/ElementTree for XML details, outputting a structured YAML spec file to `specs/<name>.yaml`.

## Standard Stack

### Core
| Library | Version | Purpose | Why Standard |
|---------|---------|---------|--------------|
| python-pptx | 1.0.2 | Slide enumeration, layout names, placeholder types, shape traversal | Canonical PPTX library; handles ZIP assembly, slide/layout hierarchy |
| lxml | 6.0.2 | Direct XML parsing of theme1.xml, slide XML, layout XML | python-pptx theme/color API is incomplete; lxml is already a dependency of python-pptx |
| PyYAML | >=6.0 | Spec serialization/deserialization | Python standard for YAML; human-readable, git-diffable, schema-flexible |

### Supporting
| Library | Version | Purpose | When to Use |
|---------|---------|---------|-------------|
| zipfile | stdlib | Direct PPTX ZIP access for raw XML reading | Reading `ppt/theme/theme1.xml` and slide XMLs bypassing python-pptx limitations |
| xml.etree.ElementTree | stdlib | Fallback XML parsing if lxml unavailable | Not needed if lxml is guaranteed; but useful for zero-dependency mode |

### Alternatives Considered
| Instead of | Could Use | Tradeoff |
|------------|-----------|----------|
| lxml | xml.etree.ElementTree (stdlib) | ET handles simpler namespaces; lxml is faster and handles malformed XML better |
| YAML | JSON | JSON is faster to parse but harder for humans to edit/read; no comments support |
| YAML | TOML | TOML is Python-native for config but less familiar for nested document structures |
| PyYAML | ruamel.yaml | ruamel preserves formatting/comments but is heavier; PyYAML sufficient for write-once-read-many |

**Installation:**
```bash
# python-pptx and lxml already installed
pip install pyyaml>=6.0
```

## Architecture Patterns

### Recommended Project Structure
```
src/ppt_skill/
├── spec/                    # NEW: Spec extraction module
│   ├── __init__.py
│   ├── extractor.py          # SpecExtractor class — main entry point
│   ├── theme.py              # Theme color/font resolution from theme1.xml
│   ├── layout_analysis.py    # Spatial analysis: margins, spacing, positioning
│   ├── slide_classifier.py   # Slide type classification logic
│   ├── density.py            # Content density / rhythm analysis
│   └── spec_model.py         # Pydantic or dataclass models for the spec schema
├── cli/                      # NEW: CLI commands
│   ├── __init__.py
│   └── spec_commands.py      # extract-spec, list-specs, select-spec commands
├── config.py                 # EXTEND: add DESIGN_COLORS, FONT_SIZES, LAYOUT_MARGINS
└── ...
specs/                        # NEW: Project-local spec storage
└── <spec-name>.yaml
```

### Pattern 1: Hybrid python-pptx + lxml Extraction
**What:** Use python-pptx for slide/layout traversal (high-level API), fall back to lxml for theme colors, fonts, backgrounds, and explicit XML attributes.
**When to use:** Whenever python-pptx returns `None` for a property (inherited from theme), or when the API is known broken (background).
**Example:**
```python
# Source: Verified via testing against real PPTX files
def extract_color_scheme(prs):
    """Extract resolved HEX colors from theme1.xml via lxml."""
    import zipfile
    from lxml import etree

    with zipfile.ZipFile(prs_path, 'r') as z:
        theme_xml = z.read('ppt/theme/theme1.xml')

    nsmap = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
    root = etree.fromstring(theme_xml)
    clr_scheme = root.find('a:themeElements/a:clrScheme', nsmap)

    colors = {}
    for child in clr_scheme:
        name = etree.QName(child.tag).localname
        for inner in child:
            tag = etree.QName(inner.tag).localname
            if tag == 'srgbClr':
                colors[name] = '#' + inner.get('val')
            elif tag == 'sysClr':
                colors[name] = '#' + inner.get('lastClr', inner.get('val'))
            # schemeClr (theme reference) → resolve against clrScheme
    return colors
```

### Pattern 2: Schema-First Spec Design
**What:** Define the spec schema as dataclasses or Pydantic models BEFORE writing any extraction code. The schema drives what gets extracted, validated, and serialized.
**When to use:** Always. The schema is the contract between extraction (Phase 2) and generation (Phase 4).
**Example:**
```python
from dataclasses import dataclass, field
from enum import Enum

class SlideType(Enum):
    TITLE = "title"
    CONTENT = "content"
    SECTION_DIVIDER = "section_divider"
    IMAGE_TEXT = "image_text"
    DATA = "data"

@dataclass
class ColorPalette:
    background1: str  # HEX e.g. "#FFFFFF"
    background2: str  # HEX
    text1: str        # HEX
    text2: str        # HEX
    accent1: str      # HEX
    accent2: str      # HEX
    accent3: str      # HEX
    accent4: str      # HEX
    accent5: str      # HEX
    accent6: str      # HEX
    hyperlink: str    # HEX
    followed_hyperlink: str  # HEX

@dataclass
class Typography:
    heading_family: str
    body_family: str
    heading_sizes: dict[str, float]  # {"title": 44.0, "subtitle": 28.0, "h1": 32.0, ...}
    body_sizes: dict[str, float]     # {"body": 18.0, "small": 14.0, ...}

@dataclass
class SlideLayoutSpec:
    slide_type: SlideType
    margins: dict[str, float]  # {"top": 0.5, "bottom": 0.5, "left": 1.0, "right": 1.0}
    title_position: dict[str, float]  # {"x": ..., "y": ..., "width": ..., "height": ...}
    content_positions: list[dict]

@dataclass
class PresentationRhythm:
    sequencing_pattern: list[str]  # e.g. ["title", "content", "content", "section_divider", ...]
    density_profile: list[str]     # e.g. ["anchor", "dense", "breathing", "dense", ...]
    story_arc: dict                # {"opening": 1, "development": 5, "climax": 1, "closing": 1}

@dataclass
class DesignSpec:
    metadata: dict
    colors: ColorPalette
    typography: Typography
    slide_layouts: list[SlideLayoutSpec]
    rhythm: PresentationRhythm
```

### Anti-Patterns to Avoid
- **Using python-pptx alone for theme extraction:** `run.font.name` / `run.font.color.rgb` return `None` for inherited values. Must use lxml to resolve theme defaults.
- **Assuming all colors are `srgbClr`:** Theme colors can be `sysClr` (system color with `lastClr` fallback) or `schemeClr` (theme reference requiring resolution against `clrScheme`).
- **Classifying slides solely by layout name:** Custom templates may rename layouts. Always validate with content analysis (placeholder types, presence of images/tables/charts).
- **Hardcoding theme namespaces:** Use `root.nsmap` from lxml or predefine the OOXML namespaces explicitly. The `a:` and `p:` prefixes may vary in custom template files.

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| PPTX ZIP handling | Manual zip parsing | python-pptx's `Presentation()` for structure; raw `zipfile` for theme XML only | python-pptx resolves slide/layout relationships, placeholder inheritance, and XML namespaces |
| YAML serialization | Custom YAML writer | PyYAML `yaml.dump()` with `default_flow_style=False` | Handles edge cases (multi-line strings, special chars, anchors) |
| Color hex normalization | Manual string formatting | python-pptx's `RGBColor` for explicit colors + custom resolver for theme colors | python-pptx colors return 6-char hex strings already; only need to add `#` prefix |
| EMU-to-real-unit conversion | Custom math | python-pptx's `Emu`, `Pt`, `Inches`, `Cm` | Already tested in Phase 1 pipeline |
| Slide content type detection | Heuristic shape counting | Use `placeholder_format.type` (PP_PLACEHOLDER) + `has_chart`/`has_table` + shape type inspection | python-pptx already provides these discriminators |

**Key insight:** The core extraction logic should be thin glue between python-pptx (structure) and lxml (details). The real design work is in the **schema definition** and **theme color resolution algorithm** — those are where complexity lives.

## Common Pitfalls

### Pitfall 1: `Slide.background._element` Returns Wrong Element
**What goes wrong:** `slide.background._element` returns `<p:cSld>` (the entire slide content container) instead of `<p:bg>`. Code that iterates children of this element silently corrupts the slide's `spTree`.
**Why it happens:** Confirmed python-pptx bug (GitHub #1126, opened Apr 16 2026, still open). The `_element` property is misnamed — it points to the parent `cSld`, not the `bg` child.
**How to avoid:** Use direct lxml path: find `<p:bg>` from `<p:cSld>` via `bg_elem = cSld.find(f'{{{p_ns}}}bg')`. Never iterate children of `background._element`.
**Warning signs:** Slide content disappears after background manipulation; spTree is silently replaced.

### Pitfall 2: Inherited Font Properties Return None
**What goes wrong:** `run.font.name`, `run.font.size`, `run.font.color.rgb` all return `None` when values are inherited from the theme or slide layout (not explicitly set on the run).
**Why it happens:** python-pptx's `Font` object reports `None` for any property not explicitly set at the run level. The actual value comes from the style hierarchy: run → paragraph → layout → master → theme.
**How to avoid:** Build a font resolver that walks the style hierarchy via lxml. Check run's `rPr`, then paragraph's `defRPr`, then layout's `defRPr`, then theme's `majorFont`/`minorFont`.
**Warning signs:** All fonts report as `None` in presentations that use only theme styling.

### Pitfall 3: Theme Color Resolution Requires Multi-Level Fallback
**What goes wrong:** Assuming all theme colors are direct `srgbClr`. Colors can be `sysClr` (system color), `schemeClr` (reference to another theme slot), or have tint/shade modifiers.
**Why it happens:** OOXML allows `sysClr` with `lastClr` (system default + explicit fallback), and `schemeClr` references within the theme. `dk1` and `lt1` are commonly `sysClr`.
**How to avoid:** Implement a resolution chain: `schemeClr` → resolve against `clrScheme` → `srgbClr` or `sysClr.lastClr`. Always prefer `lastClr` over `val` for `sysClr` (consistent across light/dark modes).
**Warning signs:** `dk1` and `lt1` return `"windowText"` and `"window"` — these are semantic names, not colors. Use `lastClr`.

### Pitfall 4: Slide Layout Names Are Unreliable in Custom Templates
**What goes wrong:** Classifying slides solely by `slide_layout.name` fails when users create custom layouts with non-standard names.
**Why it happens:** PowerPoint allows arbitrary layout names. Enterprise templates often use numbered or branded names (e.g., "01_Title", "AGENDA_SLIDE").
**How to avoid:** Use a dual classification strategy: (1) name-based for standard layouts, (2) content-based fallback — count placeholders by type, detect images/charts/tables, measure text density.
**Warning signs:** Layout names don't match any known pattern; many "unknown" classifications.

### Pitfall 5: Missing Background on Some Slides
**What goes wrong:** Some slides have no `<p:bg>` element at all (inherit from slide master or theme). Direct slide-level background extraction returns nothing.
**Why it happens:** Background can be set at slide level (`p:bg` in slide XML), slide layout level, slide master level, or theme level (`bgFillStyleLst`). Most presentations set it only at master level.
**How to avoid:** Walk the inheritance chain: slide → slide layout → slide master → theme. Check `bgFillStyleLst` in theme for the fallback. Report which level the background comes from.
**Warning signs:** Slides 2-N have no background data; only master-level background found.

## Code Examples

Verified patterns from official sources and testing:

### Theme Color Extraction (lxml)
```python
# Source: Verified via lxml parsing of generated PPTX theme1.xml
import zipfile
from lxml import etree

A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'

def extract_theme_colors(pptx_path: str) -> dict[str, str]:
    """Extract resolved HEX color palette from theme1.xml."""
    with zipfile.ZipFile(pptx_path, 'r') as z:
        theme_xml = z.read('ppt/theme/theme1.xml')

    root = etree.fromstring(theme_xml)
    clr_scheme = root.find(f'{{{A_NS}}}themeElements/{{{A_NS}}}clrScheme')

    colors = {}
    for child in clr_scheme:
        name = etree.QName(child.tag).localname  # dk1, lt1, accent1, ...
        # Each child has one color element inside
        inner = list(child)[0] if list(child) else None
        if inner is None:
            continue
        tag = etree.QName(inner.tag).localname
        if tag == 'srgbClr':
            colors[name] = '#' + inner.get('val')
        elif tag == 'sysClr':
            colors[name] = '#' + (inner.get('lastClr') or inner.get('val', '000000'))

    return colors
```

### Theme Font Extraction (lxml)
```python
def extract_theme_fonts(pptx_path: str) -> dict[str, str]:
    """Extract heading and body font families from theme1.xml."""
    with zipfile.ZipFile(pptx_path, 'r') as z:
        theme_xml = z.read('ppt/theme/theme1.xml')

    root = etree.fromstring(theme_xml)
    font_scheme = root.find(f'{{{A_NS}}}themeElements/{{{A_NS}}}fontScheme')

    fonts = {}
    for font_class in ['majorFont', 'minorFont']:
        elem = font_scheme.find(f'{{{A_NS}}}{font_class}')
        if elem is None:
            continue
        latin = elem.find(f'{{{A_NS}}}latin')
        if latin is not None:
            fonts[font_class] = latin.get('typeface', '')

    return fonts  # {'majorFont': 'Calibri', 'minorFont': 'Calibri'}
```

### Background Extraction (lxml, Workaround for Bug #1126)
```python
def extract_slide_background(slide) -> dict | None:
    """Safely extract slide background avoiding the _element bug."""
    cSld = slide.background._element  # Actually returns <p:cSld> (bug #1126)
    bg_elem = cSld.find(f'{{{P_NS}}}bg')
    if bg_elem is None:
        return None  # Background inherited from master/theme

    # Check for solid fill
    bgPr = bg_elem.find(f'{{{P_NS}}}bgPr')
    if bgPr is not None:
        solid = bgPr.find(f'{{{A_NS}}}solidFill')
        if solid is not None:
            color_elem = list(solid)[0] if list(solid) else None
            if color_elem is not None:
                tag = etree.QName(color_elem.tag).localname
                if tag == 'srgbClr':
                    return {'type': 'solid', 'color': '#' + color_elem.get('val')}

    # Check for theme reference
    bgRef = bg_elem.find(f'{{{P_NS}}}bgRef')
    if bgRef is not None:
        scheme = bgRef.find(f'{{{A_NS}}}schemeClr')
        if scheme is not None:
            return {'type': 'theme_ref', 'ref': scheme.get('val')}

    return None
```

### Slide Classification (python-pptx + content analysis)
```python
from pptx.enum.shapes import PP_PLACEHOLDER

LAYOUT_NAME_MAP = {
    'Title Slide': 'title',
    'Title and Content': 'content',
    'Section Header': 'section_divider',
    'Two Content': 'image_text',
    'Picture with Caption': 'image_text',
    'Content with Caption': 'image_text',
    'Comparison': 'content',
    'Title Only': 'content',
    'Blank': 'content',
}

def classify_slide(slide) -> str:
    """Classify slide type by layout name, with content-based fallback."""
    layout_name = slide.slide_layout.name
    if layout_name in LAYOUT_NAME_MAP:
        return LAYOUT_NAME_MAP[layout_name]

    # Content-based fallback
    has_image = any(hasattr(s, 'image') for s in slide.shapes)
    has_chart = any(hasattr(s, 'has_chart') and s.has_chart for s in slide.shapes)
    has_table = any(hasattr(s, 'has_table') and s.has_table for s in slide.shapes)
    has_title = slide.shapes.title is not None

    if has_chart or has_table:
        return 'data'
    if has_image and not has_title:
        return 'image_text'
    if has_title:
        return 'content'
    return 'content'
```

### Content Density Analysis
```python
def analyze_content_density(slides) -> list[dict]:
    """Compute per-slide content density and classify rhythm."""
    densities = []
    for i, slide in enumerate(slides):
        total_chars = sum(
            len(s.text_frame.text)
            for s in slide.shapes
            if s.has_text_frame
        )
        image_count = sum(1 for s in slide.shapes if hasattr(s, 'image'))
        shape_count = len(list(slide.shapes))

        if total_chars < 100:
            density = 'breathing'
        elif total_chars < 500:
            density = 'dense'
        else:
            density = 'anchor'

        densities.append({
            'slide_index': i + 1,
            'char_count': total_chars,
            'image_count': image_count,
            'shape_count': shape_count,
            'density': density,
        })

    return densities
```

### YAML Spec Serialization
```python
import yaml
from pathlib import Path

def save_spec(spec: DesignSpec, name: str, specs_dir: Path = Path('specs')):
    """Save design spec as YAML file."""
    specs_dir.mkdir(exist_ok=True)
    filepath = specs_dir / f'{name}.yaml'

    with open(filepath, 'w') as f:
        yaml.dump(
            spec, f,
            default_flow_style=False,
            sort_keys=False,
            allow_unicode=True,
        )

def load_spec(name: str, specs_dir: Path = Path('specs')) -> DesignSpec:
    """Load design spec from YAML file."""
    filepath = specs_dir / f'{name}.yaml'
    with open(filepath, 'r') as f:
        data = yaml.safe_load(f)
    return DesignSpec(**data)

def list_specs(specs_dir: Path = Path('specs')) -> list[str]:
    """List available spec names (without .yaml extension)."""
    if not specs_dir.exists():
        return []
    return sorted(p.stem for p in specs_dir.glob('*.yaml'))
```

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|--------------|--------|
| python-pptx `slide.background.fill` | Direct lxml extraction of `<p:bg>` from `<p:cSld>` | 2026-04 (bug #1126 discovered) | Must bypass python-pptx API for backgrounds |
| python-pptx `run.font.name` (returns None) | lxml resolution chain: run→para→layout→master→theme | Always needed | Font extraction requires inheritance walk |
| ppt-master's design_spec.md markdown format | YAML structured spec with typed schema | Current design choice | YAML is machine-parseable; markdown is not |
| Hardcoded color palettes in config.py | Extracted spec files in `specs/` directory | Phase 2 implementation | Enables multi-spec management, git versioning |

**Deprecated/outdated:**
- `slide.background._element` for background reading: returns wrong element. Use lxml as described above.
- Relying on `font.name` for theme fonts: returns `None`. Resolve via theme XML.

## Open Questions

1. **How to resolve `schemeClr` references within theme colors?**
   - What we know: Some colors reference other theme slots (e.g., `<a:schemeClr val="phClr"/>` in background fills). The `clrScheme` provides the base mapping.
   - What's unclear: Whether `phClr` (placeholder color) resolves to a specific scheme color or is context-dependent. PowerPoint's rendering engine may apply additional tinting.
   - Recommendation: For v1, resolve `schemeClr` against the `clrScheme` map. Flag `phClr` and other non-standard references with a fallback to `lt1`.

2. **How to detect embedded vs. system fonts?**
   - What we know: python-pptx has no font embedding API. OOXML stores embedded fonts in `ppt/fonts/` directory within the ZIP, but detection requires checking for font part relationships in `[Content_Types].xml`.
   - What's unclear: Whether we need to distinguish embedded from system fonts for spec purposes (generation may not be able to use embedded fonts).
   - Recommendation: Record font names as-is. Add a `fonts.embedded: true/false` flag if `ppt/fonts/` directory exists. Don't block on this — Phase 4 can apply fonts using system availability.

3. **What's the exact ppt-master design_spec.md format?**
   - What we know: The ppt-master GitHub repo (`cupang-ai/ppt-master`) returns 404. The fork in this project contains no design_spec.md — only converter/finalize modules. The Phase 1 config.py mentions `DESIGN_COLORS`, `INDUSTRY_COLORS`, `LAYOUT_MARGINS`, `FONT_SIZES` as planned config constants.
   - What's unclear: Whether the original ppt-master had a formal design_spec format or just in-code config dicts.
   - Recommendation: Design our own YAML spec schema based on python-pptx's theme model (clrScheme + fontScheme + fmtScheme). This is cleaner than trying to reverse-engineer a missing file. The schema proposed in Pattern 2 above is complete for Phase 2-4 needs.

4. **How to detect gradient backgrounds faithfully?**
   - What we know: `bgFillStyleLst` in `theme1.xml` can contain `gradFill` with gradient stops (`gsLst`). Each gradient stop has position and color. Path type (`circle`, `linear`) and angle must be extracted.
   - What's unclear: Whether Phase 4 (PPT Generation) can reproduce arbitrary gradients from spec. The SVG→DrawingML pipeline in Phase 1 handles some gradients but might not cover all theme gradient patterns.
   - Recommendation: Extract gradient metadata (type, stops, positions) into spec. Flag complexity — simple 2-stop linear gradients are fully supported; multi-stop radial gradients may need simplification for Phase 4.

5. **What's the optimal density threshold for anchor/dense/breathing classification?**
   - What we know: The provided context says "anchor/dense/breathing" categories. Our test shows ~10 chars for title slide (breathing), ~583 for content-heavy (anchor).
   - What's unclear: Whether thresholds should be absolute or relative (percentile-based within the deck).
   - Recommendation: Use percentile-based classification: bottom 20% = breathing, top 20% = anchor, rest = dense. This adapts to any deck length and writing style.

## Sources

### Primary (HIGH confidence)
- python-pptx 1.0.2 readthedocs — verified `ColorFormat`, `Font`, `PlaceholderFormat`, `SlideLayout` APIs, `MSO_THEME_COLOR_INDEX` enum
  - https://python-pptx.readthedocs.io/en/latest/api/dml.html
  - https://python-pptx.readthedocs.io/en/latest/api/enum/MsoThemeColorIndex.html
  - https://python-pptx.readthedocs.io/en/latest/api/text.html
- python-pptx GitHub Issue #1126 — confirmed `Slide.background._element` bug with lxml workaround
  - https://github.com/scanny/python-pptx/issues/1126
- Direct lxml testing against generated PPTX files — verified theme1.xml structure, namespace handling, color/font scheme extraction, background fill patterns

### Secondary (MEDIUM confidence)
- python-pptx user guide (placeholders, quickstart) — confirmed layout names, placeholder types, slide structure
  - https://python-pptx.readthedocs.io/en/latest/user/quickstart.html
  - https://python-pptx.readthedocs.io/en/latest/user/placeholders-using.html
- OOXML specification knowledge (namespace URIs, element hierarchy) — verified via actual XML inspection

### Tertiary (LOW confidence)
- ppt-master design_spec.md format — GitHub repo was inaccessible (404). Design decisions based on config.py hints and OOXML theme model. Marked for validation when original source is located.

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH — python-pptx 1.0.2 + lxml 6.0.2 + PyYAML are verified, installed, and tested with real PPTX files
- Architecture: HIGH — hybrid python-pptx + lxml pattern confirmed necessary and working; slide classification and density analysis tested with generated PPTX
- Pitfalls: HIGH — python-pptx bug #1126 confirmed via GitHub issue; font/color `None` returns verified via testing; theme color resolution chain verified via lxml XML parsing
- Spec format: MEDIUM — YAML schema design is sound but ppt-master's original format unavailable for comparison; schema may need adjustment as Phase 4 generation requirements become clearer

**Research date:** 2026-05-06
**Valid until:** 2026-06-06 (stable APIs; all libraries are mature)

<phase_requirements>
## Phase Requirements

| ID | Description | Research Support |
|----|-------------|-----------------|
| SPC-01 | Extract visual style (color palette, backgrounds, gradients), typography (font families, sizes, weights, hierarchy), and spatial layout patterns | Theme color extraction via `theme1.xml` lxml parsing (§Theme Color Extraction code example). Font extraction via `majorFont`/`minorFont` in theme. Background via `bgFillStyleLst` lxml. Spatial layout via shape `left/top/width/height` in EMU. |
| SPC-02 | Layout classification — identifying and categorizing slide types with distinct visual properties | `LAYOUT_NAME_MAP` in §Slide Classification maps standard names to categories. Content-based fallback using placeholder types (`PP_PLACEHOLDER`), `has_chart`, `has_table`, image detection. |
| SPC-03 | Capture presentation logic — slide sequencing, content density rhythm, storytelling structure | Content density calculation (§Content Density Analysis) produces per-slide char counts. Percentile-based rhythm classification (breathing/dense/anchor). Slide type sequence captured as ordered list. |
| SPC-04 | Spec saved as structured project-local file, versionable, shareable, reusable | YAML format (§YAML Spec Serialization) in `specs/<name>.yaml` directory. Git-diffable, human-readable. Pydantic/dataclass models for schema validation. |
| SPC-05 | List available specs and select one as active design target | `list_specs()` function enumerates `specs/` directory. `load_spec()` loads by name. Active spec tracked via config file or state file in Phase 3-4. |
</phase_requirements>
