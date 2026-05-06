# Phase 1: Pipeline Foundation - Research

**Researched:** 2026-05-06
**Domain:** SVG-to-DrawingML conversion (forking ppt-master's proven 17-module pipeline)
**Confidence:** HIGH

## Summary

The ppt-master project contains a mature, battle-tested SVG→DrawingML converter pipeline at `skills/ppt-master/scripts/`. The pipeline converts SVG files to native PowerPoint shapes (DrawingML XML embedded inside PPTX), producing fully editable PPTs where every shape, text box, gradient, and image is a native PowerPoint element. The codebase is production-grade, with a 17-module svg_to_pptx/ package, an 8-module svg_finalize/ post-processing package, 11,631 SVG icons, 73 chart templates, and 23 layout packs.

Forking this pipeline for ppt-skill Phase 1 means extracting the essential converter modules while understanding the full dependency graph to avoid breaking dual-consumer modules (e.g., `use_expander.py` and `tspan_flattener.py` both delegate to `svg_finalize/` submodules, and `drawingml_converter.py` calls both in-memory — removing `svg_finalize/` silently breaks native PPTX output).

**Primary recommendation:** Fork the entire `scripts/` directory, keep all 17 svg_to_pptx/ modules + svg_finalize/ submodules as a unit, strip only ppt-master-specific project management scripts (project_manager.py, image_gen.py, notes_to_audio.py, etc.), and validate with the existing quality checker and a sample SVG→PPTX round-trip.

<phase_requirements>
## Phase Requirements

| ID | Description | Research Support |
|----|-------------|-----------------|
| PIP-01 | Fork and adapt ppt-master's core SVG→DrawingML converter modules (17-module pipeline) for spec-driven generation — every shape, text box, and gradient is a native PowerPoint element | §2 documents all 17 modules, their dependency graph, and which are core vs auxiliary |
| PIP-02 | SVG quality checker validates generated SVGs against ppt-master compatibility rules before conversion | §3 documents the full banned-feature list, validation checks, and banned patterns |
| PIP-03 | Post-processing pipeline (icon embedding, tspan flattening, image alignment) produces PPTX-parsable SVG output | §4 documents the 4-stage finalize pipeline and cross-dependencies with the converter |
| PIP-04 | Inherit ppt-master's icon library (11,600+ SVG icons) and chart templates (70+ chart types) for visual variety | §5 documents template directory structure, icon libraries, chart count, and layout packs |
</phase_requirements>

## Standard Stack

### Core

| Library | Version | Purpose | Why Standard |
|---------|---------|---------|--------------|
| python-pptx | >=0.6.21 | PPTX file creation/assembly, slide dimension management | Only library that builds .pptx from Python; the pipeline uses it for base PPTX scaffolding |
| Pillow | >=9.0.0 | Image dimension reading, SVG-to-PNG fallback rendering | Intrinsic image size detection for preserveAspectRatio handling; no viable alternative |
| lxml (stdlib fallback) | built-in | XML parsing via ElementTree | ppt-master's pipeline uses stdlib `xml.etree.ElementTree` exclusively — no lxml dependency |

### Supporting

| Library | Version | Purpose | When to Use |
|---------|---------|---------|-------------|
| CairoSVG | optional | High-quality SVG→PNG for Office compat mode | Only needed if Office LTSC 2021 compatibility is required; the native shapes mode needs no PNG renderer |
| svglib + reportlab | optional | Lightweight SVG→PNG fallback | CairoSVG alternative (~5MB vs ~25MB); gradients may be lost |

### Dependencies Not Needed

The pipeline's native shapes mode (the primary mode for Phase 1) requires zero rendering libraries. All conversion is pure Python stdlib + python-pptx + Pillow. The SVG→PNG renderers (cairosvg, svglib) are only for the legacy compatibility mode, which embeds SVG as images rather than native shapes.

**Installation:**
```bash
pip install python-pptx>=0.6.21 Pillow>=9.0.0
```

## Architecture Patterns

### Module Dependency Graph (svg_to_pptx/ package)

```
pptx_cli.py ──→ pptx_discovery.py
           ──→ pptx_builder.py ──→ drawingml_converter.py ──→ drawingml_context.py
                                   │                          │
                                   ├─ drawingml_elements.py ──┤
                                   │  (rect, circle, path,    │
                                   │   text, image, etc.)     │
                                   ├─ drawingml_styles.py ────┤
                                   │  (fill, stroke, shadow)  │
                                   ├─ drawingml_paths.py ─────┤
                                   │  (path parsing/norm)     │
                                   ├─ drawingml_utils.py ─────┤
                                   │  (coords, color, fonts)  │
                                   ├─ use_expander.py ────────┤
                                   │  └─ depends on:          │
                                   │     svg_finalize.embed_icons
                                   └─ tspan_flattener.py ─────┤
                                      └─ depends on:          │
                                         svg_finalize.flatten_tspan
           ──→ pptx_dimensions.py
           ──→ pptx_media.py (PNG renderer)
           ──→ pptx_slide_xml.py (legacy mode only)
           ──→ pptx_notes.py (notes)
           ──→ pptx_narration.py (narration)
```

**Key insight:** The `drawingml_converter.py` is the central dispatcher. It calls `use_expander.py` and `tspan_flattener.py` in-memory (before XML processing), which themselves delegate to `svg_finalize/` submodules. Removing either svg_finalize/embed_icons.py or svg_finalize/flatten_tspan.py silently breaks native PPTX output — icons drop out and multi-line text collapses.

### Cross-Package Dependencies (dual-consumer modules)

| Module | Consumed By | Why Critical |
|--------|-------------|-------------|
| `svg_finalize/embed_icons.py` | `finalize_svg.py` (disk) AND `use_expander.py` (memory) | Icon expansion must work identically in both pipelines |
| `svg_finalize/flatten_tspan.py` | `finalize_svg.py` (disk) AND `tspan_flattener.py` (memory) | Tspan flattening must work identically in both pipelines |
| `svg_finalize/align_embed_images.py` | `finalize_svg.py` (disk) only | Image alignment + embedding; converter reads base64 directly |
| `svg_finalize/svg_rect_to_path.py` | `finalize_svg.py` (disk) only | Converts rounded rects before conversion |

### Recommended Project Structure for Fork

```
ppt-skill/
├── src/
│   └── ppt_skill/
│       ├── __init__.py
│       ├── converter/                    # Forked svg_to_pptx/ (17 modules)
│       │   ├── __init__.py
│       │   ├── cli.py                    # Renamed pptx_cli.py
│       │   ├── builder.py                # Renamed pptx_builder.py
│       │   ├── converter.py              # Renamed drawingml_converter.py
│       │   ├── context.py                # Renamed drawingml_context.py
│       │   ├── elements.py               # Renamed drawingml_elements.py
│       │   ├── paths.py                  # Renamed drawingml_paths.py
│       │   ├── styles.py                 # Renamed drawingml_styles.py
│       │   ├── utils.py                  # Renamed drawingml_utils.py
│       │   ├── dimensions.py             # Renamed pptx_dimensions.py
│       │   ├── discovery.py              # Renamed pptx_discovery.py
│       │   ├── media.py                  # Renamed pptx_media.py
│       │   ├── narration.py              # Renamed pptx_narration.py
│       │   ├── notes.py                  # Renamed pptx_notes.py
│       │   ├── slide_xml.py              # Renamed pptx_slide_xml.py
│       │   ├── use_expander.py
│       │   └── tspan_flattener.py
│       ├── finalize/                     # Forked svg_finalize/ (8 modules)
│       │   ├── __init__.py
│       │   ├── embed_icons.py
│       │   ├── align_embed_images.py
│       │   ├── crop_images.py
│       │   ├── embed_images.py
│       │   ├── fix_image_aspect.py
│       │   ├── flatten_tspan.py
│       │   └── svg_rect_to_path.py
│       ├── quality.py                    # Forked svg_quality_checker.py
│       ├── config.py                     # Canvas formats, color schemes
│       └── pipeline.py                   # New: unified pipeline orchestration
├── templates/                            # Forked templates/ (flat copy)
│   ├── icons/                            # 11,631 SVG icons (5 libraries)
│   │   ├── chunk-filled/
│   │   ├── tabler-filled/
│   │   ├── tabler-outline/
│   │   ├── phosphor-duotone/
│   │   └── simple-icons/
│   ├── charts/                           # 73 chart template SVGs
│   └── layouts/                          # 23 layout packs
├── tests/
│   ├── test_converter.py
│   ├── test_quality_checker.py
│   ├── test_finalize.py
│   └── fixtures/
│       ├── sample_input.svg
│       ├── sample_output.pptx
│       └── svg_pptx_pairs/
├── requirements.txt
└── SKILL.md
```

### Pattern 1: In-Memory Preprocessing Before XML Conversion

**What:** The converter applies icon expansion and tspan flattening to the SVG element tree in memory before dispatching to element converters, rather than requiring a separate `finalize_svg` disk step.

**When to use:** Always for native shapes mode. This is how the converter reads `svg_output/` directly.
**Example (from drawingml_converter.py:308-323):**
```python
# Expand <use data-icon="..."/> placeholders in-memory
icons_dir = Path(__file__).resolve().parent.parent.parent / 'templates' / 'icons'
if icons_dir.exists():
    from .use_expander import expand_use_data_icons
    expanded = expand_use_data_icons(root, icons_dir)

# Flatten positional <tspan> into independent <text> elements
from .tspan_flattener import flatten_positional_tspans
if flatten_positional_tspans(tree):
    ...
```

### Pattern 2: ConvertContext Recursive Traversal

**What:** A `ConvertContext` dataclass is passed through recursive SVG tree traversal, accumulating translate/scale/opacity/filter state. Each element converter calls `ctx.child()` to create a child context with accumulated transforms.

**Key state carried by ConvertContext:**
- `translate_x/y`, `scale_x/y` — accumulated geometric transforms
- `transform_matrix` — full affine matrix for matrix-capable elements (images)
- `inherited_styles` — CSS properties inherited from parent `<g>` elements
- `defs` — resolved `<defs>` dictionary for gradient/filter/clipPath lookups
- `anim_targets` — top-level semantic groups for per-element entrance animation
- `media_files` / `rel_entries` — accumulated during conversion for PPTX assembly

### Pattern 3: Soft Imports with Fallback

**What:** Optional dependencies (cairosvg, pptx_animations) use try/except ImportError patterns so the core pipeline works without them.

**Example (from pptx_media.py:9-20):**
```python
PNG_RENDERER: str | None = None
try:
    import cairosvg
    PNG_RENDERER = 'cairosvg'
except (ImportError, OSError):
    try:
        from svglib.svglib import svg2rlg
        PNG_RENDERER = 'svglib'
    except (ImportError, OSError):
        pass
```

### Anti-Patterns to Avoid

- **Importing `config.py` with full dependency on project structure:** The `config.py` imports assume the skill's full directory tree. When forking, paths need adjustment to be relative to the forked package root.
- **Hardcoding `templates/icons/` path from scripts/:** Multiple modules resolve `templates/icons/` via `Path(__file__).resolve().parent.parent.parent / 'templates' / 'icons'`. This 3-level parent traversal must be adjusted for the forked package structure.

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| SVG path parsing | Custom path command parser | Forked `drawingml_paths.py` | Handles 15 SVG path commands (M, L, C, Q, A, S, T, etc.) with arc-to-cubic-bezier conversion — 429 lines of battle-tested code |
| Color parsing | Custom hex/rgba parser | Forked `drawingml_utils.py` parse_hex_color() | Handles #RGB, #RRGGBB, #RRGGBBAA, named colors, url() references |
| PPTX OOXML assembly | Manual ZIP + XML | python-pptx for scaffolding + forked builder for slide XML injection | python-pptx handles [Content_Types].xml, presentation.xml, slide layout registration; builder injects DrawingML slide XML |
| Font fallback resolution | Custom font mapping table | Forked `drawingml_utils.py` FONT_FALLBACK_WIN / EA_FONTS | 40+ font mappings for macOS→Windows CJK font substitution, genre-based serif/sans-serif defaults |
| Gradient to DrawingML | Custom gradient XML builder | Forked `drawingml_styles.py` build_gradient_fill() | Handles linearGradient, radialGradient, gradientTransform, spreadMethod — 250+ lines |

**Key insight:** The pipeline is a finely tuned OOXML XML generator. Hand-rolling any part risks producing silently invalid PPTX files that PowerPoint opens but cannot edit (the "shapes become unselectable" failure mode).

## Common Pitfalls

### Pitfall 1: Removing `svg_finalize/` Breaks Native PPTX Output

**What goes wrong:** Deleting the svg_finalize/ package (thinking it's only for the svg_final/ disk pipeline) silently breaks `use_expander.py` and `tspan_flattener.py`, which import from it during in-memory conversion. Icons drop out of native PPTX output without errors.

**Why it happens:** The pipeline was designed with two consumers of svg_finalize/: the disk-based `finalize_svg.py` and the memory-based `svg_to_pptx/`. Both need the same icon expansion and tspan flattening logic.

**How to avoid:** Treat `svg_to_pptx/` and `svg_finalize/` as a coupled unit. Fork both together.

**Warning signs:** `convert_svg_to_slide_shapes()` succeeds but SVG output is missing icons and multiline text is collapsed onto one line.

### Pitfall 2: Hardcoded Paths Break After Fork

**What goes wrong:** Multiple modules resolve `templates/icons/` using `Path(__file__).resolve().parent.parent.parent / 'templates' / 'icons'`. After moving `svg_to_pptx/` into `src/ppt_skill/converter/`, the 3-parent traversal points to the wrong location.

**Why it happens:** The original code assumes `scripts/svg_to_pptx/` → `scripts/` → `skills/ppt-master/` → `templates/`. In the forked structure, this path chain changes.

**How to avoid:** Either add a configurable `TEMPLATES_DIR` constant, or adjust the relative path computation in every module that resolves templates. Audit all occurrences:
- `drawingml_converter.py:308` (icons_dir)
- `use_expander.py:33` (scripts_dir for svg_finalize import)
- `tspan_flattener.py:39` (scripts_dir for svg_finalize import)
- `finalize_svg.py:130` (icons_dir)
- `embed_icons.py:52` (DEFAULT_ICONS_DIR)

### Pitfall 3: `config.py` Import Chain Contamination

**What goes wrong:** `config.py` imports are optional (try/except) in `pptx_dimensions.py` and `project_utils.py`, but `svg_quality_checker.py` unconditionally imports several ppt-master-specific modules (`project_utils`, `error_helper`, `update_spec`).

**Why it happens:** The quality checker was built as an integrated tool within the ppt-master ecosystem. When forking, these cross-references fail.

**How to avoid:** Strip or rewrite the quality checker's ppt-master-specific imports:
- `CANVAS_FORMATS` can come from a local config module
- `ErrorHelper` is only used for verbose error formatting — safe to remove
- `_parse_spec_lock` is for spec_lock drift detection — Phase 1 doesn't need it yet
- Template-mode checks are for ppt-master's layout library management — Phase 1 doesn't need them

### Pitfall 4: Python 3.10+ Compatibility

**What goes wrong:** The codebase uses `from __future__ import annotations` and `str | None` type syntax throughout. These require Python 3.10+ for the union type syntax.

**Why it happens:** The `from __future__ import annotations` import makes `str | None` valid at type-check time on 3.7+, but at runtime the pipe operator for types was only introduced in Python 3.10.

**How to avoid:** Set `requires-python >= 3.10` in the project configuration. The forked code already uses this syntax, and Python 3.10 is a reasonable minimum for 2026.

### Pitfall 5: Negative SVG Attribute Parsing

**What goes wrong:** Many SVG attributes (x, y, width, height, font-size) can be negative in SVG but are silently zeroed or produce invalid EMU values in the converter.

**Why it happens:** The `_f()` helper converts strings to floats (returning 0 for None). The converter skips elements with `w <= 0 or h <= 0`, but doesn't guard against negative x/y causing off-canvas placement.

**How to avoid:** The quality checker should validate that shapes are within the viewBox. Add a bounds check step for Phase 1 validation.

## Code Examples

### Standalone Conversion (existing pipeline pattern)

```python
# Source: ppt-master svg_to_pptx/pptx_builder.py
from pathlib import Path
from ppt_skill.converter.converter import convert_svg_to_slide_shapes
from ppt_skill.converter.builder import create_pptx_with_native_svg

# Single SVG to PPTX with native shapes
svg_files = [Path("slide1.svg"), Path("slide2.svg")]
create_pptx_with_native_svg(
    svg_files=svg_files,
    output_path=Path("output.pptx"),
    use_native_shapes=True,
)
```

### Quality Checker Usage (existing pattern)

```python
# Source: ppt-master svg_quality_checker.py
from ppt_skill.quality import SVGQualityChecker

checker = SVGQualityChecker()
result = checker.check_file("slide1.svg")

# result['errors'] contains banned features detected:
# - "Detected forbidden <mask> element (PPT does not support SVG masks)"
# - "Detected forbidden <style> element (use inline attributes instead)"
# - "Detected forbidden rgba() color (use fill-opacity/stroke-opacity instead)"
```

### Minimal Pipeline Entry Point (what to build for Phase 1)

```python
# New: ppt_skill/pipeline.py
"""Unified pipeline: SVG → quality check → convert → PPTX."""
from pathlib import Path
from ppt_skill.quality import SVGQualityChecker
from ppt_skill.converter.builder import create_pptx_with_native_svg


def convert_svg_to_pptx(svg_files: list[Path], output_path: Path) -> bool:
    """Convert SVG files to native-shape PPTX with quality validation."""
    # 1. Quality check
    checker = SVGQualityChecker()
    for svg in svg_files:
        result = checker.check_file(str(svg))
        if not result['passed']:
            raise ValueError(f"SVG quality check failed for {svg}: {result['errors']}")

    # 2. Convert
    return create_pptx_with_native_svg(
        svg_files=svg_files,
        output_path=output_path,
        use_native_shapes=True,
    )
```

## Complete Banned Feature List (from svg_quality_checker.py)

The quality checker validates these banned features. Each is a hard error:

| Category | Banned Feature | Error Message |
|----------|---------------|---------------|
| XML | Malformed XML (HTML entities: &nbsp; &mdash; etc.) | "Invalid XML — SVG must be well-formed XML" |
| Clipping/Masking | `<mask>` elements | "PPT does not support SVG masks" |
| Clipping/Masking | `clip-path` on non-image elements | "clip-path is only allowed on <image> elements" |
| Clipping/Masking | `clip-path` referencing undefined `<clipPath>` | "no matching <clipPath> definition found" |
| Style System | `<style>` elements | "use inline attributes instead" |
| Style System | `class` attribute | "use inline styles instead" |
| Style System | `id` attribute + `<style>` together | "CSS selectors forbidden" |
| Style System | `<?xml-stylesheet>` processing instruction | "external CSS references forbidden" |
| Style System | `<link rel="stylesheet">` | "external CSS references forbidden" |
| Style System | `@import` | "external CSS references forbidden" |
| Structure | `<foreignObject>` | "use <tspan> for manual line breaks" |
| Structure | `<symbol>` + `<use>` (both present) | "use basic shapes or simple <use> instead" |
| Text/Fonts | `<textPath>` | "path text is incompatible with PPT" |
| Text/Fonts | `@font-face` | "use system font stack" |
| Animation | `<animate*>` elements | "SVG animations are not exported" |
| Animation | `<set>` elements | "SVG animations are not exported" |
| Animation | `<script>` elements | "scripts and event handlers forbidden" |
| Animation | Event attributes (onclick, onload, etc.) | "forbidden event attributes" |
| Colors | `rgba()` colors | "use fill-opacity/stroke-opacity instead" |
| Opacity | `<g opacity="...">` | "set opacity on each child element individually" |
| Opacity | `<image opacity="...">` | "use overlay mask approach" |
| Other | `<iframe>` | "should not appear in SVG" |
| Other | Undefined marker references | "marker-start/marker-end ... but no <marker> element found" |

**Quality checker also validates (warnings, not errors):**
- viewBox presence and format
- Font stack ending with PPT-safe family (Arial, Microsoft YaHei, SimSun, etc.)
- width/height consistency with viewBox
- Overly long single-line text (>100 chars)
- Image file existence and resolution vs display size
- spec_lock drift (colors, fonts, sizes not in spec_lock.md) — Phase 1 should skip this

## Templates Inventory

### Icon Libraries (11,631 SVG files)

| Library | Count | ViewBox | Style |
|---------|-------|---------|-------|
| tabler-outline | ~5,000+ | 24x24 | Stroke (currentColor) |
| simple-icons | ~3,400+ | 24x24 | Fill (brand logos) |
| phosphor-duotone | ~1,200+ | 256x256 | Duotone (single color + 0.2-opacity backplate) |
| tabler-filled | ~1,000+ | 24x24 | Fill |
| chunk-filled | ~640+ | 16x16 | Fill |

Usage: `<use data-icon="tabler-outline/home" x="100" y="200" width="48" height="48" fill="#0076A8"/>`

### Chart Templates (73 SVG files)

Covers: bar_chart, line_chart, pie_chart, donut_chart, area_chart, bubble_chart, funnel_chart, gantt_chart, radar_chart, gauge_chart, waterfall_chart, heatmap, box_plot_chart, scatter_plot, dual_axis_line_chart, treemap, sunburst, sankey_diagram, bullet_chart, butterfly_chart, dumbbell_chart, cycle_diagram, flywheel_diagram, fishbone_diagram, concentric_circles, chevron_process, agenda_list, comparison_table, basic_table, feature_matrix_table, financial_statement_table, comparison_columns, consulting_table, client_server_flow, bcg_matrix, ansoff_matrix

### Layout Packs (23 directories)

Each contains `design_spec.md` + per-page SVG files covering full deck architecture:
academic_defense, ai_ops, anthropic, china_telecom_template, exhibit, google_style, government_blue, government_red, mckinsey, medical_university, pixel_retro, psychology_attachment, smart_red, 中国电建_常规, 中国电建_现代, 中汽研_商务, 中汽研_常规, 中汽研_现代, etc.

### Copy Strategy

Templates are pure file copies — no pip install needed. Use `shutil.copytree()` or `setup.py package_data` to include them in the forked package. Total size: icons ~15-20MB (mostly simple SVGs), charts ~2MB, layouts ~5MB. With compression, this adds ~15MB to the package.

## Open Questions

1. **Template scope for Phase 1**
   - What we know: 11,631 icons, 73 charts, 23 layout packs exist in ppt-master
   - What's unclear: Which subset is needed for Phase 1 validation vs. Phase 4 PPT Generation
   - Recommendation: Copy all icons + charts during Phase 1 (they're svg_to_pptx dependencies for icon expansion). Layouts can wait for Phase 2 (Spec Extraction) or Phase 4

2. **How to handle `config.py` and `project_utils.py`**
   - What we know: Both modules are optional imports in most converter modules, but contain large amounts of ppt-master-specific configuration (DESIGN_COLORS, INDUSTRY_COLORS, LAYOUT_MARGINS, FONT_SIZES)
   - What's unclear: Whether to strip them entirely or extract the essential `CANVAS_FORMATS` dict only
   - Recommendation: Extract `CANVAS_FORMATS` (8 entries: ppt169, ppt43, wechat, xiaohongshu, moments, story, banner, a4) into a minimal `config.py`. Strip DESIGN_COLORS, INDUSTRY_COLORS, LAYOUT_MARGINS, FONT_SIZES — those belong to Phase 2's spec extraction

3. **CLI entry point design**
   - What we know: ppt-master's `pptx_cli.py` is a full CLI with 20+ arguments (transitions, animations, narration, compat mode, SVG source directory, etc.)
   - What's unclear: Which CLI arguments Phase 1 needs vs. Phase 4 PPT Generation
   - Recommendation: Phase 1 only needs `--input <svg_file>` and `--output <pptx_file>`. Build a minimal CLI in `pipeline.py`. Full CLI comes in Phase 4

4. **Test data preservation**
   - What we know: ppt-master has `examples/` and `projects/` directories with known-good SVG→PPTX pairs
   - What's unclear: Whether to copy those as test fixtures for regression testing
   - Recommendation: Copy at least 5-10 known-good SVG→PPTX pairs from ppt-master examples into `tests/fixtures/svg_pptx_pairs/` for regression validation. The quality checker's `--all` mode can batch-validate them

## Sources

### Primary (HIGH confidence — direct code inspection)
- `/Users/wuliang/project/ppt-master/skills/ppt-master/scripts/` — Complete scripts directory (36 entries)
- `/Users/wuliang/project/ppt-master/skills/ppt-master/scripts/svg_to_pptx/` — All 17 converter modules read in full
- `/Users/wuliang/project/ppt-master/skills/ppt-master/scripts/svg_finalize/` — All 8 post-processing modules read
- `/Users/wuliang/project/ppt-master/skills/ppt-master/scripts/svg_quality_checker.py` — 1,210 lines, banned feature list, all validation checks documented
- `/Users/wuliang/project/ppt-master/skills/ppt-master/scripts/finalize_svg.py` — 4-stage pipeline (embed-icons, align-images, flatten-text, fix-rounded)
- `/Users/wuliang/project/ppt-master/skills/ppt-master/templates/` — Directory structure confirmed (icons: 11,631 SVGs, charts: 73 SVGs, layouts: 23 dirs)
- `/Users/wuliang/project/ppt-master/skills/ppt-master/requirements.txt` — Dependency versions confirmed

### Secondary (MEDIUM confidence)
- `/Users/wuliang/project/ppt-master/skills/ppt-master/scripts/config.py` — Canvas formats, color schemes, SVG constraints (729 lines)
- `/Users/wuliang/project/ppt-master/skills/ppt-master/scripts/project_utils.py` — Project structure validation, canvas format detection

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH — All dependencies confirmed from requirements.txt and import analysis; the pipeline's native shapes mode needs only python-pptx + Pillow
- Architecture: HIGH — Full dependency graph mapped by reading all 25 modules; dual-consumer entanglement documented; all 4 cross-package dependencies identified
- Pitfalls: HIGH — 5 concrete pitfalls identified from code analysis, each with root cause, detection, and prevention strategy
- Templates: HIGH — Directory structure confirmed by filesystem; 11,631 icons counted via `find *.svg | wc -l`; 73 charts and 23 layouts confirmed

**Research date:** 2026-05-06
**Valid until:** 2026-08-06 (stable codebase, 90-day validity)
