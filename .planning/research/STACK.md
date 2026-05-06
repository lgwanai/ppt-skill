# Stack Research

**Domain:** AI-powered PPT generation with design spec extraction
**Researched:** 2026-05-06
**Confidence:** HIGH

## Recommended Stack

### Core Technologies

| Technology | Version | Purpose | Why Recommended |
|------------|---------|---------|-----------------|
| Python | 3.12+ | Runtime | ppt-master targets 3.10+; Pillow 12.x and PyMuPDF 1.27.x require >=3.10. Pin to 3.12 for LTS stability, typing, and match all dependency ranges. |
| python-pptx | 1.0.2 | PPTX read/write | De facto standard for OOXML PPTX manipulation in Python. Reads layouts, themes, shapes; writes native DrawingML. 10+ years mature, MIT license, actively maintained (v1.0.2 released Aug 2024). No viable alternative exists. |
| lxml | 6.1.0 | Deep OOXML traversal & SVG parsing | Required for deep .pptx inspection (theme XML, slide layouts, color schemes) beyond what python-pptx's high-level API exposes. Also used by ppt-master's SVG parser to read/viewBox attributes. 20+ years of libxml2/libxslt reliability. |
| PyMuPDF | 1.27.2.3 | PDF→Markdown for source material | Used by ppt-master's `pdf_to_md.py` pipeline. Only needed if users provide PDF source material (not core for spec extraction from PPTX). AGPL licensed — confirm compatibility before distribution. Required for spec extraction only if analyzing PDF-exported PPTs as references. |

### Supporting Libraries

| Library | Version | Purpose | When to Use |
|---------|---------|---------|-------------|
| Pillow | 12.2.0 | Image processing, PNG fallback rendering | Required: (a) extract embedded images from PPTX during spec analysis, (b) PNG fallback rendering in Office compatibility mode, (c) image aspect ratio correction. |
| svglib | 1.5.0 | SVG→PNG fallback conversion | Lightweight (~5MB) SVG→ReportLab→PNG renderer. Used as Office compatibility fallback when target PowerPoint doesn't support native SVG (Office LTSC 2021 and earlier). Prefer cairosvg for fidelity, but svglib has fewer system dependencies. |
| cairosvg | 2.7+ | SVG→PNG high-fidelity fallback | Higher fidelity than svglib (full gradient/filter support). Requires `brew install cairo` on macOS. Use in production when visual quality matters for compatibility mode. |
| reportlab | 4.0+ | PDF output for svglib | Transitive dependency of svglib. Not directly used in our pipeline. |
| PyYAML | 6.0+ | Spec file serialization | Design specs saved as structured YAML files (`specs/name.yaml`). Human-readable, git-diffable, easy for LLM to generate and parse. |
| openpyxl | 3.1.0 | Excel→Markdown for source material | Only if users provide .xlsx as source material. Transitively needed by ppt-master's `excel_to_md.py` pipeline. |

### Development Tools

| Tool | Purpose | Notes |
|------|---------|-------|
| uv | Package management | Modern Python package manager (Astral, 2025+). Faster than pip, lock-file support, no venv complexity. Use `uv pip install` or `uv sync` with `pyproject.toml`. |
| pytest | Testing | Standard Python testing framework. Test SVG→DrawingML converters, spec extraction logic, PPTX generation integrity. |
| ruff | Linting & formatting | Single tool for linting + formatting. Replaces flake8/isort/black. Fast (Rust-based). |

## Installation

```bash
# Core — PPTX read/write + deep OOXML + YAML spec files
uv pip install python-pptx==1.0.2 lxml==6.1.0 PyYAML>=6.0

# Image processing (extract images from PPTX, PNG fallback)
uv pip install Pillow==12.2.0

# SVG→PNG fallback for Office compatibility mode (pick one)
# Option A: cairosvg (high fidelity, needs system cairo)
brew install cairo           # macOS only
uv pip install cairosvg>=2.7

# Option B: svglib (lightweight, no system deps)
uv pip install svglib>=1.5.0 reportlab>=4.0

# Optional — PDF source material support
uv pip install PyMuPDF==1.27.2.3  # AGPL — verify license compatibility

# Dev
uv pip install pytest ruff
```

## Alternatives Considered

| Category | Recommended | Alternative | Why Not |
|----------|-------------|-------------|---------|
| PPTX library | python-pptx 1.0.2 | python-docx | Wrong format — reads/writes .docx, not .pptx |
| PPTX library | python-pptx 1.0.2 | pptx-template (jinja2-pptx) | Template-fill approach — great for variable substitution into fixed templates, but we generate programmatic shapes from SVG; mismatched paradigm |
| PPTX library | python-pptx 1.0.2 | aspose.slides (commercial) | Paid license (~$1000+/year), not practical for open-source skill |
| PPTX reading | python-pptx 1.0.2 | openxml (manual zip+xml) | Reinventing the wheel; python-pptx handles the ZIP container and relationship management |
| PPTX reading | python-pptx 1.0.2 | OOXML SDK (C#/.NET) | Wrong language ecosystem; Python + python-pptx is simpler and sufficient |
| XML parsing | lxml 6.1.0 | xml.etree.ElementTree (stdlib) | lxml is 5-10x faster with better XPath, namespace handling, and the `cleaner` module for stripping namespaces from SVG — critical for our SVG parser |
| SVG→DrawingML | Fork ppt-master's 17 modules | Build from scratch | ppt-master's SVG→DrawingML converter is proven (22 example projects, 309 pages), handles 17 shape types, and has been battle-tested with Claude/GPT/Gemini |
| SVG→DrawingML | Fork ppt-master | python-pptx's built-in SVG support | python-pptx v1.0.2 has no native SVG→DrawingML path — it only inserts SVG as an OLE object or linked image, not as editable shapes |
| PDF analysis | PyMuPDF | pdfplumber | pdfplumber text-only; PyMuPDF extracts images, tables, fonts, and layout metadata — better for source-to-content pipeline |
| Spec format | YAML | JSON | YAML is more human-readable, supports comments, and LLMs produce cleaner YAML output with fewer formatting errors |
| Package manager | uv | pip/poetry/pipenv | uv is the 2025–2026 standard: 10-100x faster than pip, no venv complexity, lockfile support, compatible with pip/poetry workflows |
| Image processing | Pillow 12.2.0 | opencv-python | Overkill — we only need resize, format conversion, basic manipulation. Pillow is lighter (~5MB vs ~50MB). |

## What NOT to Use

| Avoid | Why | Use Instead |
|-------|-----|-------------|
| **Image-flattening PPTX tools** (Gamma, Beautiful.ai, etc.) | Produce one large image per slide — not editable shapes. Violates core requirement. | python-pptx + SVG→DrawingML pipeline |
| **HTML-to-PPTX converters** | Convert HTML/CSS to PPTX but lose native PowerPoint shapes, produce layout artifacts, and can't handle complex SVG→DrawingML conversion. | Direct SVG→DrawingML converter |
| **LibreOffice UNO API (headless)** | Requires LibreOffice installation, ~500MB dependency, slow subprocess-based conversion, fragile to version mismatches. | python-pptx for direct OOXML manipulation |
| **AppleScript/VBA automation** | Platform-locked, requires installed PowerPoint, not automatable in CLI/server. Violates "no PowerPoint required" constraint. | python-pptx (pure Python, cross-platform) |
| **pptx-template / jinja2-pptx** | Template-fill paradigm — great for report generation, but our pipeline generates shapes from SVG dynamically. Templates are fixed; our layouts are generated. | Direct shape construction via python-pptx |
| **Node.js/TypeScript stack** | ppt-master's core pipeline (SVG→DrawingML, PPTX construction, source-to-MD) is Python-only. Mixing Node.js adds complexity with zero benefit. No Node.js equivalent of python-pptx/lxml exists. | Python-only stack |
| **Python <3.10** | python-pptx 1.0.2 requires >=3.8, but Pillow 12.x and PyMuPDF 1.27.x require >=3.10. Pin to 3.12 for forward compatibility. | Python 3.12+ |

## Stack Patterns by Variant

**If spec extraction is the priority (Phase 1):**
- Use python-pptx + lxml for PPTX traversal
- LLM analyzes extracted raw data (text, fonts, colors, layouts) and produces structured YAML
- No SVG→DrawingML needed yet — this phase is read-only with structured output

**If generation is the priority (Phase 2+):**
- Use the full ppt-master SVG→DrawingML pipeline (fork + adapt)
- python-pptx assembles the final PPTX
- cairosvg or svglib for Office compatibility PNG fallback
- LLM generates design-spec-compliant SVG per slide, then converter runs

**If dual-platform skill format is needed:**
- Write SKILL.md per Opencode/Claude Code conventions
- All executable logic in Python scripts under `skills/ppt-spec-skill/scripts/`
- Skill file references scripts via relative paths
- Use `#!/usr/bin/env python3` shebangs for portability

## Version Compatibility

| Package A | Compatible With | Notes |
|-----------|-----------------|-------|
| python-pptx 1.0.2 | Python 3.8–3.12 | Tested at v1.0.2; no blockers for 3.13 but not yet in classifiers |
| lxml 6.1.0 | Python 3.8–3.14 | Full 3.13 and 3.14 wheel support added in 6.x |
| Pillow 12.2.0 | Python 3.10–3.14 | Dropped 3.9 support at v11.x |
| PyMuPDF 1.27.2.3 | Python 3.10–3.14 | Requires >=3.10 since v1.25.x |
| svglib 1.5.0 | Python 3.8+ | Lightweight, broad compatibility |
| cairosvg 2.7+ | Requires system Cairo | macOS: `brew install cairo`; Ubuntu: `apt install libcairo2-dev` |

## Core Architecture of ppt-master's SVG→DrawingML Pipeline

ppt-master v2.6.0 uses 17 SVG→DrawingML converter modules. Key converters:

1. **Shape converters**: rect→p:sp, circle/ellipse→p:sp, line→p:cxnSp, path→p:cxnSp, polygon/polyline→p:cxnSp
2. **Text converters**: text→p:sp with txBody (font, size, color, alignment, word wrap)
3. **Image converters**: image→p:pic with rasters and clip rects
4. **Group/SVG converters**: g/svg→p:grpSp (group shapes for logical layers)
5. **Style mapping**: SVG fill→DrawingML solidFill, SVG stroke→DrawingML ln/outline, opacity mapping
6. **Chart converters**: For ppt-master's 70+ chart templates (likely not needed in MVP)
7. **Layout engine**: 22 layout templates, absolute positioning from SVG coordinates

Our fork will need: (a) shape converters, (b) text converters, (c) style mapping, (d) image embedding — the chart and layout template modules can be deferred.

## Skill Format: Opencode vs Claude Code

Both platforms use Markdown-based skill definitions with frontmatter metadata. Key differences:

| Aspect | Opencode Skill | Claude Code Skill |
|--------|---------------|-------------------|
| Location | `~/.opencode/skills/name/SKILL.md` | `.claude/skills/name/SKILL.md` or registered plugin |
| Frontmatter | `description`, `color`, `tools` | `description`, `tools`, `model` |
| Scripts | Shebanged Python in `skills/name/scripts/` | Same pattern |
| Discovery | Directory scan | Directory scan or marketplace |

**Strategy**: Write a single SKILL.md compatible with both. Core logic in Python scripts. Provide install instructions for both platforms. The skill definition file is platform-agnostic — only the installation path differs.

## Sources

- [PyPI: python-pptx 1.0.2](https://pypi.org/project/python-pptx/) — Verified latest version, release date Aug 7 2024 — HIGH confidence
- [PyPI: lxml 6.1.0](https://pypi.org/project/lxml/) — Verified latest version, release date Apr 18 2026 — HIGH confidence
- [PyPI: PyMuPDF 1.27.2.3](https://pypi.org/project/PyMuPDF/) — Verified latest version, release date Apr 24 2026 — HIGH confidence
- [PyPI: Pillow 12.2.0](https://pypi.org/project/pillow/) — Verified latest version, release date Apr 1 2026 — HIGH confidence
- [ppt-master v2.6.0 README](https://github.com/hugohe3/ppt-master) — Architecture overview, dependency list, SVG→DrawingML design — HIGH confidence
- [ppt-master skills/requirements.txt](https://raw.githubusercontent.com/hugohe3/ppt-master/main/skills/ppt-master/requirements.txt) — Verified dependency declarations — HIGH confidence
- [python-pptx docs](https://python-pptx.readthedocs.io/en/latest/) — Official documentation — HIGH confidence
- [ppt-master skills SKILL.md](https://github.com/hugohe3/ppt-master/tree/main/skills/ppt-master) — Skill format reference for Claude Code — MEDIUM confidence

---

*Stack research for: PPT Spec Skill (AI-powered PPT generation)*
*Researched: 2026-05-06*
