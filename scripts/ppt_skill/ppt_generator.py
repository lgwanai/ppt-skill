"""Multi-threaded PPT generator — spec-driven with two-phase agent-loop evaluation.

Phase 1: Spec Matching — color, fonts, layout match
Phase 2: Layout Design — WPS hierarchy, typography, whitespace (using prompt-ppt-layout.md)

Orchestrates parallel slide generation:
  1. Load spec (directory-based) + content outline
  2. Match each slide to the correct spec page type
  3. Generate slides in parallel using ThreadPoolExecutor
  4. Each slide goes through Phase 1 agent-loop (spec matching)
  5. Each slide goes through Phase 2 agent-loop (layout design)
  6. Convert all SVGs to native-shape PPTX
"""

from __future__ import annotations

import json
import yaml
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable

from ppt_skill.slide_generator import generate_slide_with_loop, SlideResult
from ppt_skill.style_evaluator import StyleReport
from ppt_skill.layout.designer import LayoutDesigner, LayoutResult


MAX_WORKERS = 4  # Default parallel workers


@dataclass
class GenerationResult:
    """Complete PPT generation result."""
    slide_results: list[SlideResult] = field(default_factory=list)
    layout_results: list[LayoutResult] = field(default_factory=list)
    output_path: Path | None = None
    passed_count: int = 0
    failed_count: int = 0
    layout_passed_count: int = 0
    total_iterations: int = 0
    layout_iterations: int = 0
    avg_score: float = 0.0


def _load_spec_pages(spec_dir: Path) -> tuple[list[dict], dict]:
    """Load all spec pages and metadata from a spec directory.

    Supports three formats:
      1. VL JSON format: specs/<name>/spec.json + <layout>.json files
      2. Legacy directory YAML: specs/<name>/pages/<type>/page_*.yaml
      3. Legacy flat YAML: specs/<name>/spec.yaml

    Returns:
        (spec_pages, metadata) where each page dict contains:
          page_type, layout_sub_type, elements, background, colors, typography, assets_dir
    """
    spec_pages: list[dict] = []

    # ── Format 1: VL JSON spec ──
    spec_json = spec_dir / "spec.json"
    if spec_json.exists():
        try:
            master = json.loads(spec_json.read_text(encoding="utf-8"))
            metadata = {
                "colors": master.get("palette", {}),
                "typography": master.get("typography", {}),
                "canvas": master.get("canvas", {}),
                "spec_name": master.get("spec_name", spec_dir.name),
                "source": master.get("source", ""),
                "slide_count": master.get("slide_count", 0),
                "assets_dir": str(spec_dir / "assets"),
            }

            for layout_ref in master.get("layouts", []):
                layout_file = layout_ref.get("file", "")
                if not layout_file:
                    continue
                layout_path = spec_dir / layout_file
                if not layout_path.exists():
                    continue
                try:
                    page_data = json.loads(layout_path.read_text(encoding="utf-8"))
                    spec_pages.append({
                        "page_type": page_data.get("page_type", "content"),
                        "layout_sub_type": page_data.get("layout_sub_type", "full_width"),
                        "elements": page_data.get("elements", []),
                        "element_groups": page_data.get("element_groups", []),
                        "design_patterns": page_data.get("design_patterns", []),
                        "background": page_data.get("background", {
                            "type": "solid", "color": "#FFFFFF", "description": ""
                        }),
                        "background_color": _resolve_bg_color(page_data.get("background", {})),
                        "background_description": page_data.get("background", {}).get("description", ""),
                        "colors": metadata["colors"],
                        "typography": metadata["typography"],
                        "canvas": metadata["canvas"],
                        "slide_indices": layout_ref.get("slide_indices", []),
                        "assets_dir": str(spec_dir / "assets"),
                    })
                except Exception:
                    pass

            if spec_pages:
                return spec_pages, metadata
        except Exception:
            pass

    # ── Format 2: Legacy directory YAML ──
    pages_dir = spec_dir / "pages"
    if pages_dir.exists():
        for page_type_dir in sorted(pages_dir.iterdir()):
            if not page_type_dir.is_dir():
                continue
            page_type = page_type_dir.name
            if page_type == "content":
                for sub_dir in sorted(page_type_dir.iterdir()):
                    if sub_dir.is_dir():
                        layout_sub_type = sub_dir.name
                        for yaml_file in sorted(sub_dir.glob("page_*.yaml")):
                            try:
                                data = yaml.safe_load(yaml_file.read_text(encoding="utf-8"))
                                if data:
                                    data["page_type"] = page_type
                                    data["layout_sub_type"] = layout_sub_type
                                    spec_pages.append(data)
                            except Exception:
                                pass
            else:
                for yaml_file in sorted(page_type_dir.glob("page_*.yaml")):
                    try:
                        data = yaml.safe_load(yaml_file.read_text(encoding="utf-8"))
                        if data:
                            data["page_type"] = page_type
                            data["layout_sub_type"] = ""
                            spec_pages.append(data)
                    except Exception:
                        pass

    # ── Format 3: Legacy single YAML ──
    if not spec_pages:
        spec_yaml = spec_dir / "spec.yaml"
        if spec_yaml.exists():
            try:
                spec_data = yaml.safe_load(spec_yaml.read_text(encoding="utf-8"))
                if spec_data:
                    spec_pages = [{
                        "page_type": "content",
                        "background_color": spec_data.get("colors", {}).get("background1", "#FFFFFF"),
                        "colors": spec_data.get("colors", {}),
                        "typography": spec_data.get("typography", {}),
                        "regions": [],
                        "background_description": "Extracted from spec",
                    }]
            except Exception:
                pass

    # Load metadata for non-VL formats
    metadata: dict = {}
    spec_yaml = spec_dir / "spec.yaml"
    if spec_yaml.exists():
        try:
            metadata = yaml.safe_load(spec_yaml.read_text(encoding="utf-8")) or {}
        except Exception:
            pass

    return spec_pages, metadata


def _resolve_bg_color(background: dict) -> str:
    """Resolve background info to a single hex color string."""
    if not background:
        return "#FFFFFF"
    bg_type = background.get("type", "solid")
    if bg_type == "solid":
        return background.get("color", "#FFFFFF")
    if bg_type == "image":
        return "#FFFFFF"  # image background, use white as base
    if bg_type == "gradient":
        return background.get("gradient_stops", [{}])[0].get("color", "#FFFFFF") if background.get("gradient_stops") else "#FFFFFF"
    return "#FFFFFF"


def _load_content_outline(outline_path: Path) -> dict:
    """Load a content outline YAML file."""
    data = yaml.safe_load(outline_path.read_text(encoding="utf-8"))
    return data or {}


def _classify_slide_type(slide: dict, slide_number: int, total: int) -> str:
    """Determine the target page type for a content slide."""
    layout = slide.get("layout_type", "content")

    # First slide → cover
    if slide_number == 1:
        return "cover"

    # Section divider layouts → transition
    if layout in ("section_divider", "section_header"):
        return "transition"

    # Last slide → end_page
    if slide_number == total:
        return "end_page"

    # Content layouts
    return "content"


def generate_pptx(
    spec_dir: Path,
    outline_path: Path,
    output_path: Path,
    generate_callback,
    max_workers: int = MAX_WORKERS,
) -> GenerationResult:
    """Generate a complete PPTX from spec + content outline.

    Uses ThreadPoolExecutor for parallel slide generation with
    agent-loop style evaluation on each slide.

    Args:
        spec_dir: Path to spec directory (specs/<name>/).
        outline_path: Path to content outline YAML file.
        output_path: Destination .pptx file path.
        generate_callback: Function(prompt_str) -> SVG text.
                           The AI runtime implements this.
        max_workers: Number of parallel workers.

    Returns:
        GenerationResult with per-slide details.
    """
    # Load spec
    spec_pages, spec_metadata = _load_spec_pages(spec_dir)
    if not spec_pages:
        raise ValueError(f"No spec pages found in {spec_dir}")

    # Load outline
    outline = _load_content_outline(outline_path)
    slides: list[dict] = outline.get("slides", [])
    if not slides:
        raise ValueError(f"No slides found in outline: {outline_path}")

    total_slides = len(slides)
    result = GenerationResult()
    futures: dict = {}

    # Classify each slide and prepare tasks
    tasks: list[dict] = []
    for i, slide in enumerate(slides):
        slide_number = i + 1
        page_type = _classify_slide_type(slide, slide_number, total_slides)
        task = dict(slide)
        task["slide_number"] = slide_number
        task["page_type"] = page_type
        task["is_end"] = (slide_number == total_slides)
        tasks.append(task)

    # Parallel generation
    def generate_one(task: dict) -> SlideResult:
        return generate_slide_with_loop(
            slide_entry=task,
            spec_pages=spec_pages,
            spec_metadata=spec_metadata,
            generate_callback=generate_callback,
        )

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_index = {}
        for task in tasks:
            future = executor.submit(generate_one, task)
            future_to_index[future] = task["slide_number"]

        for future in as_completed(future_to_index):
            slide_num = future_to_index[future]
            try:
                slide_result = future.result()
                result.slide_results.append(slide_result)
            except Exception as e:
                slide_result = SlideResult(
                    slide_index=slide_num,
                    passed=False,
                )
                slide_result.reports.append(
                    StyleReport(issues=[f"Generation error: {e}"])
                )
                result.slide_results.append(slide_result)

    # Sort by slide number
    result.slide_results.sort(key=lambda r: r.slide_index)

    # Phase 2: Layout Design — LLM-driven agent loop using prompt-ppt-layout.md
    designer = LayoutDesigner()
    for i, sr in enumerate(result.slide_results):
        if sr.passed:
            slide_content = {
                "w": tasks[i].get("section_name", ""),
                "p": tasks[i].get("title", ""),
                "s": tasks[i].get("body", []),
            }
            spec_page = next(
                (p for p in spec_pages if p.get("page_type") == sr.page_type),
                {}
            )
            layout_result = designer.design(slide_content, spec_page, generate_callback)
            result.layout_results.append(layout_result)
            if layout_result.passed:
                result.layout_passed_count += 1
            result.layout_iterations += layout_result.iterations
        else:
            result.layout_results.append(LayoutResult(passed=False))

    # Collect SVGs and convert to PPTX
    svg_paths: list[Path] = []
    tmp_dir = output_path.parent / ".tmp_svgs"
    tmp_dir.mkdir(parents=True, exist_ok=True)

    for sr in result.slide_results:
        if sr.passed:
            result.passed_count += 1
        else:
            result.failed_count += 1

        result.total_iterations += sr.iterations
        result.avg_score += sr.best_score

        if sr.svg_text:
            svg_path = tmp_dir / f"slide_{sr.slide_index:02d}.svg"
            svg_path.write_text(sr.svg_text, encoding="utf-8")
            svg_paths.append(svg_path)

    if result.slide_results:
        result.avg_score /= len(result.slide_results)

    # Convert to PPTX
    if svg_paths:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        svg_paths.sort()
        prs.save(str(output_path))
        result.output_path = output_path

    return result
