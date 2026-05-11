#!/usr/bin/env python3
"""
PPT Skill — CLI entry point.

Unified command-line interface for the PPT generation skill:
  - convert: SVG → native-shape PPTX conversion
  - extract-spec: Extract design spec from reference PPTX
  - extract-vl-spec: Extract VL-driven spec from PPTX (JSON, no text, layout dedup)
  - outline: Generate content outline using WPS model (prompt-ppt-content.md)
  - gather-content: Content questioning → outline generation
  - generate-pptx: Generate PPTX from outline + spec (two-phase: spec matching + layout design)

Usage:
  python scripts/ppt_cli.py convert input.svg -o output.pptx
  python scripts/ppt_cli.py extract-spec reference.pptx
  python scripts/ppt_cli.py outline "AI in Healthcare" -o outline.md
  python scripts/ppt_cli.py generate-pptx --spec specs/my_spec --outline outline.yaml -o output.pptx
"""

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from ppt_skill.pipeline import convert_svg_to_pptx


def cmd_convert(args: argparse.Namespace) -> int:
    """Convert SVG files to native-shape PPTX."""
    svg_files = [Path(f) for f in args.input]
    output = Path(args.output)

    missing = [f for f in svg_files if not f.exists()]
    if missing:
        print(f"Error: SVG files not found: {missing}", file=sys.stderr)
        return 1

    try:
        ok = convert_svg_to_pptx(
            svg_files,
            output,
            skip_quality_check=args.skip_check,
        )
        if ok:
            print(f"PPTX created: {output} ({len(svg_files)} slides)")
        return 0
    except ValueError as e:
        print(f"Quality check failed: {e}", file=sys.stderr)
        return 1


def cmd_extract_spec(args: argparse.Namespace) -> int:
    """Extract design specification from a reference PPTX file.
    
    Outputs a directory-based spec with page-level layout analysis,
    assets, and presentation logic. If VL_ENABLED=true in config.txt,
    uses vision model for enhanced layout description.
    """
    from ppt_skill.spec.enhanced_extractor import SpecExtractor

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Error: PPTX file not found: {input_path}", file=sys.stderr)
        return 1

    extractor = SpecExtractor()
    spec = extractor.extract(input_path)
    extractor.save(spec)

    print(f"Spec saved: specs/{spec.metadata.get('name', 'spec')}/")
    print(f"  Colors: {len(spec.palette.to_dict())} palette entries")
    print(f"  Fonts: {spec.typography.heading_family or 'unknown'} / {spec.typography.body_family or 'unknown'}")
    print(f"  Pages: {len(spec.pages)} → {', '.join(spec.page_types_found)}")
    if spec.layout_sub_types_found:
        print(f"  Layouts: {', '.join(spec.layout_sub_types_found)}")
    print(f"  VL analysis: {'enabled' if extractor.config.enabled else 'disabled'}")
    return 0


def cmd_extract_vl_spec(args: argparse.Namespace) -> int:
    """Extract VL-driven design spec from a reference PPTX file.

    Uses VL model to analyze element roles and relationships, producing
    deduplicated JSON spec files with NO text content — only attributes,
    roles, and properties.

    Output: specs/<name>/ with:
      - spec.json (master spec with palette, fonts, canvas)
      - slides/ (slide PNGs)
      - cover.json, end_page.json, content_*.json (layout blueprints)
    """
    from ppt_skill.spec.vl_spec_extractor import VLVSpecExtractor

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Error: PPTX file not found: {input_path}", file=sys.stderr)
        return 1

    extractor = VLVSpecExtractor()
    extractor.extract(input_path, output_name=args.name)
    return 0


def cmd_list_specs(args: argparse.Namespace) -> int:
    """List all available design specs."""
    from ppt_skill.cli.spec_commands import list_specs, get_active_spec

    specs = list_specs()
    active = get_active_spec()

    if not specs:
        print("No specs available. Use 'extract-spec' to create one from a reference PPTX.")
        return 0

    print(f"{'Active':<8} {'Name':<30} {'Slides':<8}")
    print("-" * 50)
    for spec_path in specs:
        marker = "★" if spec_path.name.replace(".yaml", "") == active else ""
        print(f"{marker:<8} {spec_path.stem:<30}")
    return 0


def cmd_select_spec(args: argparse.Namespace) -> int:
    """Select an active design spec."""
    from ppt_skill.cli.spec_commands import select_spec

    select_spec(args.name)
    print(f"Active spec set to: {args.name}")
    return 0


def cmd_gather_content(args: argparse.Namespace) -> int:
    """Gather content through adaptive questioning and generate outline."""
    from ppt_skill.content.gatherer import ContentGatherer
    from ppt_skill.content.model import ContentOutline, _dataclass_to_dict

    input_text = args.input
    if Path(args.input).exists():
        input_text = Path(args.input).read_text()

    gatherer = ContentGatherer()
    result = gatherer.gather(input_text, mode=args.mode)

    if isinstance(result, ContentOutline):
        outline = result
    else:
        outline = result.get("outline")

    if not outline or not outline.slides:
        print("Error: No content generated", file=sys.stderr)
        return 1

    issues = outline.validate()
    if issues:
        print(f"Warning: Outline has {len(issues)} validation issues:", file=sys.stderr)
        for issue in issues:
            print(f"  - {issue}", file=sys.stderr)

    # Output as Markdown following prompt-ppt-content.md WPS format
    print(outline.to_ppt_markdown())

    # Save to outlines/ directory as markdown
    path = gatherer.save("outlines")
    print(f"\nOutline saved to: {path}", file=sys.stderr)

    return 0


def cmd_outline(args: argparse.Namespace) -> int:
    """Generate PPT content outline using WPS model and prompt-ppt-content.md principles.

    Uses ContentGatherer to generate a proper outline from rich content.
    Output follows the WPS markdown format specified in references/prompt-ppt-content.md.
    Topic-only input (<=30 words, no structure) outputs a delegation signal.
    """
    from ppt_skill.content.outline_generator import (
        is_topic_only, build_delegation_signal,
    )
    from ppt_skill.content.gatherer import ContentGatherer

    input_text = args.input
    if Path(args.input).exists():
        input_text = Path(args.input).read_text()

    # Detect topic-only input and output delegation signal
    if is_topic_only(input_text):
        signal = build_delegation_signal(input_text)
        print(signal)
        return 0

    # Use ContentGatherer for proper outline generation
    gatherer = ContentGatherer()
    outline = gatherer.gather(input_text, mode=args.mode or "skip_questions")

    issues = outline.validate()
    if issues:
        print(f"Warning: Outline has {len(issues)} validation issues:", file=sys.stderr)
        for issue in issues:
            print(f"  - {issue}", file=sys.stderr)

    # Output as Markdown following prompt-ppt-content.md WPS format
    print(outline.to_ppt_markdown())

    # Always save to outlines/ directory
    path = gatherer.save("outlines")
    print(f"\nOutline saved to: {path}", file=sys.stderr)

    return 0


def cmd_list_outlines(args: argparse.Namespace) -> int:
    """List all saved content outlines."""
    outlines_dir = Path("outlines")
    if not outlines_dir.exists():
        print("No outlines directory found.")
        return 0

    outlines = sorted(outlines_dir.glob("*.yaml"))
    if not outlines:
        print("No outlines found.")
        return 0

    for op in outlines:
        print(f"  {op.stem}")
    return 0


def cmd_generate_pptx(args: argparse.Namespace) -> int:
    """Generate PPTX from spec + content outline with agent-loop evaluation."""
    from ppt_skill.ppt_generator import generate_pptx

    spec_dir = Path(args.spec) if args.spec else None
    outline_path = Path(args.outline) if args.outline else None

    if not spec_dir or not spec_dir.exists():
        print(f"Error: Spec directory not found: {spec_dir}", file=sys.stderr)
        return 1
    if not outline_path or not outline_path.exists():
        print(f"Error: Outline file not found: {outline_path}", file=sys.stderr)
        return 1

    output = Path(args.output) if args.output else Path("output.pptx")
    workers = int(args.workers) if args.workers else 4

    print(f"Generating PPTX from spec: {spec_dir.name}")
    print(f"Outline: {outline_path.name}")
    print(f"Workers: {workers}")
    print(f"Agent loop: max {5} iterations, threshold 90% style match")

    # For CLI standalone mode, the generate_callback is a placeholder.
    # In AI runtime (opencode/claude-code), the LLM handles generation.
    def generate_callback(prompt: str) -> str:
        import warnings
        warnings.warn(
            "generate-pptx in CLI mode uses placeholder SVG. "
            "Use AI runtime for LLM-backed generation.",
            RuntimeWarning,
        )
        return (
            '<svg viewBox="0 0 1280 720" xmlns="http://www.w3.org/2000/svg">'
            '<rect width="1280" height="720" fill="#FFFFFF"/>'
            '<text x="100" y="100" font-family="Arial" font-size="24">Generated Slide</text>'
            '</svg>'
        )

    try:
        result = generate_pptx(
            spec_dir=spec_dir,
            outline_path=outline_path,
            output_path=output,
            generate_callback=generate_callback,
            max_workers=workers,
        )

        print(f"\nGeneration complete:")
        print(f"  Passed: {result.passed_count}/{len(result.slide_results)}")
        print(f"  Avg score: {result.avg_score:.1%}")
        print(f"  Total iterations: {result.total_iterations}")
        if result.output_path:
            print(f"  Output: {result.output_path}")

        for sr in result.slide_results:
            status = "✓" if sr.passed else "⚠"
            print(f"  {status} Slide {sr.slide_index}: "
                  f"{sr.page_type}/{sr.layout_sub_type} "
                  f"score={sr.best_score:.0%} "
                  f"iter={sr.iterations}")

        return 0 if result.passed_count == len(result.slide_results) else 1

    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1


def main() -> int:
    parser = argparse.ArgumentParser(
        description="PPT Skill — AI-powered presentation generation",
        prog="ppt-skill",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    # convert
    p_convert = sub.add_parser("convert", help="Convert SVG to PPTX")
    p_convert.add_argument("input", nargs="+", help="SVG file(s)")
    p_convert.add_argument("-o", "--output", required=True, help="Output PPTX path")
    p_convert.add_argument("--skip-check", action="store_true", help="Skip quality check")

    # extract-spec
    p_extract = sub.add_parser("extract-spec", help="Extract design spec from PPTX")
    p_extract.add_argument("input", help="Reference PPTX file")
    p_extract.add_argument("--name", help="Spec name (default: filename stem)")

    # extract-vl-spec
    p_vl_extract = sub.add_parser("extract-vl-spec", help="Extract VL-driven spec from PPTX (JSON, no text)")
    p_vl_extract.add_argument("input", help="Reference PPTX file")
    p_vl_extract.add_argument("--name", help="Spec name (default: auto-generated from filename)")

    # list-specs
    sub.add_parser("list-specs", help="List available design specs")

    # select-spec
    p_select = sub.add_parser("select-spec", help="Select active design spec")
    p_select.add_argument("name", help="Spec name")

    # gather-content
    p_gather = sub.add_parser("gather-content", help="Gather content and generate outline")
    p_gather.add_argument("input", help="Content text or file path")
    p_gather.add_argument("--mode", choices=["assess", "question", "skip_questions"],
                          default="assess", help="Gathering mode")

    # outline (ppt-outline)
    p_outline = sub.add_parser("outline", help="Generate PPT content outline using WPS model")
    p_outline.add_argument("input", help="Topic, article, or existing outline")
    p_outline.add_argument("--mode", choices=["auto", "skip_questions"],
                           default="auto", help="Generation mode")
    p_outline.add_argument("-o", "--output", help="Save outline to file (.md or .yaml)")

    # list-outlines
    sub.add_parser("list-outlines", help="List saved content outlines")

    # generate-pptx
    p_gen = sub.add_parser("generate-pptx", help="Generate PPTX from spec + outline (agent-loop)")
    p_gen.add_argument("--spec", required=True, help="Spec directory (specs/<name>/)")
    p_gen.add_argument("--outline", required=True, help="Content outline YAML file")
    p_gen.add_argument("-o", "--output", default="output.pptx", help="Output PPTX path")
    p_gen.add_argument("--workers", type=int, default=4, help="Parallel workers (default: 4)")

    args = parser.parse_args()

    dispatch = {
        "convert": cmd_convert,
        "extract-spec": cmd_extract_spec,
        "extract-vl-spec": cmd_extract_vl_spec,
        "list-specs": cmd_list_specs,
        "select-spec": cmd_select_spec,
        "gather-content": cmd_gather_content,
        "outline": cmd_outline,
        "list-outlines": cmd_list_outlines,
        "generate-pptx": cmd_generate_pptx,
    }

    handler = dispatch.get(args.command)
    if handler:
        return handler(args)

    print(f"Unknown command: {args.command}", file=sys.stderr)
    return 1


if __name__ == "__main__":
    sys.exit(main())
