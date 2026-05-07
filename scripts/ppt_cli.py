#!/usr/bin/env python3
"""
PPT Skill — CLI entry point.

Unified command-line interface for the PPT generation skill:
  - convert: SVG → native-shape PPTX conversion
  - extract-spec: Extract design spec from reference PPTX
  - gather-content: Content questioning → outline generation
  - generate-pptx: Generate PPTX from outline + spec (Phase 4 placeholder)

Usage:
  python scripts/ppt_cli.py convert input.svg -o output.pptx
  python scripts/ppt_cli.py extract-spec reference.pptx
  python scripts/ppt_cli.py gather-content "topic or content text"
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
    print(f"  Colors: {len(spec.colors.to_dict())} palette entries")
    print(f"  Fonts: {spec.typography.heading_family or 'unknown'} / {spec.typography.body_family or 'unknown'}")
    print(f"  Pages: {len(spec.pages)} → {', '.join(spec.page_types_found)}")
    if spec.layout_sub_types_found:
        print(f"  Layouts: {', '.join(spec.layout_sub_types_found)}")
    print(f"  Assets: {spec.asset_count}")
    print(f"  VL analysis: {'enabled' if extractor.config.enabled else 'disabled'}")
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
    from ppt_skill.content.model import ContentOutline

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
        print(f"Warning: Outline has {len(issues)} validation issues:")
        for issue in issues:
            print(f"  - {issue}")

    yaml_str = outline.to_yaml()
    print(yaml_str)
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

    # list-outlines
    sub.add_parser("list-outlines", help="List saved content outlines")

    args = parser.parse_args()

    dispatch = {
        "convert": cmd_convert,
        "extract-spec": cmd_extract_spec,
        "list-specs": cmd_list_specs,
        "select-spec": cmd_select_spec,
        "gather-content": cmd_gather_content,
        "list-outlines": cmd_list_outlines,
    }

    handler = dispatch.get(args.command)
    if handler:
        return handler(args)

    print(f"Unknown command: {args.command}", file=sys.stderr)
    return 1


if __name__ == "__main__":
    sys.exit(main())
