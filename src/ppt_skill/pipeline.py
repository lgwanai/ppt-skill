"""PPT Skill — unified SVG-to-PPTX pipeline."""

import argparse
from pathlib import Path

from ppt_skill.converter.builder import create_pptx_with_native_svg
from ppt_skill.quality import SVGQualityChecker


def convert_svg_to_pptx(
    svg_files: list[Path],
    output_path: Path,
    *,
    skip_quality_check: bool = False,
) -> bool:
    """
    Convert SVG files to native-shape PPTX with quality validation.

    Args:
        svg_files: List of SVG file paths to convert (one per slide)
        output_path: Destination .pptx file path
        skip_quality_check: If True, bypass quality validation (for trusted SVGs)

    Returns:
        True if conversion succeeded

    Raises:
        ValueError: If quality check fails and skip_quality_check is False
        FileNotFoundError: If any SVG file does not exist
    """
    # 1. Validate inputs
    for svg_path in svg_files:
        if not svg_path.exists():
            raise FileNotFoundError(f"SVG file not found: {svg_path}")

    # 2. Quality check (unless skipped)
    if not skip_quality_check:
        checker = SVGQualityChecker()
        for svg_path in svg_files:
            result = checker.check_file(str(svg_path))
            if not result.get("passed", False):
                errors = result.get("errors", [])
                raise ValueError(
                    f"SVG quality check failed for {svg_path}:\n"
                    + "\n".join(f"  - {e}" for e in errors)
                )

    # 3. Convert to PPTX with native shapes
    return create_pptx_with_native_svg(
        svg_files=svg_files,
        output_path=output_path,
        use_native_shapes=True,
        use_compat_mode=False,
        verbose=False,
    )


def main() -> None:
    """CLI entry point: python -m ppt_skill.pipeline --input ... --output ..."""
    parser = argparse.ArgumentParser(
        description="Convert SVG files to native-shape PowerPoint slides"
    )
    parser.add_argument(
        "--input",
        "-i",
        nargs="+",
        required=True,
        type=Path,
        help="One or more SVG files to convert (one per slide)",
    )
    parser.add_argument(
        "--output",
        "-o",
        required=True,
        type=Path,
        help="Output .pptx file path",
    )
    parser.add_argument(
        "--skip-check",
        action="store_true",
        help="Skip SVG quality validation",
    )

    args = parser.parse_args()

    try:
        convert_svg_to_pptx(
            svg_files=list(args.input),
            output_path=args.output,
            skip_quality_check=args.skip_check,
        )
        print(f"PPTX created: {args.output}")
    except (ValueError, FileNotFoundError) as e:
        print(f"Error: {e}", flush=True)
        exit(1)


if __name__ == "__main__":
    main()
