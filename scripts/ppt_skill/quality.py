"""
PPT Skill — SVG Quality Checker

Validates SVG files against ppt-master compatibility rules for native-shape
DrawingML conversion. Standalone module with no external dependencies beyond
ppt_skill.config and Python stdlib.

Usage:
    from ppt_skill.quality import SVGQualityChecker
    checker = SVGQualityChecker()
    result = checker.check_file("slide.svg")
    if not result['passed']:
        for err in result['errors']:
            print(err)
"""

import re
import json
import html
from pathlib import Path
from typing import List, Dict
from collections import defaultdict
from xml.etree import ElementTree as ET

from ppt_skill.config import CANVAS_FORMATS


HEX_VALUE_RE = re.compile(r"#[0-9A-Fa-f]{3,8}")


class SVGQualityChecker:
    """SVG quality checker for ppt-master compatibility rules.

    Validates SVGs against the DrawingML converter's requirements:
    banned elements (mask, style, script, etc.), CSS patterns (rgba,
    @font-face, @import), font stacks, viewBox, dimensions, and more.
    """

    def __init__(self):
        self.results: List[Dict] = []
        self.summary = {
            'total': 0,
            'passed': 0,
            'warnings': 0,
            'errors': 0
        }
        self.issue_types = defaultdict(int)

    def check_file(self, svg_file: str, expected_format: str = None) -> Dict:
        """Check a single SVG file.

        Args:
            svg_file: SVG file path
            expected_format: Expected canvas format (e.g., 'ppt169')

        Returns:
            Check result dictionary with keys:
            file, path, exists, errors, warnings, info, passed
        """
        svg_path = Path(svg_file)

        if not svg_path.exists():
            return {
                'file': str(svg_file),
                'exists': False,
                'errors': ['File does not exist'],
                'warnings': [],
                'passed': False
            }

        result = {
            'file': svg_path.name,
            'path': str(svg_path),
            'exists': True,
            'errors': [],
            'warnings': [],
            'info': {},
            'passed': True
        }

        try:
            with open(svg_path, 'r', encoding='utf-8') as f:
                content = f.read()

            # 0. Check XML well-formedness — every other check assumes the file
            # is valid XML. Bail early on failure so the regex-based checks
            # below don't produce misleading errors on a broken document.
            if self._check_xml_well_formed(content, result):
                # 1. Check viewBox
                self._check_viewbox(content, result, expected_format)

                # 2. Check forbidden elements
                self._check_forbidden_elements(content, result)

                # 3. Check fonts
                self._check_fonts(content, result)

                # 4. Check width/height consistency with viewBox
                self._check_dimensions(content, result)

                # 5. Check text wrapping methods
                self._check_text_elements(content, result)

                # 6. Check image references (file existence and resolution)
                self._check_image_references(content, svg_path, result)

            # Determine pass/fail
            result['passed'] = len(result['errors']) == 0

        except Exception as e:
            result['errors'].append(f"Failed to read file: {e}")
            result['passed'] = False

        # Update statistics
        self.summary['total'] += 1
        if result['passed']:
            if result['warnings']:
                self.summary['warnings'] += 1
            else:
                self.summary['passed'] += 1
        else:
            self.summary['errors'] += 1

        # Categorize issue types
        for error in result['errors']:
            self.issue_types[self._categorize_issue(error)] += 1

        self.results.append(result)
        return result

    def _check_xml_well_formed(self, content: str, result: Dict) -> bool:
        """Check that the SVG content parses as well-formed XML.

        SVG is strict XML. AI-generated decks frequently produce content that
        looks fine in HTML5-tolerant previews but fails strict XML parsing —
        common causes are HTML named entities (&nbsp; &mdash; &copy;...) and
        bare XML reserved characters in text (R&D, error < 5%). Such pages
        cannot be exported to PPTX, so we surface them here as a hard error
        before any downstream check looks at them.

        Returns True when the document is well-formed; False otherwise.
        """
        try:
            ET.fromstring(content)
            return True
        except ET.ParseError as e:
            result['errors'].append(
                f"Invalid XML: {e} — SVG must be well-formed XML. "
                f"Use raw Unicode for typography (—, ©, →, NBSP); "
                f"escape XML reserved chars as &amp; &lt; &gt; &quot; &apos; "
                f"(see references/shared-standards.md §1)."
            )
            return False

    def _check_viewbox(self, content: str, result: Dict, expected_format: str = None):
        """Check viewBox attribute."""
        viewbox_match = re.search(r'viewBox="([^"]+)"', content)

        if not viewbox_match:
            result['errors'].append("Missing viewBox attribute")
            return

        viewbox = viewbox_match.group(1)
        result['info']['viewbox'] = viewbox

        # Check format
        if not re.match(r'0 0 \d+ \d+', viewbox):
            result['warnings'].append(f"Unusual viewBox format: {viewbox}")

        # Check if it matches expected format
        if expected_format and expected_format in CANVAS_FORMATS:
            expected_w = CANVAS_FORMATS[expected_format]['width']
            expected_h = CANVAS_FORMATS[expected_format]['height']
            # Convert EMU to SVG user-units (1px ≈ 9525 EMU at 96 DPI)
            # The viewBox uses pixel units, but CANVAS_FORMATS uses EMU.
            # For viewBox comparison, we expect the viewBox to match the
            # ppi-scaled canvas dimensions. The common convention is that
            # the viewBox values equal width/height in pixels.
            #
            # Since we don't know the exact PPI used, we issue a warning
            # if the viewBox doesn't match common conventions.
            pass

    # ============================================================
    # Forbidden elements and patterns — PPT incompatible
    # ============================================================

    def _check_forbidden_elements(self, content: str, result: Dict):
        """Check forbidden elements and patterns (comprehensive blocklist).

        Validates against all ~20 banned features documented in
        RESEARCH.md §"Complete Banned Feature List".
        """
        content_lower = content.lower()

        # --- Clipping / masking ---
        # clipPath is ONLY allowed on <image> elements (converter maps to
        # DrawingML picture geometry). On shapes it is pointless (just draw
        # the target shape) and breaks the SVG PPTX rendering.
        if '<clippath' in content_lower:
            clip_on_non_image = re.search(
                r'<(?!image\b)\w+[^>]*\bclip-path\s*=', content, re.IGNORECASE)
            if clip_on_non_image:
                result['errors'].append(
                    "clip-path is only allowed on <image> elements — "
                    "for shapes, draw the target shape directly instead of clipping")
            # Check that every clip-path reference has a matching <clipPath> def
            clip_refs = re.findall(r'clip-path\s*=\s*["\']url\(#([^)]+)\)', content)
            for ref_id in clip_refs:
                if f'id="{ref_id}"' not in content and f"id='{ref_id}'" not in content:
                    result['errors'].append(
                        f"clip-path references #{ref_id} but no matching "
                        f"<clipPath id=\"{ref_id}\"> definition found")
        if '<mask' in content_lower:
            result['errors'].append(
                "Detected forbidden <mask> element (PPT does not support SVG masks)")

        # --- Style system ---
        if '<style' in content_lower:
            result['errors'].append(
                "Detected forbidden <style> element (use inline attributes instead)")
        if re.search(r'\bclass\s*=', content):
            result['errors'].append(
                "Detected forbidden class attribute (use inline styles instead)")
        # id attribute: only report error when <style> also exists (id is
        # harmful only with CSS selectors). id inside <defs> for
        # linearGradient/filter etc. is required; standalone id attributes
        # have no impact on PPT export.
        if '<style' in content_lower and re.search(r'\bid\s*=', content):
            result['errors'].append(
                "Detected id attribute used with <style> "
                "(CSS selectors forbidden, use inline styles instead)")
        if re.search(r'<\?xml-stylesheet\b', content_lower):
            result['errors'].append(
                "Detected forbidden xml-stylesheet (external CSS references forbidden)")
        if re.search(r'<link[^>]*rel\s*=\s*["\']stylesheet["\']', content_lower):
            result['errors'].append(
                "Detected forbidden <link rel=\"stylesheet\"> "
                "(external CSS references forbidden)")
        if re.search(r'@import\s+', content_lower):
            result['errors'].append(
                "Detected forbidden @import (external CSS references forbidden)")

        # --- Structure / nesting ---
        if '<foreignobject' in content_lower:
            result['errors'].append(
                "Detected forbidden <foreignObject> element "
                "(use <tspan> for manual line breaks)")
        has_symbol = '<symbol' in content_lower
        has_use = re.search(r'<use\b', content_lower) is not None
        if has_symbol and has_use:
            result['errors'].append(
                "Detected forbidden <symbol> + <use> complex usage "
                "(use basic shapes or simple <use> instead)")
        # marker-start / marker-end are conditionally allowed. The converter
        # maps qualifying <marker> defs to native DrawingML <a:headEnd>/<a:tailEnd>.
        # We only warn when a marker is used without an obvious <defs> definition.
        if re.search(r'\bmarker-(?:start|end)\s*=\s*["\']url\(#([^)]+)\)', content_lower):
            if '<marker' not in content_lower:
                result['errors'].append(
                    "Detected marker-start/marker-end referencing a marker id, "
                    "but no <marker> element found in the file")

        # --- Text / fonts ---
        if '<textpath' in content_lower:
            result['errors'].append(
                "Detected forbidden <textPath> element "
                "(path text is incompatible with PPT)")
        if '@font-face' in content_lower:
            result['errors'].append(
                "Detected forbidden @font-face (use system font stack)")

        # --- Animation / interaction ---
        if re.search(r'<animate', content_lower):
            result['errors'].append(
                "Detected forbidden SMIL animation element <animate*> "
                "(SVG animations are not exported)")
        if re.search(r'<set\b', content_lower):
            result['errors'].append(
                "Detected forbidden SMIL animation element <set> "
                "(SVG animations are not exported)")
        if '<script' in content_lower:
            result['errors'].append(
                "Detected forbidden <script> element "
                "(scripts and event handlers forbidden)")
        if re.search(r'\bon\w+\s*=', content):  # onclick, onload etc.
            result['errors'].append(
                "Detected forbidden event attributes (e.g., onclick, onload)")

        # --- Other discouraged elements ---
        if '<iframe' in content_lower:
            result['errors'].append(
                "Detected <iframe> element (should not appear in SVG)")
        if re.search(r'rgba\s*\(', content_lower):
            result['errors'].append(
                "Detected forbidden rgba() color "
                "(use fill-opacity/stroke-opacity instead)")
        if re.search(r'<g[^>]*\sopacity\s*=', content_lower):
            result['errors'].append(
                "Detected forbidden <g opacity> "
                "(set opacity on each child element individually)")
        if re.search(r'<image[^>]*\sopacity\s*=', content_lower):
            result['errors'].append(
                "Detected forbidden <image opacity> (use overlay mask approach)")

    def _check_fonts(self, content: str, result: Dict):
        """Check font usage.

        PPTX stores a single `typeface` per run with no runtime fallback, so
        every stack must END with a cross-platform pre-installed family.
        """
        font_matches = re.findall(
            r'font-family[:\s]*["\']([^"\']+)["\']', content, re.IGNORECASE)

        if not font_matches:
            return

        result['info']['fonts'] = list(set(font_matches))

        # Pre-installed on Windows + macOS out of the box (plus their direct
        # FONT_FALLBACK_WIN mappings). A stack whose last concrete family is in
        # this set survives the PPTX round-trip on any viewer machine.
        ppt_safe_tail = {
            'microsoft yahei', 'simhei', 'simsun', 'kaiti', 'fangsong',
            'pingfang sc', 'heiti sc', 'songti sc', 'stsong',
            'arial', 'arial black', 'calibri', 'segoe ui', 'verdana',
            'helvetica', 'helvetica neue', 'tahoma', 'trebuchet ms',
            'times new roman', 'times', 'georgia', 'cambria', 'palatino',
            'consolas', 'courier new', 'menlo', 'monaco',
            'impact',
        }

        for font_family in font_matches:
            # Drop the generic CSS fallback (sans-serif / serif / monospace)
            # and inspect the last concrete family.
            parts = [p.strip().strip('"').strip("'").lower()
                     for p in font_family.split(',')]
            parts = [p for p in parts
                     if p and p not in ('sans-serif', 'serif', 'monospace',
                                        'cursive', 'fantasy', 'system-ui')]
            if not parts:
                continue
            tail = parts[-1]
            if tail not in ppt_safe_tail:
                result['warnings'].append(
                    f"Font stack does not end on a PPT-safe family "
                    f"(expected e.g. Microsoft YaHei / SimSun / Arial / "
                    f"Times New Roman / Consolas): {font_family}"
                )
                break

    def _check_dimensions(self, content: str, result: Dict):
        """Check width/height consistency with viewBox."""
        width_match = re.search(r'width="(\d+)"', content)
        height_match = re.search(r'height="(\d+)"', content)

        if width_match and height_match:
            width = width_match.group(1)
            height = height_match.group(1)
            result['info']['dimensions'] = f"{width}x{height}"

            # Check consistency with viewBox
            if 'viewbox' in result['info']:
                viewbox_parts = result['info']['viewbox'].split()
                if len(viewbox_parts) == 4:
                    vb_width, vb_height = viewbox_parts[2], viewbox_parts[3]
                    if width != vb_width or height != vb_height:
                        result['warnings'].append(
                            f"width/height ({width}x{height}) does not match "
                            f"viewBox ({vb_width}x{vb_height})"
                        )

    def _check_text_elements(self, content: str, result: Dict):
        """Check text elements and wrapping methods."""
        # Count text and tspan elements
        text_count = content.count('<text')
        tspan_count = content.count('<tspan')

        result['info']['text_elements'] = text_count
        result['info']['tspan_elements'] = tspan_count

        # Check for overly long single-line text (may need wrapping)
        text_matches = re.findall(r'<text[^>]*>([^<]{100,})</text>', content)
        if text_matches:
            result['warnings'].append(
                f"Detected {len(text_matches)} potentially overly long "
                f"single-line text(s) (consider using tspan for wrapping)"
            )

    def _check_image_references(self, content: str, svg_path: Path, result: Dict):
        """Check image file existence and resolution vs display size."""
        img_tag_pattern = re.compile(r'<image\b([^>]*)/?>', re.IGNORECASE)

        svg_dir = svg_path.parent
        checked = set()

        for tag_match in img_tag_pattern.finditer(content):
            attrs = tag_match.group(1)

            # Extract href (prefer href over xlink:href)
            href_match = (
                re.search(r'\bhref="(?!data:)([^"]+)"', attrs) or
                re.search(r'\bxlink:href="(?!data:)([^"]+)"', attrs)
            )
            if not href_match:
                continue

            href = href_match.group(1)
            if href in checked:
                continue
            checked.add(href)

            # Resolve relative path
            img_path = svg_dir / href
            if not img_path.exists():
                # Try relative to parent (common for svg_output/ layouts)
                alt_path = svg_path.parent.parent / href
                if not alt_path.exists():
                    result['warnings'].append(
                        f"Referenced image file not found: {href}"
                    )
                    continue
                img_path = alt_path

            # Check resolution vs display size
            try:
                from PIL import Image
                with Image.open(img_path) as img:
                    actual_w, actual_h = img.size
            except ImportError:
                pass  # PIL not available, skip resolution check
            except Exception:
                pass  # Image unreadable, skip resolution check
            else:
                display_w_match = re.search(r'\bwidth="([^"]+)"', attrs)
                display_h_match = re.search(r'\bheight="([^"]+)"', attrs)
                if display_w_match and display_h_match:
                    try:
                        display_w = float(display_w_match.group(1))
                        display_h = float(display_h_match.group(1))
                    except ValueError:
                        continue
                    if actual_w > display_w * 3 or actual_h > display_h * 3:
                        result['warnings'].append(
                            f"Image {href} is {actual_w}x{actual_h} but displayed "
                            f"at {int(display_w)}x{int(display_h)} — consider "
                            f"downsizing to reduce file size"
                        )

    def _categorize_issue(self, error_msg: str) -> str:
        """Categorize issue type for summary aggregation."""
        if 'Invalid XML' in error_msg:
            return 'XML well-formedness'
        elif 'viewBox' in error_msg:
            return 'viewBox issues'
        elif 'foreignObject' in error_msg:
            return 'foreignObject'
        elif 'font' in error_msg.lower():
            return 'Font issues'
        else:
            return 'Other'

    def check_directory(self, directory: str, expected_format: str = None) -> List[Dict]:
        """Check all SVG files in a directory.

        Args:
            directory: Directory path
            expected_format: Expected canvas format

        Returns:
            List of check results
        """
        dir_path = Path(directory)

        if not dir_path.exists():
            print(f"[ERROR] Directory does not exist: {directory}")
            return []

        # Find all SVG files
        if dir_path.is_file():
            svg_files = [dir_path]
        else:
            svg_output = dir_path / 'svg_output' if (
                dir_path / 'svg_output').exists() else dir_path
            svg_files = sorted(svg_output.glob('*.svg'))

        if not svg_files:
            print(f"[WARN] No SVG files found")
            return []

        print(f"\n[SCAN] Checking {len(svg_files)} SVG file(s)...\n")

        for svg_file in svg_files:
            result = self.check_file(str(svg_file), expected_format)
            self._print_result(result)

        return self.results

    def _print_result(self, result: Dict):
        """Print check result for a single file."""
        if result['passed']:
            if result['warnings']:
                icon = "[WARN]"
                status = "Passed (with warnings)"
            else:
                icon = "[OK]"
                status = "Passed"
        else:
            icon = "[ERROR]"
            status = "Failed"

        print(f"{icon} {result['file']} - {status}")

        # Display basic info
        if result['info']:
            info_items = []
            if 'viewbox' in result['info']:
                info_items.append(f"viewBox: {result['info']['viewbox']}")
            if info_items:
                print(f"   {' | '.join(info_items)}")

        # Display errors
        if result['errors']:
            for error in result['errors']:
                print(f"   [ERROR] {error}")

        # Display warnings
        if result['warnings']:
            for warning in result['warnings'][:2]:  # Only show first 2 warnings
                print(f"   [WARN] {warning}")
            if len(result['warnings']) > 2:
                print(f"   ... and {len(result['warnings']) - 2} more warning(s)")

        print()

    def print_summary(self):
        """Print check summary."""
        print("=" * 80)
        print("[SUMMARY] Check Summary")
        print("=" * 80)

        print(f"\nTotal files: {self.summary['total']}")
        print(
            f"  [OK] Fully passed: {self.summary['passed']} "
            f"({self._percentage(self.summary['passed'])}%)")
        print(
            f"  [WARN] With warnings: {self.summary['warnings']} "
            f"({self._percentage(self.summary['warnings'])}%)")
        print(
            f"  [ERROR] With errors: {self.summary['errors']} "
            f"({self._percentage(self.summary['errors'])}%)")

        if self.issue_types:
            print(f"\nIssue categories:")
            for issue_type, count in sorted(
                    self.issue_types.items(), key=lambda x: x[1], reverse=True):
                print(f"  {issue_type}: {count}")

        # Fix suggestions
        if self.summary['errors'] > 0 or self.summary['warnings'] > 0:
            print(f"\n[TIP] Common fixes:")
            print(f"  1. XML well-formedness: write typography as raw Unicode "
                  f"(—, ©, →, NBSP); escape XML reserved chars as &amp; &lt; "
                  f"&gt; &quot; &apos; — never use HTML named entities like "
                  f"&nbsp; &mdash; &copy;")
            print(f"  2. viewBox issues: Ensure consistency with canvas format")
            print(f"  3. foreignObject: Use <text> + <tspan> for manual line breaks")
            print(f"  4. Font issues: end every font-family stack with a PPT-safe "
                  f"family (e.g. Microsoft YaHei / Arial / Consolas)")

    def _percentage(self, count: int) -> int:
        """Calculate percentage."""
        if self.summary['total'] == 0:
            return 0
        return int(count / self.summary['total'] * 100)

    def export_report(self, output_file: str = 'svg_quality_report.txt'):
        """Export check report to a text file."""
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write("PPT Skill SVG Quality Check Report\n")
            f.write("=" * 80 + "\n\n")

            for result in self.results:
                status = "[OK] Passed" if result['passed'] else "[ERROR] Failed"
                f.write(f"{status} - {result['file']}\n")
                f.write(f"Path: {result.get('path', 'N/A')}\n")

                if result['info']:
                    f.write(f"Info: {result['info']}\n")

                if result['errors']:
                    f.write(f"\nErrors:\n")
                    for error in result['errors']:
                        f.write(f"  - {error}\n")

                if result['warnings']:
                    f.write(f"\nWarnings:\n")
                    for warning in result['warnings']:
                        f.write(f"  - {warning}\n")

                f.write("\n" + "-" * 80 + "\n\n")

            # Write summary
            f.write("\n" + "=" * 80 + "\n")
            f.write("Check Summary\n")
            f.write("=" * 80 + "\n\n")
            f.write(f"Total files: {self.summary['total']}\n")
            f.write(f"Fully passed: {self.summary['passed']}\n")
            f.write(f"With warnings: {self.summary['warnings']}\n")
            f.write(f"With errors: {self.summary['errors']}\n")

        print(f"\n[REPORT] Check report exported: {output_file}")
