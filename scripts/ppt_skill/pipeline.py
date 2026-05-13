#!/usr/bin/env python3
"""PPT Generation Pipeline — parse outline, match templates, generate PPTX.

Steps:
  1. parse outline → page list + content JSON
  2. match pages to template pages (paginate.py)
  3. merge content into templates + generate PPTX (merge_and_generate.py)

Pre-requisites (one-time setup per template):
  - extract_ppt_elements.py → individual page JSONs
  - ppt_to_png.py → page screenshots
  - vl_analyze.py → page analysis JSON
"""

import sys
from pathlib import Path

from ppt_skill.outline_parser import outlines_to_pages_file
from ppt_skill.paginate import main as paginate_main
from ppt_skill.merge_and_generate import main as merge_main


def full_pipeline(outline_path: str, output_pptx: str,
                  spec_dir: str = 'specs/business-trip',
                  template_pptx: str = None,
                  page_analysis: str = None):
    """Run the complete PPT generation pipeline.

    Args:
        outline_path: Path to markdown outline (e.g. outlines/agent_course.md)
        output_pptx: Output PPTX file path
        spec_dir: Directory with template files (page_analysis.json, ppt_demo1_page*.json)
        template_pptx: Source PPTX for extraction (first run only)
        page_analysis: Path to vl_analyze output JSON
    """
    import os
    cwd = os.getcwd()

    # Step 0: One-time template extraction (if source PPTX provided)
    if template_pptx and Path(template_pptx).exists():
        print("=== STEP 0: Extract template elements ===")
        os.system(f"cd {spec_dir} && python {cwd}/scripts/ppt_skill/extract_spec_v2.py ../{template_pptx}")
        # Convert to PNG
        png_dir = Path(template_pptx).stem + '_png'
        os.system(f"cd {spec_dir} && python {cwd}/scripts/ppt_skill/ppt_to_png.py {template_pptx} {png_dir}")
        # VL analysis
        if page_analysis:
            os.system(f"cd {spec_dir} && python {cwd}/scripts/ppt_skill/vl_analyze.py {png_dir} {page_analysis}")

    # Step 1: Parse outline → content pages JSON
    print("\n=== STEP 1: Parse outline ===")
    content_json = f"{Path(outline_path).stem}_content.json"
    pages = outlines_to_pages_file(outline_path, content_json)
    print(f"  Parsed {len(pages)} pages → {content_json}")

    # Step 2: Match content pages to templates
    print("\n=== STEP 2: Match templates ===")
    if not page_analysis:
        # Try to find analysis file
        for candidate in [Path(spec_dir) / 'page_analysis.json',
                          Path(spec_dir) / 'demo1_analysis.json']:
            if candidate.exists():
                page_analysis = str(candidate)
                break
    if not page_analysis:
        print("  WARNING: No page_analysis.json found, using deterministic matching")
        # Fall back to deterministic: just number pages sequentially
        # Create a minimal analysis from the content directly
        pages_output = f"{Path(outline_path).stem}_pages.json"
        # Simple mapping: just assign template pages round-robin
        _simple_match(pages, pages_output)
    else:
        pages_output = f"{Path(outline_path).stem}_pages.json"
        # Run paginate.py from the spec dir (it reads markdown directly)
        os.system(f"cd {spec_dir} && python {cwd}/scripts/ppt_skill/paginate.py {cwd}/{outline_path} {cwd}/{page_analysis}")
        # paginate.py outputs to stdout, need to save it
        print(f"  (paginate.py outputs to stdout, redirect or use merge directly)")

    # Step 3: Merge and generate PPTX
    print(f"\n=== STEP 3: Generate PPTX ===")
    merge_main(pages_output, output_pptx, 0, None)

    print(f"\n=== Done: {output_pptx} ===")


def _simple_match(pages: list, output_path: str):
    """Fallback: simple sequential matching when no VL analysis available."""
    import json, glob
    # Find available template pages
    tmpl = {}
    for f in sorted(glob.glob('ppt_demo1_page*.json')):
        import re
        m = re.search(r'page(\d+)', f)
        if m: tmpl[int(m.group(1))] = f
    
    output = []
    for p in pages:
        matched = (
            'ppt_demo1.json' if p['page_type'] == 'cover'
            else 'ppt_demo1_page2.json'  # default to content page template
        )
        output.append({
            **p,
            'template_page': matched,
            'match_reason': 'deterministic fallback',
        })
    
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"  Created {output_path} ({len(output)} pages)")
