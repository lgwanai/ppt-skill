"""Agent loop pipeline: generate → screenshot → VL compare → fix → repeat.

Full pipeline:
  1. Generate slide using provided generator function
  2. Save to temp PPTX, screenshot via QuickLook (qlmanage)
  3. VL model (Doubao) compares screenshot against spec expectations
  4. Extract fix instructions from VL feedback
  5. Apply fixes to generation parameters
  6. Repeat until score > threshold or max iterations

Usage:
    result = agent_loop_refine(generate_fn, spec_page, slide_index)
    if result.passed:
        # slide is ready
"""

from __future__ import annotations

import base64
import io
import os
import subprocess
import tempfile
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable

from pptx import Presentation

from ppt_skill.ppt_style_eval import evaluate_slide, StyleReport


MAX_ITERATIONS = 5
PASS_THRESHOLD = 0.85


@dataclass
class AgentLoopResult:
    slide_index: int = 0
    iterations: int = 0
    best_score: float = 0.0
    best_slide_data: dict | None = None
    reports: list[StyleReport] = field(default_factory=list)
    vl_feedback: list[str] = field(default_factory=list)
    passed: bool = False


# ── Screenshot via QuickLook ────────────────────────────────────────


def screenshot_pptx(pptx_path: Path, output_dir: Path) -> Path | None:
    """Render PPTX first slide to PNG using macOS QuickLook.

    Returns path to PNG, or None on failure.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    try:
        subprocess.run(
            ["qlmanage", "-t", "-s", "1920", "-o", str(output_dir), str(pptx_path)],
            capture_output=True, timeout=15,
        )
        # Find the generated thumbnail
        for f in output_dir.iterdir():
            if f.suffix == ".png" and pptx_path.stem in f.name:
                return f
        # Try any PNG
        pngs = sorted(output_dir.glob("*.png"))
        return pngs[0] if pngs else None
    except Exception:
        return None


# ── VL Comparison ───────────────────────────────────────────────────


def vl_compare_screenshot(
    screenshot_path: Path,
    spec_page: dict,
    previous_issues: list[str] | None = None,
) -> dict:
    """Send screenshot to VL model for comparison against spec.

    Returns {"score": 0-100, "issues": [...], "fixes": [...]}
    """
    try:
        from PIL import Image
    except ImportError:
        return {"score": 0, "issues": ["PIL not available"], "fixes": []}

    from ppt_skill.config import get_vl_client, get_vl_config

    vl_cfg = get_vl_config()
    api_key = vl_cfg["api_key"] or os.environ.get("VL_API_KEY", "")
    if not api_key:
        return {"score": 0, "issues": ["No VL API key configured"], "fixes": []}

    client = get_vl_client()
    model = vl_cfg["model"]

    # Load and compress image
    try:
        img = Image.open(screenshot_path).convert('RGB')
        if img.width > 1024:
            ratio = 1024 / img.width
            img = img.resize((1024, int(img.height * ratio)))
        buf = io.BytesIO()
        img.save(buf, 'JPEG', quality=65)
        img_b64 = base64.b64encode(buf.getvalue()).decode()
    except Exception as e:
        return {"score": 0, "issues": [f"Screenshot error: {e}"], "fixes": []}

    # Build spec description
    spec_desc = "Expected style:\n"
    spec_desc += f"  Background: {spec_page.get('background', {}).get('description', 'white')}\n"
    for e in spec_page.get("elements", []):
        ts = e.get("text_style", {}) or {}
        pos = e.get("position", {})
        spec_desc += (
            f"  {e['role']}: pos=({pos.get('x',0):.2f},{pos.get('y',0):.2f}) "
            f"font={ts.get('font_family','?')} {ts.get('font_size_pt',0):.0f}pt "
            f"bold={ts.get('font_weight','?')} color={ts.get('font_color','?')} "
            f"align={ts.get('text_alignment','?')}\n"
        )

    prev = ""
    if previous_issues:
        prev = "\nPrevious issues to fix:\n" + "\n".join(f"  - {i}" for i in previous_issues)

    prompt = f"""Compare this slide screenshot against the specification. {prev}
{spec_desc}

Rate style fidelity 0-100. List specific issues with exact values (e.g. "title should be 40pt not 36pt").
Return JSON: {{"score": 85, "issues": ["title size wrong: 36pt should be 40pt"], "fixes": ["change title font size to 40pt"]}}
"""
    try:
        response = client.chat.completions.create(
            model=model,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{img_b64}"}},
                ],
            }],
            max_tokens=1024, temperature=0.1,
        )
        import json, re
        text = response.choices[0].message.content
        match = re.search(r'\{[^}]+\}', text, re.DOTALL)
        if match:
            return json.loads(match.group(0))
        return {"score": 0, "issues": [text], "fixes": []}
    except Exception as e:
        return {"score": 0, "issues": [f"VL error: {e}"], "fixes": []}


# ── Style fix application ────────────────────────────────────────────


def apply_fixes_from_vl(fixes: list[str], style_params: dict) -> dict:
    """Parse VL fix instructions and update style parameters.

    Example fixes:
      "change title font size to 40pt" → style_params["title"]["size"] = 40
      "title should be bold" → style_params["title"]["bold"] = True
      "subtitle color should be #0070C0" → style_params["subtitle"]["color"] = "#0070C0"
    """
    import re
    for fix in fixes:
        fix_lower = fix.lower()
        # Which element?
        element = "body"
        for role in ["title", "subtitle", "body", "header"]:
            if role in fix_lower:
                element = role
                break

        if element not in style_params:
            style_params[element] = {}

        # Font size
        m = re.search(r'(\d+)\s*pt', fix)
        if m and "size" in fix_lower:
            style_params[element]["size"] = int(m.group(1))

        # Bold
        if "bold" in fix_lower:
            if "not bold" in fix_lower or "unbold" in fix_lower:
                style_params[element]["bold"] = False
            else:
                style_params[element]["bold"] = True

        # Color
        m = re.search(r'#[0-9A-Fa-f]{6}', fix)
        if m and "color" in fix_lower:
            style_params[element]["color"] = m.group(0)

    return style_params


# ── Agent Loop ────────────────────────────────────────────────────────


def agent_loop_refine(
    generate_fn: Callable[[dict, int], Path],
    spec_page: dict,
    slide_index: int,
    max_iterations: int = MAX_ITERATIONS,
    use_vl: bool = True,
) -> AgentLoopResult:
    """Run agent loop: generate → evaluate → fix → repeat.

    Args:
        generate_fn: (style_params, slide_index) → Path to generated PPTX file.
                     Called with updated params each iteration.
        spec_page: Spec page dict with expected styles.
        slide_index: Slide number.
        max_iterations: Max refinement iterations.
        use_vl: Whether to use VL model for comparison.

    Returns AgentLoopResult with best slide and scores.
    """
    result = AgentLoopResult(slide_index=slide_index)
    style_params: dict = {}
    previous_issues: list[str] | None = None
    best_score = 0.0
    best_pptx: Path | None = None

    for iteration in range(1, max_iterations + 1):
        # 1. Generate
        try:
            pptx_path = generate_fn(style_params, slide_index)
        except Exception as e:
            result.reports.append(StyleReport(issues=[f"Generation error: {e}"]))
            continue

        if not pptx_path or not pptx_path.exists():
            result.reports.append(StyleReport(issues=["PPTX not created"]))
            continue

        # 2. Screenshot
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            screenshot = screenshot_pptx(pptx_path, tmp)

            # 3. Evaluate — programmatic check first
            try:
                prs = Presentation(str(pptx_path))
                report = evaluate_slide(
                    list(prs.slides)[0], spec_page,
                    prs.slide_width, prs.slide_height,
                )
            except Exception as e:
                report = StyleReport(issues=[f"Evaluation error: {e}"])

            result.reports.append(report)

            # Track best
            if report.overall > best_score:
                best_score = report.overall
                result.best_score = best_score

            if report.passed and report.overall >= PASS_THRESHOLD:
                result.iterations = iteration
                result.passed = True
                return result

            # 4. VL comparison (optional)
            if use_vl and screenshot:
                try:
                    vl_result = vl_compare_screenshot(screenshot, spec_page, previous_issues)
                    vl_score = vl_result.get("score", 0) / 100.0
                    if vl_score > best_score:
                        best_score = vl_score
                        result.best_score = best_score
                    if vl_score >= PASS_THRESHOLD:
                        result.iterations = iteration
                        result.passed = True
                        return result

                    # Apply VL fixes for next iteration
                    fixes = vl_result.get("fixes", [])
                    if fixes:
                        result.vl_feedback.extend(fixes)
                        style_params = apply_fixes_from_vl(fixes, style_params)
                except Exception:
                    pass

            # 5. Prepare fix instructions from programmatic report
            previous_issues = report.issues

    result.iterations = max_iterations
    result.best_score = best_score
    result.passed = best_score >= PASS_THRESHOLD
    return result
