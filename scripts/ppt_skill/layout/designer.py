"""Layout design agent loop using prompt-ppt-layout.md principles.

Applies professional layout design to slides:
- WPS hierarchy (What/Point/Support)
- Typography rules (sans-serif, 1.5x spacing)
- Color rules (single primary, dark text)
- Whitespace (margins, breathing room)

Usage:
    designer = LayoutDesigner()
    result = designer.design(slide_content, spec_page, llm_callback)
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable


@dataclass
class LayoutParams:
    """Layout parameters for a slide."""

    # Typography
    font_family: str = "Arial"  # Sans-serif only
    w_font_size: float = 10.0  # Navigation (small)
    p_font_size: float = 24.0  # Point (large)
    s_font_size: float = 12.0  # Support (medium)
    p_font_weight: str = "bold"
    s_font_weight: str = "regular"
    line_spacing: float = 1.5

    # Colors
    primary_color: str = "#0070C0"  # Single primary color
    text_color: str = "#333333"  # Dark gray for body
    background_color: str = "#FFFFFF"  # White or light

    # Whitespace (as percentage of page)
    left_margin: float = 0.10  # 10%
    right_margin: float = 0.10
    top_margin: float = 0.08
    bottom_margin: float = 0.08

    # Layout
    w_position: str = "top"  # Navigation position
    p_position: str = "upper"  # Point position
    s_position: str = "lower"  # Support position

    # Icons
    icon_name: str = ""  # Bootstrap icon name
    icon_position: str = "right"  # Icon placement

    def to_dict(self) -> dict:
        return {
            "font_family": self.font_family,
            "w_font_size": self.w_font_size,
            "p_font_size": self.p_font_size,
            "s_font_size": self.s_font_size,
            "p_font_weight": self.p_font_weight,
            "s_font_weight": self.s_font_weight,
            "line_spacing": self.line_spacing,
            "primary_color": self.primary_color,
            "text_color": self.text_color,
            "background_color": self.background_color,
            "left_margin": self.left_margin,
            "right_margin": self.right_margin,
            "top_margin": self.top_margin,
            "bottom_margin": self.bottom_margin,
            "w_position": self.w_position,
            "p_position": self.p_position,
            "s_position": self.s_position,
            "icon_name": self.icon_name,
            "icon_position": self.icon_position,
        }

    @classmethod
    def from_dict(cls, data: dict) -> LayoutParams:
        return cls(**{k: v for k, v in data.items() if hasattr(cls, k)})


@dataclass
class LayoutIssue:
    """Layout validation issue."""

    rule: str
    message: str
    severity: str  # error, warning


@dataclass
class LayoutResult:
    """Result of layout design."""

    passed: bool = False
    iterations: int = 0
    params: LayoutParams | None = None
    issues: list[LayoutIssue] = field(default_factory=list)
    fixes_applied: list[str] = field(default_factory=list)
    llm_feedback: list[str] = field(default_factory=list)


class LayoutDesigner:
    """Apply professional layout design using LLM-driven agent loop."""

    def __init__(self):
        self.layout_prompt = self._load_layout_prompt()
        self.max_iterations = 3

    def _load_layout_prompt(self) -> str:
        """Load prompt-ppt-layout.md content."""
        prompt_path = Path("references/prompt-ppt-layout.md")
        if prompt_path.exists():
            return prompt_path.read_text(encoding="utf-8")
        return ""

    def _build_layout_prompt(
        self,
        slide_content: dict,
        spec_page: dict,
        previous_fixes: list[str] | None = None,
    ) -> str:
        """Build the LLM prompt for layout design using prompt-ppt-layout.md."""
        w = slide_content.get("w", "")
        p = slide_content.get("p", "")
        s = slide_content.get("s", [])

        page_type = spec_page.get("page_type", "content")
        colors = spec_page.get("colors", {})
        typography = spec_page.get("typography", {})

        fix_section = ""
        if previous_fixes:
            fix_section = (
                "\n## Previous Issues to Fix\n"
                + "\n".join(f"  - {fix}" for fix in previous_fixes)
                + "\n\nFix ALL of the above issues in your response.\n"
            )

        bg_color = colors.get("lt1") or colors.get("background1") or colors.get("background", "#FFFFFF")
        primary_color_display = colors.get("accent1") or colors.get("primary", "#0070C0")
        text_color_display = colors.get("dk1") or colors.get("text1") or colors.get("body", "#333333")
        heading_font = typography.get("heading_family", "Arial")
        body_font = typography.get("body_family", "Arial")

        return f"""You are a professional slide layout designer. Apply the following layout principles to redesign this slide.

## Layout Design Principles
{self.layout_prompt}

## Slide Content
- **W (What/Navigation)**: {w}
- **P (Point/核心观点)**: {p}
- **S (Support/支撑论据)**: {chr(10).join(f'  • {b}' for b in s) if s else 'N/A'}
- **Page Type**: {page_type}

## Design Spec
- Background color: {bg_color}
- Primary color: {primary_color_display}
- Text color: {text_color_display}
- Heading font: {heading_font}
- Body font: {body_font}

## Requirements
1. Font family MUST be sans-serif (e.g. Arial, Helvetica)
2. W (navigation) should be small text (8-12pt) at top
3. P (point) should be large (22-28pt), bold, prominent
4. S (support) should be medium (11-16pt), regular weight, below P
5. Line spacing 1.5x
6. Single primary color for accents; dark text (#333-#444) for body
7. Margins: 8-12% on all sides
8. Suggest a Bootstrap icon name (from bi-* set) that matches the content{fix_section}

## Output Format
Return ONLY valid JSON with these keys:
{{
  "font_family": "Arial",
  "w_font_size": 10.0,
  "p_font_size": 24.0,
  "s_font_size": 12.0,
  "p_font_weight": "bold",
  "s_font_weight": "regular",
  "line_spacing": 1.5,
  "primary_color": "#0070C0",
  "text_color": "#333333",
  "background_color": "#FFFFFF",
  "left_margin": 0.10,
  "right_margin": 0.10,
  "top_margin": 0.08,
  "bottom_margin": 0.08,
  "w_position": "top",
  "p_position": "upper",
  "s_position": "lower",
  "icon_name": "bi-graph-up",
  "icon_position": "right"
}}
"""

    def design(
        self,
        slide_content: dict,
        spec_page: dict,
        generate_callback: Callable[[str], str] | None = None,
    ) -> LayoutResult:
        """Design layout for a slide using LLM-driven agent loop.

        Uses the LLM (via generate_callback) with prompt-ppt-layout.md
        to generate optimal layout parameters, then validates and iterates.

        Args:
            slide_content: Slide data with w, p, s fields
            spec_page: Design spec page
            generate_callback: LLM function (prompt_str) -> response text.
                               If None, falls back to rule-based params.

        Returns:
            LayoutResult with layout params and validation
        """
        result = LayoutResult()

        if generate_callback is None:
            return self._fallback_rule_based(slide_content, spec_page, result)

        previous_fixes: list[str] | None = None

        for iteration in range(1, self.max_iterations + 1):
            result.iterations = iteration

            prompt = self._build_layout_prompt(slide_content, spec_page, previous_fixes)

            try:
                response = generate_callback(prompt)
                params_data = self._parse_llm_response(response)
                params = LayoutParams.from_dict(params_data) if params_data else self._initial_params(slide_content, spec_page)
            except Exception:
                params = self._initial_params(slide_content, spec_page)

            result.params = params

            issues = self._validate_layout(params, slide_content)
            result.issues = issues

            errors = [i for i in issues if i.severity == "error"]
            warnings = [i for i in issues if i.severity == "warning"]

            if not errors:
                result.passed = True
                if warnings:
                    result.fixes_applied.extend(
                        self._generate_fixes(warnings)
                    )
                return result

            fixes = self._generate_fixes(errors)
            result.fixes_applied.extend(fixes)
            result.llm_feedback.extend(fixes)
            previous_fixes = fixes

        return result

    def _parse_llm_response(self, response: str) -> dict | None:
        """Parse JSON from LLM response, handling markdown fences."""
        if not response:
            return None

        json_match = re.search(r'\{[^{}]*\}', response, re.DOTALL)
        if json_match:
            try:
                return json.loads(json_match.group(0))
            except json.JSONDecodeError:
                return None
        return None

    def _fallback_rule_based(
        self,
        slide_content: dict,
        spec_page: dict,
        result: LayoutResult,
    ) -> LayoutResult:
        """Fallback: generate layout params from spec data (no LLM)."""
        params = self._initial_params(slide_content, spec_page)
        result.params = params

        for iteration in range(1, self.max_iterations + 1):
            result.iterations = iteration

            issues = self._validate_layout(params, slide_content)
            result.issues = issues

            errors = [i for i in issues if i.severity == "error"]
            if not errors:
                result.passed = True
                return result

            fixes = self._generate_fixes(issues)
            result.fixes_applied.extend(fixes)
            params = self._apply_fixes(params, fixes)
            result.params = params

        return result

    def _initial_params(self, slide_content: dict, spec_page: dict) -> LayoutParams:
        """Generate initial layout params from content and spec."""
        colors = spec_page.get("colors", {})
        typography = spec_page.get("typography", {})

        # Handle both VL JSON palette (accent1, dk1, lt1) and legacy YAML formats
        primary_color = (
            colors.get("accent1") or colors.get("primary") or "#0070C0"
        )
        if isinstance(primary_color, list):
            primary_color = primary_color[0] if primary_color else "#0070C0"

        text_color = (
            colors.get("dk1") or colors.get("text1") or colors.get("body") or "#000000"
        )
        background_color = (
            colors.get("lt1") or colors.get("background1") or colors.get("background") or "#FFFFFF"
        )

        font_family = typography.get("heading_family", "Arial")
        if not font_family or font_family in ("宋体", "仿宋", "楷体"):
            font_family = "Arial"

        return LayoutParams(
            font_family=font_family,
            primary_color=primary_color,
            text_color=text_color,
            background_color=background_color,
        )

    def _validate_layout(self, params: LayoutParams, content: dict) -> list[LayoutIssue]:
        """Validate layout against design rules."""
        issues: list[LayoutIssue] = []

        if params.font_family in ("宋体", "仿宋", "楷体", "SimSun", "FangSong", "KaiTi"):
            issues.append(LayoutIssue(
                rule="sans_serif",
                message=f"Font '{params.font_family}' is serif. Use sans-serif.",
                severity="error",
            ))

        if params.p_font_size <= params.s_font_size:
            issues.append(LayoutIssue(
                rule="hierarchy",
                message=f"P font size ({params.p_font_size}) must be larger than S ({params.s_font_size})",
                severity="error",
            ))

        if params.p_font_weight != "bold":
            issues.append(LayoutIssue(
                rule="bold_point",
                message="Point (P) should be bold for emphasis",
                severity="warning",
            ))

        if params.line_spacing < 1.3 or params.line_spacing > 1.7:
            issues.append(LayoutIssue(
                rule="line_spacing",
                message=f"Line spacing should be 1.5x, got {params.line_spacing}",
                severity="warning",
            ))

        if params.left_margin < 0.08 or params.right_margin < 0.08:
            issues.append(LayoutIssue(
                rule="margins",
                message="Page margins should be at least 8%",
                severity="warning",
            ))

        if params.text_color.upper() not in ("#333333", "#333", "#444444", "#444", "#000000", "#000"):
            issues.append(LayoutIssue(
                rule="text_color",
                message="Body text should be dark gray or black",
                severity="warning",
            ))

        return issues

    def _generate_fixes(self, issues: list[LayoutIssue]) -> list[str]:
        """Generate fix instructions from issues."""
        fixes: list[str] = []

        for issue in issues:
            if issue.rule == "sans_serif":
                fixes.append("Change font_family to Arial or Helvetica")
            elif issue.rule == "hierarchy":
                fixes.append("Set p_font_size to 24pt and s_font_size to 12pt")
            elif issue.rule == "bold_point":
                fixes.append("Set p_font_weight to bold")
            elif issue.rule == "line_spacing":
                fixes.append("Set line_spacing to 1.5")
            elif issue.rule == "margins":
                fixes.append("Set margins to 10% left/right and 8% top/bottom")
            elif issue.rule == "text_color":
                fixes.append("Set text_color to #333333")

        return fixes

    def _apply_fixes(self, params: LayoutParams, fixes: list[str]) -> LayoutParams:
        """Apply fix instructions to layout params."""
        for fix in fixes:
            fix_lower = fix.lower()

            if "font_family" in fix_lower:
                params.font_family = "Arial"
            elif "p_font_size" in fix_lower:
                params.p_font_size = 24.0
            elif "s_font_size" in fix_lower:
                params.s_font_size = 12.0
            elif "p_font_weight" in fix_lower:
                params.p_font_weight = "bold"
            elif "line_spacing" in fix_lower:
                params.line_spacing = 1.5
            elif "margins" in fix_lower:
                params.left_margin = 0.10
                params.right_margin = 0.10
                params.top_margin = 0.08
                params.bottom_margin = 0.08
            elif "text_color" in fix_lower:
                params.text_color = "#333333"

        return params

    def get_icon_suggestion(self, content: dict) -> str:
        """Suggest Bootstrap icon based on content."""
        text = " ".join([
            content.get("w", ""),
            content.get("p", ""),
            " ".join(content.get("s", [])),
        ]).lower()

        icon_keywords = {
            "bi-graph-up": ["data", "chart", "growth", "increase", "revenue", "sales"],
            "bi-people": ["team", "people", "staff", "employee", "customer"],
            "bi-lightbulb": ["idea", "innovation", "new", "creative", "solution"],
            "bi-check-circle": ["success", "achieve", "complete", "done", "goal"],
            "bi-arrow-right": ["next", "process", "step", "flow", "progress"],
            "bi-gear": ["setting", "config", "system", "technical", "tool"],
            "bi-cpu": ["ai", "technology", "machine", "compute", "algorithm"],
            "bi-currency-dollar": ["money", "price", "cost", "budget", "finance"],
            "bi-trophy": ["award", "win", "champion", "best", "top"],
            "bi-globe": ["global", "world", "international", "market"],
        }

        for icon, keywords in icon_keywords.items():
            if any(kw in text for kw in keywords):
                return icon

        return ""


__all__ = ["LayoutDesigner", "LayoutParams", "LayoutResult", "LayoutIssue"]
