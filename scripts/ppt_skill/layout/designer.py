"""Layout design agent loop using prompt-ppt-layout.md principles.

Applies professional layout design to slides:
- WPS hierarchy (What/Point/Support)
- Typography rules (sans-serif, 1.5x spacing)
- Color rules (single primary, dark text)
- Whitespace (margins, breathing room)

Usage:
    designer = LayoutDesigner()
    styled_slide = designer.design(slide_content, spec_page)
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


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


class LayoutDesigner:
    """Apply professional layout design using agent loop."""

    def __init__(self):
        self.layout_prompt = self._load_layout_prompt()
        self.max_iterations = 3

    def _load_layout_prompt(self) -> str:
        """Load prompt-ppt-layout.md content."""
        prompt_path = Path("references/prompt-ppt-layout.md")
        if prompt_path.exists():
            return prompt_path.read_text(encoding="utf-8")
        return ""

    def design(
        self,
        slide_content: dict,
        spec_page: dict,
        generate_callback: Any = None,
    ) -> LayoutResult:
        """Design layout for a slide using agent loop.

        Args:
            slide_content: Slide data with w, p, s fields
            spec_page: Design spec page
            generate_callback: Function to generate layout (LLM-backed)

        Returns:
            LayoutResult with layout params and validation
        """
        result = LayoutResult()
        params = self._initial_params(slide_content, spec_page)

        for iteration in range(1, self.max_iterations + 1):
            result.iterations = iteration
            result.params = params

            # Validate layout
            issues = self._validate_layout(params, slide_content)
            result.issues = issues

            # Check for errors
            errors = [i for i in issues if i.severity == "error"]
            if not errors:
                result.passed = True
                return result

            # Apply fixes
            fixes = self._generate_fixes(issues)
            result.fixes_applied.extend(fixes)
            params = self._apply_fixes(params, fixes)

        return result

    def _initial_params(self, slide_content: dict, spec_page: dict) -> LayoutParams:
        """Generate initial layout params from content and spec."""
        # Extract from spec
        colors = spec_page.get("colors", {})
        typography = spec_page.get("typography", {})

        primary_color = colors.get("primary", "#0070C0")
        if isinstance(primary_color, list):
            primary_color = primary_color[0] if primary_color else "#0070C0"

        font_family = typography.get("heading_family", "Arial")
        if not font_family or font_family in ("宋体", "仿宋", "楷体"):
            font_family = "Arial"  # Force sans-serif

        return LayoutParams(
            font_family=font_family,
            primary_color=primary_color,
            text_color=colors.get("body", "#333333"),
            background_color=colors.get("background", "#FFFFFF"),
        )

    def _validate_layout(self, params: LayoutParams, content: dict) -> list[LayoutIssue]:
        """Validate layout against design rules."""
        issues: list[LayoutIssue] = []

        # Rule 1: Sans-serif only
        if params.font_family in ("宋体", "仿宋", "楷体", "SimSun", "FangSong", "KaiTi"):
            issues.append(LayoutIssue(
                rule="sans_serif",
                message=f"Font '{params.font_family}' is serif. Use sans-serif.",
                severity="error",
            ))

        # Rule 2: P larger than S
        if params.p_font_size <= params.s_font_size:
            issues.append(LayoutIssue(
                rule="hierarchy",
                message=f"P font size ({params.p_font_size}) must be larger than S ({params.s_font_size})",
                severity="error",
            ))

        # Rule 3: P should be bold
        if params.p_font_weight != "bold":
            issues.append(LayoutIssue(
                rule="bold_point",
                message="Point (P) should be bold for emphasis",
                severity="warning",
            ))

        # Rule 4: Line spacing should be 1.5
        if params.line_spacing < 1.3 or params.line_spacing > 1.7:
            issues.append(LayoutIssue(
                rule="line_spacing",
                message=f"Line spacing should be 1.5x, got {params.line_spacing}",
                severity="warning",
            ))

        # Rule 5: Margins should be adequate
        if params.left_margin < 0.08 or params.right_margin < 0.08:
            issues.append(LayoutIssue(
                rule="margins",
                message="Page margins should be at least 8%",
                severity="warning",
            ))

        # Rule 6: Text should be dark
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

        # Keywords to icons mapping
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
