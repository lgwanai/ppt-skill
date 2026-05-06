"""Tests for SVG quality checker — banned feature detection."""

from pathlib import Path

from ppt_skill.quality import SVGQualityChecker

FIXTURES = Path(__file__).parent / "fixtures"
BANNED = FIXTURES / "banned_features_svg"


class TestBannedFeatures:
    """Each banned feature must be detected with a specific error."""

    def setup_method(self):
        self.checker = SVGQualityChecker()

    def test_rejects_mask(self):
        result = self.checker.check_file(str(BANNED / "mask.svg"))
        assert not result["passed"]
        assert any("mask" in e.lower() for e in result["errors"])

    def test_rejects_rgba(self):
        result = self.checker.check_file(str(BANNED / "rgba.svg"))
        assert not result["passed"]
        assert any("rgba" in e.lower() for e in result["errors"])

    def test_rejects_fontface(self):
        result = self.checker.check_file(str(BANNED / "fontface.svg"))
        assert not result["passed"]
        assert any(
            "font-face" in e.lower() or "@font-face" in e.lower()
            for e in result["errors"]
        )

    def test_rejects_html_entities(self):
        result = self.checker.check_file(str(BANNED / "html_entities.svg"))
        assert not result["passed"]
        # HTML entities trigger XML well-formedness errors
        assert len(result["errors"]) > 0

    def test_rejects_style_tag(self):
        result = self.checker.check_file(str(BANNED / "style_tag.svg"))
        assert not result["passed"]
        assert any("style" in e.lower() for e in result["errors"])


class TestCleanSVGs:
    """Clean SVGs must pass quality check without errors."""

    def setup_method(self):
        self.checker = SVGQualityChecker()

    def test_accepts_simple_rect(self):
        result = self.checker.check_file(str(FIXTURES / "sample_simple.svg"))
        assert result["passed"], f"Errors: {result.get('errors', [])}"

    def test_accepts_text(self):
        result = self.checker.check_file(str(FIXTURES / "sample_text.svg"))
        assert result["passed"], f"Errors: {result.get('errors', [])}"

    def test_accepts_gradient(self):
        result = self.checker.check_file(str(FIXTURES / "sample_gradient.svg"))
        assert result["passed"], f"Errors: {result.get('errors', [])}"
