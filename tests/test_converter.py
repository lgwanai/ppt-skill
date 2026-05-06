"""Tests for SVG→DrawingML converter pipeline."""

from ppt_skill.converter.context import ConvertContext
from ppt_skill.converter.converter import convert_svg_to_slide_shapes
from ppt_skill.converter.paths import parse_svg_path, path_commands_to_drawingml
from ppt_skill.converter.utils import parse_hex_color


class TestConverterImports:
    """Verify all key converter symbols are importable."""

    def test_import_converter_dispatcher(self):
        assert callable(convert_svg_to_slide_shapes)

    def test_import_convert_context(self):
        assert ConvertContext is not None

    def test_import_color_parser(self):
        assert callable(parse_hex_color)

    def test_import_path_parser(self):
        assert callable(parse_svg_path)

    def test_import_path_commands_converter(self):
        assert callable(path_commands_to_drawingml)


class TestFinalizeImports:
    """Verify all finalize post-processing symbols are importable."""

    def test_import_icon_resolver(self):
        from ppt_skill.finalize.embed_icons import resolve_icon_path

        assert callable(resolve_icon_path)

    def test_import_tspan_flattener(self):
        from ppt_skill.finalize.flatten_tspan import flatten_text_with_tspans

        assert callable(flatten_text_with_tspans)


class TestUtils:
    """Unit tests for utility functions (no SVG files needed)."""

    def test_parse_hex_color_6digit(self):
        result = parse_hex_color("#4472C4")
        assert result is not None

    def test_parse_hex_color_3digit(self):
        result = parse_hex_color("#FFF")
        assert result is not None

    def test_parse_hex_color_invalid(self):
        result = parse_hex_color("rgb(1,2,3)")
        assert result is None
