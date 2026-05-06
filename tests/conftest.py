"""Pytest fixtures for ppt-skill tests."""

from pathlib import Path

import pytest


@pytest.fixture
def fixtures_dir():
    return Path(__file__).parent / "fixtures"


@pytest.fixture
def banned_svg_dir(fixtures_dir):
    return fixtures_dir / "banned_features_svg"


@pytest.fixture
def temp_output(tmp_path):
    """Temporary .pptx output path that auto-cleans."""
    return tmp_path / "test_output.pptx"
