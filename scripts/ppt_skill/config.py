"""PPT Skill — configuration loader + canvas formats + design spec defaults.

Provides:
  - load_config() — reads config.txt from project root, returns a flat dict
  - get_llm_client() / get_vl_client() — OpenAI-compatible clients
  - CANVAS_FORMATS — canvas dimension dictionary (EMU)
  - DESIGN_COLORS, FONT_SIZES, LAYOUT_MARGINS — design spec defaults

Dimensions are in EMU (English Metric Units, 1px = 9525 EMU at 96 DPI).
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any


def _find_config() -> Path | None:
    """Search for config.txt in project root (parent of scripts/)."""
    # Check several likely locations
    candidates = [
        Path("config.txt"),
        Path(__file__).parent.parent.parent / "config.txt",  # scripts/../config.txt
    ]
    for c in candidates:
        if c.exists():
            return c.resolve()
    return None


def load_config() -> dict[str, str]:
    """Load key=value configuration from config.txt.

    Returns a flat dict with ALL config keys (VL_*, LLM_*, etc.).
    Missing config file returns empty dict — callers should provide defaults.
    """
    config_path = _find_config()
    if not config_path:
        return {}

    config: dict[str, str] = {}
    for line in config_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        config[key.strip()] = value.strip().strip('"').strip("'")

    return config


def get_llm_config(config: dict[str, str] | None = None) -> dict[str, str]:
    """Extract LLM configuration from config dict (or load on demand)."""
    if config is None:
        config = load_config()
    return {
        "provider": config.get("LLM_PROVIDER", "openai"),
        "model": config.get("LLM_MODEL", "gpt-4o"),
        "api_key": config.get("LLM_API_KEY", ""),
        "api_base": config.get("LLM_API_BASE", ""),
        "max_tokens": config.get("LLM_MAX_TOKENS", "8192"),
    }


def get_vl_config(config: dict[str, str] | None = None) -> dict[str, str]:
    """Extract VL configuration from config dict (or load on demand)."""
    if config is None:
        config = load_config()
    return {
        "provider": config.get("VL_PROVIDER", "openai"),
        "model": config.get("VL_MODEL", "gpt-4o"),
        "api_key": config.get("VL_API_KEY", ""),
        "api_base": config.get("VL_API_BASE", ""),
        "max_tokens": config.get("VL_MAX_TOKENS", "4096"),
        "enabled": config.get("VL_ENABLED", "true").lower() == "true",
    }


def _build_openai_client(
    api_key: str,
    api_base: str | None = None,
) -> Any:
    """Build an OpenAI-compatible client."""
    from openai import OpenAI

    kwargs: dict = {"api_key": api_key}
    if api_base:
        kwargs["base_url"] = api_base
    return OpenAI(**kwargs)


def get_llm_client(config: dict[str, str] | None = None) -> Any:
    """Get an OpenAI-compatible client configured for LLM (text) usage."""
    cfg = get_llm_config(config)
    api_key = cfg["api_key"] or os.environ.get("LLM_API_KEY", "")
    return _build_openai_client(api_key, cfg["api_base"] or None)


def get_vl_client(config: dict[str, str] | None = None) -> Any:
    """Get an OpenAI-compatible client configured for VL (vision) usage."""
    cfg = get_vl_config(config)
    api_key = cfg["api_key"] or os.environ.get("VL_API_KEY", "")
    return _build_openai_client(api_key, cfg["api_base"] or None)


CANVAS_FORMATS = {
    "ppt169":      {"width": 12192000, "height": 6858000,  "name": "PPT 16:9"},
    "ppt43":       {"width": 9144000,  "height": 6858000,  "name": "PPT 4:3"},
    "wechat":      {"width": 10800000, "height": 19200000, "name": "WeChat Article"},
    "xiaohongshu": {"width": 10800000, "height": 14400000, "name": "Rednote Post"},
    "moments":     {"width": 10800000, "height": 12600000, "name": "WeChat Moments"},
    "story":       {"width": 10800000, "height": 19200000, "name": "Instagram Story"},
    "banner":      {"width": 19200000, "height": 10800000, "name": "Banner"},
    "a4":          {"width": 11900000, "height": 16840000, "name": "A4 Document"},
}

# =========================================================================
# Phase 2–4: Design Spec Placeholders
#
# These constants provide fallback values when no spec is selected,
# preventing None crashes in Phase 4 generation. Actual values come
# from extracted spec YAML files in the specs/ directory.
#
# Phase 2 (Spec Extraction) populates these at runtime from a loaded spec.
# Phase 4 (PPT Generation) consumes them for slide styling.
# =========================================================================

# Default colors — overridden by extracted spec in Phase 2
DESIGN_COLORS: dict[str, str] = {
    "dk1":    "#000000",  # dark 1 — primary dark (background1)
    "dk2":    "#44546A",  # dark 2 — secondary dark (background2)
    "lt1":    "#FFFFFF",  # light 1 — primary light (text1)
    "lt2":    "#E7E6E6",  # light 2 — secondary light (text2)
    "accent1": "#4472C4",
    "accent2": "#ED7D31",
    "accent3": "#A5A5A5",
    "accent4": "#FFC000",
    "accent5": "#5B9BD5",
    "accent6": "#70AD47",
    "hlink":   "#0563C1",  # hyperlink
    "folHlink":"#954F72",  # followed hyperlink
}

# Default font sizes (in points) — overridden by extracted spec
FONT_SIZES: dict[str, float] = {
    "title":    44.0,
    "subtitle": 28.0,
    "h1":       32.0,
    "h2":       24.0,
    "body":     18.0,
    "small":    14.0,
    "caption":  11.0,
}

# Default margins (in inches) for standard 16:9 slides — overridden by extracted spec
LAYOUT_MARGINS: dict[str, float] = {
    "top":          0.5,
    "bottom":       0.5,
    "left":         1.0,
    "right":        1.0,
    "title_x":      1.0,
    "title_y":      0.5,
    "title_width":  8.0,
    "title_height": 1.2,
}
