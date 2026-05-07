"""PPT Skill — minimal configuration for Phase 1 pipeline.

Contains only CANVAS_FORMATS — the canvas dimension dictionary needed by
the quality checker and converter pipeline. All ppt-master-specific
configuration (DESIGN_COLORS, INDUSTRY_COLORS, LAYOUT_MARGINS, FONT_SIZES)
belong to Phase 2 (Spec Extraction) and are intentionally excluded.

Dimensions are in EMU (English Metric Units, 1px = 9525 EMU at 96 DPI).
"""

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
