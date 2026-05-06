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
