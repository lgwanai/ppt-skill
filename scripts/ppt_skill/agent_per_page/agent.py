"""PageAgent — DEPRECATED. Use ppt_skill.paginate + ppt_skill.merge_and_generate instead."""
# The agent-per-page approach has been replaced by:
#   1. extract_spec_v2.py — extract template elements
#   2. vl_analyze.py — VL model analysis
#   3. paginate.py (match_template) — content → template matching
#   4. gen_pptx.py (generate_ppt) — JSON → PPTX with proper shapes
#   5. merge_and_generate.py — multi-page pipeline
