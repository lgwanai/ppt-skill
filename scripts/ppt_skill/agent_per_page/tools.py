"""Agent-per-page PPT generation tools — delegates to outline_parser and gen_pptx."""
# ── Tool 1: get_page ──────────────────────────────────────────────
from ppt_skill.outline_parser import parse_outline

_PAGE_CACHE = {}


def get_page(outline_path: str, completed: set) -> dict | None:
    """Fetch next uncompleted page from outline. Results cached."""
    if outline_path not in _PAGE_CACHE:
        raw = open(outline_path, encoding='utf-8').read()
        _PAGE_CACHE[outline_path] = parse_outline(raw)
    for p in _PAGE_CACHE[outline_path]:
        if p['index'] not in completed:
            return p
    return None


# ── Re-export for backward compat ──────────────────────────────────
# These are now handled by gen_pptx.py, extract_spec_v2.py, vl_analyze.py, paginate.py
