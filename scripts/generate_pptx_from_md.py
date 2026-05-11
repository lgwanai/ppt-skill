#!/usr/bin/env python3
"""
Generate PPTX from a markdown outline that follows prompt-ppt-content.md format.

Rules:
  ## 第一章：[名称]   → section divider (NOT a slide, used for context)
  ### 转场页：[目标]   → transition slide
  ### [页面标题]       → content slide
  ## 封面             → cover slide
  ## 结尾页            → end slide

Usage:
  python3 scripts/generate_pptx_from_md.py \\
    --spec specs/agent_course/ \\
    --outline outlines/agent_course.md \\
    -o output.pptx
"""

import argparse, re, os, sys, yaml
from pathlib import Path
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

SW, SH = 12192000, 6858000

C_B = RGBColor(0x00,0x70,0xC0)
C_B2 = RGBColor(0x44,0x72,0xC4)
C_D = RGBColor(0x44,0x54,0x6A)
C_W = RGBColor(0xFF,0xFF,0xFF)
C_O = RGBColor(0xED,0x7D,0x31)

def nx(x): return int(x*SW)
def ny(y): return int(y*SH)
def nw(w): return int(w*SW)
def nh(h): return int(h*SH)

TL = Emu(457200)
TW = Emu(SW) - TL - Emu(914400)
BL = Emu(365760)
BW = Emu(SW) - BL * 2
CW = Emu(SW * 8 // 10)
CX = Emu(SW // 10)
BODY_START = ny(0.14)

def _hex(h):
    h = h.lstrip('#')
    return int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)

def add_textbox(slide, x, y, w, h, text, sz=14, bold=False, color=C_B, align=PP_ALIGN.LEFT, ls=1.2):
    b = slide.shapes.add_textbox(x, y, w, h)
    tf = b.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text
    run.font.name = '微软雅黑'
    run.font.size = Pt(sz)
    run.font.bold = bold
    run.font.color.rgb = color
    p.alignment = align
    p.line_spacing = Pt(sz * ls)
    return b

def add_img(slide, path, x, y, w, h):
    if os.path.exists(path):
        slide.shapes.add_picture(path, x, y, w, h)

def cover_bg(slide):
    for f, px, py, pw, ph in [
        ('/tmp/ppt_assets/3_自定义版式_0.png', 0, 0, 1, 1),
        ('/tmp/ppt_assets/3_自定义版式_2.png', 0.762, 0.0558, 0.1884, 0.1032),
    ]:
        add_img(slide, f, nx(px), ny(py), nw(pw), nh(ph))

def content_bg(slide):
    for f, px, py, pw, ph in [
        ('/tmp/ppt_assets/统一模板_2.png', 0, 0.2333, 0.7691, 0.7691),
        ('/tmp/ppt_assets/统一模板_1.png', 0.8316, 0.0233, 0.1405, 0.0770),
    ]:
        add_img(slide, f, nx(px), ny(py), nw(pw), nh(ph))

# ── Parse markdown outline ─────────────────────────────────────────

def parse_markdown(path: str) -> list[dict]:
    """Parse outline markdown into list of slide dicts.

    Returns [{"type":"cover|transition|content|end",
              "title": "...",
              "body": ["line1","line2"]}, ...]
    """
    slides = []
    raw = open(path, encoding='utf-8').read()
    current_chapter = ""
    current_title = ""
    current_body = []
    last_was_header = False

    for line in raw.split('\n'):
        if line.startswith('## ') and not line.startswith('### '):
            # Save previous slide
            if current_title:
                slides.append(_make_slide(current_title, current_body, current_chapter))
                current_body = []

            # ## line
            h = line.lstrip('# ').strip()
            if h == '封面':
                slides.append({"type": "cover", "title": "封面", "body": []})
            elif h == '目录':
                slides.append({"type": "toc", "title": "目录", "body": []})
            elif h == '结尾页':
                slides.append({"type": "end", "title": "结尾页", "body": []})
            else:
                # Section divider – capture chapter name
                current_chapter = re.sub(r'^[第][一二三四五六七八九十]+章[：:]', '', h).strip()
            current_title = ""
            last_was_header = True

        elif line.startswith('### '):
            if current_title:
                slides.append(_make_slide(current_title, current_body, current_chapter))
                current_body = []

            h = line.lstrip('# ').strip()
            current_title = h
            last_was_header = True

        elif line.strip() and not line.startswith('---') and not line.startswith('|') and not line.startswith('```'):
            # Body line
            text = line.strip().lstrip('- *').strip()
            if text and not text.startswith('|') and not text.startswith('- |'):
                current_body.append(text)
            last_was_header = False

        elif not line.strip():
            last_was_header = False

    if current_title:
        slides.append(_make_slide(current_title, current_body, current_chapter))

    return slides


def _make_slide(title: str, body: list[str], chapter: str) -> dict:
    slide_type = "content"
    if title.startswith("转场页"):
        slide_type = "transition"
    elif title == "封面" or title == "目录":
        slide_type = title if title in ("cover", "toc") else "content"
    return {"type": slide_type, "title": re.sub(r'^转场页[：:]', '', title).strip(),
            "body": body, "chapter": chapter}


# ── Generate slide ────────────────────────────────────────────────────

def generate_slide(prs, slide_data: dict, spec_dir: str, assets_ready: bool):
    """Add one slide to prs based on slide_data."""
    st = slide_data["type"]
    title = slide_data["title"]
    body = slide_data["body"]
    s = prs.slides.add_slide(prs.slide_layouts[6])

    if st == "cover":
        cover_bg(s)
        add_textbox(s, nx(0.1137), ny(0.0979), nw(0.7882), nh(0.4),
                    "第四课\n图片素材生成、PPT 生成与 Agent 初探",
                    sz=40, color=C_B, align=PP_ALIGN.CENTER, ls=1.5)
        add_textbox(s, nx(0.077), ny(0.625), nw(0.846), nh(0.18),
                    "联通支付AI办公效能提升系列课程 · 第四课\n[日期] · [讲师姓名]",
                    sz=20, bold=True, color=C_B, align=PP_ALIGN.CENTER, ls=1.3)
        return

    if st == "end":
        cover_bg(s)
        text = title if title != "结尾页" else "感谢参与，下次见！"
        add_textbox(s, CX, ny(0.35), CW, nh(0.15), text,
                    sz=42, bold=True, color=C_B, align=PP_ALIGN.CENTER, ls=1.5)
        add_textbox(s, CX, ny(0.52), CW, nh(0.08), "第四课：图片素材生成、PPT 生成与 Agent 初探  结束",
                    sz=20, color=C_D, align=PP_ALIGN.CENTER)
        if body:
            add_textbox(s, CX, ny(0.65), CW, nh(0.06),
                        "  |  ".join(body[:3]),
                        sz=14, color=C_D, align=PP_ALIGN.CENTER)
        return

    if st == "transition":
        content_bg(s)
        add_textbox(s, nx(0.021), ny(0.049), BW, nh(0.2),
                    title, sz=36, bold=True, color=C_B, align=PP_ALIGN.LEFT, ls=1.5)
        if body:
            add_textbox(s, BL, ny(0.35), BW, nh(0.4),
                        "\n".join(body[:4]), sz=18, color=C_D, align=PP_ALIGN.LEFT, ls=1.5)
        return

    # Content page
    content_bg(s)
    add_textbox(s, nx(0.021), ny(0.049), BW, nh(0.08), title,
                sz=24, bold=True, color=C_B)
    by = BODY_START
    for item in body:
        add_textbox(s, BL, by, BW, Emu(500000), item, sz=14, color=C_D, ls=1.3)
        by += Emu(600000)


# ── Main ──────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--spec', required=True, help='Spec directory')
    parser.add_argument('--outline', required=True, help='Markdown outline file')
    parser.add_argument('-o', '--output', default='output.pptx')
    args = parser.parse_args()

    slides = parse_markdown(args.outline)
    print(f"Parsed {len(slides)} slides from outline:")
    for i, s in enumerate(slides):
        print(f"  {i+1}. [{s['type']:10s}] {s['title'][:50]}")

    # Count by type
    counts = {}
    for s in slides:
        counts[s['type']] = counts.get(s['type'], 0) + 1

    prs = Presentation()
    prs.slide_width = SW
    prs.slide_height = SH

    assets_ready = os.path.exists('/tmp/ppt_assets/3_自定义版式_0.png')
    if not assets_ready:
        print("WARNING: Assets not found in /tmp/ppt_assets/. Run spec extract first.")
    else:
        print(f"Assets ready: {assets_ready}")

    for slide_data in slides:
        generate_slide(prs, slide_data, args.spec, assets_ready)

    prs.save(args.output)
    print(f"\nSaved: {args.output}")
    print(f"Slides: {len(slides)} ({', '.join(f'{k}={v}' for k, v in counts.items())})")


if __name__ == '__main__':
    main()
