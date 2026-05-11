#!/usr/bin/env python3
"""Generate PPTX from markdown outline, following spec data."""
import argparse, re, os, json, sys
from pathlib import Path
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

SW, SH = 12192000, 6858000
FONT = '微软雅黑'

def nx(v): return int(v * SW)
def ny(v): return int(v * SH)
def nw(v): return int(v * SW)
def nh(v): return int(v * SH)

# ── Spec ────────────────────────────────────────────────────────────────

def load_spec(spec_dir):
    d = Path(spec_dir)
    spec = json.load(open(d / 'spec.json'))
    layouts = {}
    for l in spec['layouts']:
        key = f"{l['page_type']}_{l.get('layout_sub_type','full')}"
        layouts[key] = json.load(open(d / l['file']))
    return spec['palette'], spec['typography'], layouts

def first_elem_by_role(elements, roles):
    for r in roles:
        for e in elements:
            if e.get('role') == r or (not e.get('role') and r == 'title'):
                return e
    return elements[0] if elements else {}

def elem_val(elem, field, default):
    ts = elem.get('text_style', {}) or {}
    if field == 'font':      return ts.get('font_family', FONT) or FONT
    if field == 'size':      return ts.get('font_size_pt', default) or default
    if field == 'weight':    return ts.get('font_weight', 'normal') or 'normal'
    if field == 'color':     return ts.get('font_color', f'#{default}') if isinstance(default, int) else ts.get('font_color', default) or default
    if field in ('x','y','w','h'):
        return elem.get('position', {}).get(field, default)
    return default

def _hex(h):
    h = h.lstrip('#')
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)

# ── PPTX ─────────────────────────────────────────────────────────────────

def textbox(slide, x_emu, y_emu, w_emu, h_emu, text, size, bold=False, color=None, align=PP_ALIGN.LEFT, ls=1.5):
    box = slide.shapes.add_textbox(x_emu, y_emu, w_emu, h_emu)
    tf = box.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = align; p.line_spacing = Pt(size * ls)
    r = p.add_run(); r.text = text; r.font.name = FONT; r.font.size = Pt(size)
    r.font.bold = bold
    if color: r.font.color.rgb = color
    return box

def auto_textbox(slide, x, y, w, lines, size, bold=False, color=None, align=PP_ALIGN.LEFT, ls=1.5):
    """Auto-height text box for multiple lines."""
    num_lines = len(lines)
    h = max(Emu(300000), Emu(int(num_lines * size * ls * 18000)))
    text = '\n'.join(lines)
    return textbox(slide, Emu(int(x)), Emu(int(y)), Emu(int(w)), h, text, size, bold, color, align, ls)

# ── Outline parser ───────────────────────────────────────────────────────

def parse_outline(path):
    raw = open(path, encoding='utf-8').read()
    slides = []
    chapter = ""; ch_num = 0; title = ""; body = []
    lines = raw.split('\n')

    for i, line in enumerate(lines):
        if line.startswith('## ') and not line.startswith('### '):
            # Save previous
            if title:
                slides.append({'type': _type(title), 'title': title, 'body': body, 'chapter': chapter, 'ch_num': ch_num})
                body = []
            h = line.lstrip('# ').strip()
            if h == '封面':
                # Look ahead for title/subtitle
                for j in range(i+1, min(i+10, len(lines))):
                    t = lines[j].strip()
                    if t.startswith('- 主标题：'): title = t.replace('- 主标题：', '').strip()
                    elif t.startswith('- 副标题：'): body.append(t.replace('- 副标题：', '').strip())
                    elif t.startswith('- '): body.append(t[2:])
                slides.append({'type': 'cover', 'title': title, 'body': body, 'chapter': '', 'ch_num': 0})
                title = ""; body = []
            elif h == '目录':
                slides.append({'type': 'toc', 'title': '目录', 'body': [], 'chapter': '', 'ch_num': 0})
            elif h == '结尾页':
                for j in range(i+1, min(i+10, len(lines))):
                    t = lines[j].strip()
                    if t.startswith('- '): body.append(t[2:])
                slides.append({'type': 'end', 'title': '结尾页', 'body': body, 'chapter': '', 'ch_num': 0})
                body = []
            else:
                ch_num += 1
                chapter = re.sub(r'^[第][一二三四五六七八九十]+章[：:]', '', h).strip()
            title = ""
        elif line.startswith('### '):
            if title:
                slides.append({'type': _type(title), 'title': title, 'body': body, 'chapter': chapter, 'ch_num': ch_num})
                body = []
            title = line.lstrip('# ').strip()
        elif line.strip() and not line.startswith(('---', '|', '```')):
            t = line.strip().lstrip('- *').strip()
            if t and len(t) > 3: body.append(t)
    if title:
        slides.append({'type': _type(title), 'title': title, 'body': body, 'chapter': chapter, 'ch_num': ch_num})
    return slides

def _type(title):
    if title.startswith('转场页'): return 'transition'
    return 'content'

# ── Main ──────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--spec', required=True); parser.add_argument('--outline', required=True)
    parser.add_argument('-o', '--output', default='output.pptx')
    args = parser.parse_args()

    palette, typo, layouts = load_spec(args.spec)
    slides = parse_outline(args.outline)
    print(f"Outline: {len(slides)} slides | Spec: {len(layouts)} layouts")

    prs = Presentation(); prs.slide_width = SW; prs.slide_height = SH

    # Get cover and content layout elements from spec
    cover_elems, content_elems = [], []
    for k, v in layouts.items():
        if 'cover' in k: cover_elems = v.get('elements', [])
        elif 'content' in k and not content_elems: content_elems = v.get('elements', [])

    for sd in slides:
        st = sd['type']
        elems = cover_elems if st in ('cover', 'end') else content_elems
        s = prs.slides.add_slide(prs.slide_layouts[6])

        if st in ('cover', 'end'):
            # Use cover layout element positions
            title_e = first_elem_by_role(elems, ['title'])
            subtitle_e = first_elem_by_role(elems, ['subtitle'])

            tx = nx(elem_val(title_e, 'x', 0.114))
            ty = ny(elem_val(title_e, 'y', 0.098))
            tw = nw(elem_val(title_e, 'w', 0.788))
            tsz = elem_val(title_e, 'size', 40)
            tcol_rgb = RGBColor(*_hex(elem_val(title_e, 'color', '4472C4')))

            title_text = sd['title'] if st == 'cover' else "感谢参与，下次见！"
            if st == 'end':
                tx, ty, tw, tsz = nx(0.10), ny(0.30), nw(0.80), 42

            auto_textbox(s, tx, ty, tw,
                        title_text.split('\n') if '\n' in title_text else [title_text],
                        tsz, True, tcol_rgb, PP_ALIGN.CENTER)

            if st == 'cover' and sd.get('body'):
                sx = nx(elem_val(subtitle_e, 'x', 0.077))
                sy = ny(elem_val(subtitle_e, 'y', 0.625))
                sw = nw(elem_val(subtitle_e, 'w', 0.846))
                ssz = elem_val(subtitle_e, 'size', 20)
                scol = RGBColor(*_hex(elem_val(subtitle_e, 'color', '4472C4')))
                auto_textbox(s, sx, sy, sw, sd['body'][:1], ssz, True, scol, PP_ALIGN.CENTER)
            elif st == 'end' and sd.get('body'):
                auto_textbox(s, nx(0.15), ny(0.50), nw(0.70), sd['body'][:3],
                            14, False, RGBColor(0xFF,0xFF,0xFF), PP_ALIGN.CENTER, 1.5)
            continue

        # Content/Transition pages
        title_e = first_elem_by_role(elems, ['title'])
        body_e = first_elem_by_role(elems, ['body'])

        tx = nx(elem_val(title_e, 'x', 0.021))
        ty = ny(elem_val(title_e, 'y', 0.049))
        tw = nw(elem_val(title_e, 'w', 0.95))
        tsz = elem_val(title_e, 'size', 24)
        tcol = RGBColor(*_hex(elem_val(title_e, 'color', '0070C0')))

        if st == 'transition':
            tsz = 28; ty = ny(0.15)
            # Chapter number
            auto_textbox(s, nx(0.10), ny(0.08), nw(0.15),
                        [f"{sd['ch_num']:02d}"], 64, True, tcol, PP_ALIGN.LEFT)

        auto_textbox(s, tx, ty, tw,
                    [sd['title']],
                    tsz, True, tcol, PP_ALIGN.LEFT if st != 'cover' else PP_ALIGN.CENTER)

        # Body (cap at 4 items for space)
        bx = nx(elem_val(body_e, 'x', 0.04))
        by = ny(elem_val(body_e, 'y', 0.12))
        bw = nw(elem_val(body_e, 'w', 0.92))
        bsz = elem_val(body_e, 'size', 14)
        bcol_str = elem_val(body_e, 'color', '333333')
        bcol = RGBColor(*_hex(bcol_str))

        by_emu = int(by)
        max_y = int(SH * 0.95)
        for item in sd.get('body', [])[:4]:
            if by_emu > max_y: break
            auto_textbox(s, int(bx), by_emu, int(bw), [item[:120]], bsz, False, bcol, PP_ALIGN.LEFT, 1.5)
            chars = len(item)
            chars_per_line = max(1, bw / (bsz * 500))
            lines = max(1, int(chars / chars_per_line) + 1)
            by_emu += int(lines * bsz * 1.6 * 18000)

    prs.save(args.output)
    print(f"Saved: {args.output} ({len(prs.slides)} slides)")

    # Run validator
    from ppt_skill.slide_validator import validate_presentation
    prs2 = Presentation(args.output)
    issues = validate_presentation(prs2)
    b = sum(1 for vs in issues.values() for v in vs if v.severity == 'BLOCKER')
    w = sum(1 for vs in issues.values() for v in vs if v.severity == 'WARNING')
    print(f"Validator: {b} blockers, {w} warnings")
    sys.exit(1 if b > 0 else 0)

if __name__ == '__main__':
    main()
