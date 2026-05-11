#!/usr/bin/env python3
"""
PPTX Generator v2 — Clean layout following prompt-ppt-layout.md rules.

Key improvements:
1. Each body item is a self-contained text box with proper auto-height
2. W/P/S hierarchy: nav(9pt light) → point(22pt bold accent) → body(12pt dark)
3. Consistent spacing between elements  
4. Stops at page bottom margin to prevent overflow
5. Color blocks, separator lines, chapter numbers on transition pages
6. Reads spec JSON for element positions and style guidance
"""
import argparse, re, os, json, sys
from pathlib import Path
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

SW, SH = 12192000, 6858000
FONT = '微软雅黑'

# ── Layout constants ──
MARGIN_L = int(SW * 0.10)
MARGIN_R = int(SW * 0.10)
CONTENT_W = SW - MARGIN_L - MARGIN_R
MARGIN_TOP = int(SH * 0.06)
MARGIN_BOT = int(SH * 0.06)
LINE_SPACING = 1.5

# Colors — single accent from spec palette
ACCENT = RGBColor(0x44, 0x72, 0xC4)
DARK   = RGBColor(0x33, 0x33, 0x33)
LIGHT  = RGBColor(0x99, 0x99, 0x99)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
BG     = RGBColor(0xFA, 0xFA, 0xFA)

# Font sizes per layout prompt hierarchy
W_SIZE = 9      # navigation
P_SIZE = 22     # core point
S_SIZE = 12     # body text
NOTE_SIZE = 8   # footnotes

def _emu(v): return Emu(int(v))

# ── Spec loader ────────────────────────────────────────────────────

def load_spec(spec_dir):
    d = Path(spec_dir)
    sp = json.load(open(d / 'spec.json'))
    layouts = {}
    for l in sp.get('layouts', []):
        key = f"{l['page_type']}_{l.get('layout_sub_type','')}"
        layouts[key] = json.load(open(d / l['file']))
    return sp, layouts

def spec_val(elem, field, default):
    ts = elem.get('text_style', {}) or {}
    pos = elem.get('position', {}) or {}
    if field in ('x','y','w','h'): return pos.get(field, default)
    if field == 'font':     return ts.get('font_family', FONT) or FONT
    if field == 'size':     return ts.get('font_size_pt', default) or default
    if field == 'weight':   return True if (ts.get('font_weight') or '') == 'bold' else False
    if field == 'color':    c = ts.get('font_color', default) or default; return c
    return default

def _hx(h): h = h.lstrip('#'); return int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)

# ── PPTX helpers ─────────────────────────────────────────────────────

def textbox(s, x, y, w, lines, size, bold=False, color=DARK, align=PP_ALIGN.LEFT, ls=LINE_SPACING):
    """Create text box with proper auto-height. Returns bottom y position used."""
    if not lines: return y
    # Count total display lines (accounting for CJK wrapping)
    cpl = max(1, int(w / (size * 800)))  # chars per line (CJK estimate)
    total_lines = sum(max(1, -(-len(l) // cpl)) for l in lines)
    h = max(Emu(250000), Emu(int(total_lines * size * ls * 18000)))
    box = s.shapes.add_textbox(Emu(x), Emu(y), Emu(w), h)
    tf = box.text_frame; tf.word_wrap = True
    para = tf.paragraphs[0]; para.alignment = align; para.line_spacing = Pt(size * ls)
    r = para.add_run(); r.text = lines[0]; r.font.name = FONT; r.font.size = Pt(size)
    r.font.bold = bold; r.font.color.rgb = color
    for line in lines[1:]:
        p = tf.add_paragraph(); p.alignment = align; p.line_spacing = Pt(size * ls)
        r = p.add_run(); r.text = line; r.font.name = FONT; r.font.size = Pt(size)
        r.font.color.rgb = color
    return y + h

def rect(s, x, y, w, h, fill, alpha=False):
    shape = s.shapes.add_shape(1, Emu(x), Emu(y), Emu(w), Emu(h))
    shape.fill.solid(); shape.fill.fore_color.rgb = fill; shape.line.fill.background()
    return shape

def hline(s, x, y, w, color=ACCENT, width=Emu(13000)):
    s.shapes.add_connector(1, Emu(x), Emu(y), Emu(x+w), Emu(y)).line.color.rgb = color

# ── Outline parser ────────────────────────────────────────────────────

def parse_outline(path):
    raw = open(path, encoding='utf-8').read(); lines = raw.split('\n')
    slides = []; chapter = ""; ch = 0; title = ""; body = []
    for i, line in enumerate(lines):
        if line.startswith('## ') and not line.startswith('### '):
            if title: slides.append({'t':_type(title),'ti':title,'b':body,'ch':chapter,'n':ch}); body = []
            h = line.lstrip('# ').strip()
            if h == '封面':
                ti = h; sb = []
                for j in range(i+1, min(i+10, len(lines))):
                    t = lines[j].strip()
                    if t.startswith('- 主标题：'): ti = t.replace('- 主标题：','').strip()
                    elif t.startswith('- 副标题：'): sb.append(t.replace('- 副标题：','').strip())
                    elif t.startswith('- '): sb.append(t[2:])
                slides.append({'t':'cover','ti':ti,'b':sb,'ch':'','n':0}); title = ""; body = []
            elif h == '目录': slides.append({'t':'toc','ti':'目录','b':[],'ch':'','n':0})
            elif h == '结尾页':
                eb = []
                for j in range(i+1, min(i+10, len(lines))):
                    t = lines[j].strip()
                    if t.startswith('- '): eb.append(t[2:])
                slides.append({'t':'end','ti':'结尾页','b':eb,'ch':'','n':0}); body = []
            else: ch += 1; chapter = re.sub(r'^[第][一二三四五六七八九十]+章[：:]','',h).strip()
            title = ""
        elif line.startswith('### '):
            if title: slides.append({'t':_type(title),'ti':title,'b':body,'ch':chapter,'n':ch}); body = []
            title = line.lstrip('# ').strip()
        elif line.strip() and not line.startswith(('---','|','```')):
            t = line.strip().lstrip('- *').strip()
            if t and len(t) > 3: body.append(t)
    if title: slides.append({'t':_type(title),'ti':title,'b':body,'ch':chapter,'n':ch})
    return slides

def _type(t): return 'transition' if t.startswith('转场页') else 'content'

# ── Generate one slide ───────────────────────────────────────────────

def gen_slide(prs, sd, spec, layouts):
    st = sd['t']; title = sd['ti']; body = sd['b']; chapter = sd['ch']; ch_num = sd['n']
    s = prs.slides.add_slide(prs.slide_layouts[6])

    # ── Cover ──
    if st == 'cover':
        # Background
        rect(s, 0, 0, SW, int(SH*0.78), ACCENT)  # accent color upper section
        rect(s, 0, int(SH*0.78), SW, int(SH*0.22), WHITE)  # white bottom
        
        # Title
        y = textbox(s, int(SW*0.12), int(SH*0.20), int(SW*0.76),
                   title.split('\n'), 38, True, WHITE, PP_ALIGN.CENTER, 1.3)
        
        # Subtitle
        if body:
            textbox(s, int(SW*0.12), y + 20000, int(SW*0.76),
                   [body[0]], 16, False, RGBColor(0xCC,0xDD,0xEE), PP_ALIGN.CENTER, 1.3)
        
        # Bottom info
        textbox(s, int(SW*0.12), int(SH*0.82), int(SW*0.76),
               [l for l in body[1:3]], 11, False, LIGHT, PP_ALIGN.CENTER, 1.3)
        return

    # ── End ──
    if st == 'end':
        rect(s, 0, 0, SW, SH, ACCENT)
        textbox(s, int(SW*0.10), int(SH*0.28), int(SW*0.80),
               ["感谢参与，下次见！"], 42, True, WHITE, PP_ALIGN.CENTER, 1.5)
        textbox(s, int(SW*0.10), int(SH*0.45), int(SW*0.80),
               ["第四课：图片素材生成、PPT 生成与 Agent 初探"], 16, False, RGBColor(0xCC,0xDD,0xEE), PP_ALIGN.CENTER, 1.3)
        if body:
            textbox(s, int(SW*0.15), int(SH*0.58), int(SW*0.70),
                   body[:3], 12, False, WHITE, PP_ALIGN.CENTER, 1.5)
        return

    # ── All other pages ──
    # Background
    rect(s, 0, 0, SW, SH, WHITE)
    
    # W: Navigation (top, small, light) — only if chapter exists
    y = MARGIN_TOP
    if chapter:
        y = textbox(s, MARGIN_L, y, CONTENT_W, [chapter], W_SIZE, False, LIGHT)
        y += 30000

    # Transition page header
    if st == 'transition':
        # Big chapter number
        textbox(s, MARGIN_L, y, int(SW*0.15), [f"{ch_num:02d}"], 64, True, ACCENT, PP_ALIGN.LEFT)
        textbox(s, MARGIN_L, y + 60000, CONTENT_W, [title], 28, True, ACCENT)
        y += int(SH * 0.18)
        # Decorative line
        hline(s, MARGIN_L, y, int(SW*0.25), ACCENT, Emu(20000))
        y += 50000
        # Body (max 3 items)
        for item in body[:3]:
            if y > SH - MARGIN_BOT - 50000: break
            y = textbox(s, MARGIN_L, y, CONTENT_W, [item[:120]], 14, False, DARK) + 40000
        return

    # Content page
    # P: Title as core point (big, bold, accent)
    y = textbox(s, MARGIN_L, y, CONTENT_W, [title], P_SIZE, True, ACCENT)
    y += 20000
    
    # Thin separator
    hline(s, MARGIN_L, y, int(SW*0.12), ACCENT, Emu(10000))
    y += 30000

    # S: Body items — each as a separate block with subtle left border
    max_items = 5
    for i, item in enumerate(body[:max_items]):
        if y > SH - MARGIN_BOT - 40000: break
        
        # Detect sub-headers (short bold text)
        is_header = len(item) < 50 and (item.endswith('：') or item.endswith(':') or '**' in item or
                    any(item.startswith(w) for w in ['矛盾', '局限', 'Step', '①','②','③','④']))
        sz = S_SIZE + 2 if is_header else S_SIZE
        bld = is_header
        clr = ACCENT if is_header else DARK
        
        # Left color bar for key items
        if is_header or i % 3 == 0:
            rect(s, MARGIN_L - 5000, y, 4000, int(S_SIZE*LINE_SPACING*20000), ACCENT)
        
        y = textbox(s, MARGIN_L + 20000, y, CONTENT_W - 20000, [item[:120]], sz, bld, clr) + 35000

# ── Main ─────────────────────────────────────────────────────────────────

def main():
    p = argparse.ArgumentParser()
    p.add_argument('--spec', required=True); p.add_argument('--outline', required=True)
    p.add_argument('-o','--output',default='output.pptx'); p.add_argument('--loop',action='store_true')
    args = p.parse_args()

    spec, layouts = load_spec(args.spec)
    slides = parse_outline(args.outline)
    print(f"Outline: {len(slides)} slides | Spec: {len(layouts)} layouts")

    prs = Presentation(); prs.slide_width = SW; prs.slide_height = SH

    for i, sd in enumerate(slides):
        gen_slide(prs, sd, spec, layouts)

    prs.save(args.output)
    print(f"Saved: {args.output} ({len(prs.slides)} slides)")

    # Run validator
    try:
        from ppt_skill.slide_validator import validate_presentation
        prs2 = Presentation(args.output)
        issues = validate_presentation(prs2)
        b = sum(1 for vs in issues.values() for v in vs if v.severity == 'BLOCKER')
        w = sum(1 for vs in issues.values() for v in vs if v.severity == 'WARNING')
        print(f"Validator: {b} blockers, {w} warnings")
        sys.exit(1 if b > 0 else 0)
    except Exception:
        pass

if __name__ == '__main__':
    main()
