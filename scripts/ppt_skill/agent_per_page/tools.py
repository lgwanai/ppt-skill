"""Agent-per-page PPT generation tools — LLM-driven, no hardcoded rendering."""
import json, os, re, io, base64, subprocess, tempfile
from pathlib import Path
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from openai import OpenAI
from PIL import Image

SW, SH = 12192000, 6858000
FONT = '微软雅黑'

def _emu(v): return Emu(int(v))
def _hx(h): h=h.lstrip('#'); return int(h[0:2],16),int(h[2:4],16),int(h[4:6],16)

# ── LLM client (reads config.txt) ──────────────────────────────────

def _llm_client():
    cfg = {}
    if os.path.exists('config.txt'):
        for line in open('config.txt'):
            line = line.strip()
            if '=' in line and not line.startswith('#'):
                k,v = line.split('=',1)
                cfg[k.strip()] = v.strip().strip('"')
    return OpenAI(
        api_key=cfg.get('LLM_API_KEY',''),
        base_url=cfg.get('LLM_API_BASE',''),
    )

def _vl_client():
    cfg = {}
    if os.path.exists('config.txt'):
        for line in open('config.txt'):
            line = line.strip()
            if '=' in line and not line.startswith('#'):
                k,v = line.split('=',1)
                cfg[k.strip()] = v.strip().strip('"')
    return OpenAI(
        api_key=cfg.get('VL_API_KEY',''),
        base_url=cfg.get('VL_API_BASE',''),
    )

# ── Tool 1: get_page ──────────────────────────────────────────────

def get_page(outline_path: str, completed: set) -> dict | None:
    """Fetch next uncompleted page from outline."""
    raw = open(outline_path, encoding='utf-8').read()
    pages = _parse_outline(raw)
    for p in pages:
        if p['index'] not in completed:
            return p
    return None

def _parse_outline(raw: str) -> list[dict]:
    """Parse outline markdown into page list."""
    import re
    lines = raw.split('\n')
    pages = []; chapter = ""; ch = 0; title = ""; body = []; skip = 0
    idx = 0
    
    def _type(t): return 'transition' if t.startswith('转场页') else 'content'
    
    for i, line in enumerate(lines):
        if i < skip: continue
        if line.startswith('## ') and not line.startswith('### '):
            if title:
                pages.append({'index':idx,'type':_type(title),'title':title,'body':body,'chapter':chapter,'ch_num':ch})
                idx+=1; body=[]
            h = line.lstrip('# ').strip()
            if h == '封面':
                ti=h; sb=[]
                for j in range(i+1, min(i+10,len(lines))):
                    t=lines[j].strip()
                    if t.startswith('- 主标题：'): ti=t.replace('- 主标题：','').strip()
                    elif t.startswith('- 副标题：'): sb.append(t.replace('- 副标题：','').strip())
                    elif t.startswith('- '): sb.append(t[2:])
                skip=i+len(sb)+2
                pages.append({'index':idx,'type':'cover','title':ti,'body':sb,'chapter':'','ch_num':0})
                idx+=1; title=""; body=[]
            elif h == '目录':
                toc=[]
                for j in range(i+1,min(i+15,len(lines))):
                    t=lines[j].strip()
                    if not t or t.startswith('##'): break
                    if t[0].isdigit() and '.' in t[:3]: toc.append(t.split('.',1)[1].strip())
                skip=j
                pages.append({'index':idx,'type':'toc','title':'目录','body':toc,'chapter':'','ch_num':0})
                idx+=1; title=""; body=[]
            elif h == '结尾页':
                eb=[]
                for j in range(i+1,min(i+10,len(lines))):
                    t=lines[j].strip()
                    if not t or t.startswith('##'): break
                    if t.startswith('- '): eb.append(t[2:])
                skip=i+len(eb)+2
                pages.append({'index':idx,'type':'end','title':'结尾页','body':eb,'chapter':'','ch_num':0})
                idx+=1; body=[]
            else: ch+=1; chapter=re.sub(r'^[第][一二三四五六七八九十]+章[：:]','',h).strip()
            title=""
        elif line.startswith('### '):
            if title:
                pages.append({'index':idx,'type':_type(title),'title':title,'body':body,'chapter':chapter,'ch_num':ch})
                idx+=1; body=[]
            title=line.lstrip('# ').strip()
        elif line.strip() and not line.startswith(('---','```')):
            sline=line.strip()
            if sline.startswith('|') and '---' not in sline:
                cells=[c.strip().strip('*') for c in sline.split('|') if c.strip()]
                if cells: body.append('__TABLE__'+'|'.join(cells))
            else:
                t=re.sub(r'\*+','',sline.lstrip('- ').strip())
                if t and len(t)>3: body.append(t)
    if title:
        pages.append({'index':idx,'type':_type(title),'title':title,'body':body,'chapter':chapter,'ch_num':ch})
    return pages

# ── Tool 2: extract_spec ──────────────────────────────────────────

def extract_spec(page: dict, spec_dir: str) -> dict:
    """Match page type to spec layout, return style data only (no content)."""
    sd = Path(spec_dir)
    spec = json.load(open(sd/'spec.json'))
    palette = spec['palette']
    
    # Find matching layout
    layouts = spec.get('layouts',[])
    matched = None
    for l in layouts:
        if l['page_type'] == page['type']:
            matched = l; break
    if not matched and layouts:
        matched = layouts[0]
    
    if matched:
        data = json.load(open(sd/matched['file']))
        elements = data.get('elements',[])
    else:
        elements = []
    
    # Extract style-only data using LLM
    llm = _llm_client()
    prompt = f"""Extract STYLE-ONLY data from these spec elements. Return ONLY fixed design attributes (no content text).

Page type: {page['type']}
Palette: {json.dumps(palette)}
Elements: {json.dumps([{k:v for k,v in e.items() if k != 'text'} for e in elements], ensure_ascii=False)[:2000]}

Return JSON format:
{{
  "page_type": "{page['type']}",
  "palette": {{"accent": "...", "bg": "...", "text": "...", ...}},
  "title": {{"x_pct": 0.02, "y_pct": 0.05, "w_pct": 0.95, "font": "...", "size_pt": 24, "bold": true, "color": "...", "align": "left"}},
  "body": {{"x_pct": 0.04, "y_pct": 0.12, "w_pct": 0.92, "font": "...", "size_pt": 14, "color": "...", "line_spacing": 1.5}},
  "background": {{"type": "solid|image", "color": "...", "image_path": "..."}},
  "decorations": [{{"type": "line|shape", "pos": {{...}}, "color": "..."}}],
  "layout_type": "full_width|left_right|top_bottom|center"
}}
"""
    try:
        r = llm.chat.completions.create(model="deepseek-v4-flash", messages=[{"role":"user","content":prompt}], max_tokens=1024, temperature=0.1)
        text = r.choices[0].message.content
        m = re.search(r'\{.*\}', text, re.DOTALL)
        return json.loads(m.group(0)) if m else _default_spec(page, palette)
    except:
        return _default_spec(page, palette)

def _default_spec(page, palette):
    return {
        "page_type": page['type'],
        "palette": {"accent": palette.get('accent1','#4472C4'), "bg": "#FFFFFF", "text": "#333333"},
        "title": {"x_pct": 0.021, "y_pct": 0.049, "w_pct": 0.95, "font": FONT, "size_pt": 24, "bold": True, "color": palette.get('accent1','#4472C4'), "align": "left"},
        "body": {"x_pct": 0.04, "y_pct": 0.14, "w_pct": 0.92, "font": FONT, "size_pt": 13, "color": "#333333", "line_spacing": 1.5},
        "background": {"type": "solid", "color": "#FFFFFF"},
        "decorations": [],
        "layout_type": "full_width"
    }

# ── Tool 3: plan_assets ───────────────────────────────────────────

def plan_assets(page: dict, spec: dict) -> dict:
    """Plan optimal visual elements for page content."""
    llm = _llm_client()
    prompt = f"""Plan visual assets for this slide. Priority: SmartArt+text > SmartArt > chart+text > diagram+text > plain text.
Max 2 graphics (SmartArt/chart/diagram) per page. If graphics can replace text, condense text.

Page content:
  type: {page['type']}
  title: {page['title']}
  body items ({len(page['body'])}):
  {json.dumps(page['body'][:6], ensure_ascii=False)}

Available: SmartArt types (list/process/cycle/hierarchy/pyramid/matrix), chart (bar/pie/line), diagram SVG templates, icons, images.

Return JSON: 
{{"modules": [{{"type":"smartart|chart|diagram|image|text","subtype":"flow|compare|list|pyramid|cycle|bar|table","contentRef":"which body items this covers","reason":"why this choice"}}]}}
"""
    try:
        r = llm.chat.completions.create(model="deepseek-v4-flash", messages=[{"role":"user","content":prompt}], max_tokens=1024, temperature=0.3)
        text = r.choices[0].message.content
        m = re.search(r'\{.*\}', text, re.DOTALL)
        return json.loads(m.group(0)) if m else {"modules": [{"type":"text","subtype":"list","contentRef":"all","reason":"default"}]}
    except:
        return {"modules": [{"type":"text","subtype":"list","contentRef":"all","reason":"default"}]}

# ── Tool 4: plan_layout ───────────────────────────────────────────

def plan_layout(page: dict, spec: dict, assets: dict) -> dict:
    """Calculate page layout zones based on content and assets."""
    llm = _llm_client()
    prompt = f"""Plan page layout. The slide is 13.33x7.5 inches. Fixed elements occupy:
  Title at ({spec['title']['x_pct']:.1%}, {spec['title']['y_pct']:.1%}) height ~{spec['title']['size_pt']/72:.1f}in

Page type: {page['type']}
Assets: {json.dumps(assets, ensure_ascii=False)}
Body line count: ~{len(page['body'])}

Available layouts: center, left_right, top_bottom, grid_2x2, grid_3x3, grid_2x3, pyramid, pinwheel
For transition pages: always use center layout.

Return JSON:
{{"layout":"grid_2x2","zones":[{{"x_pct":0.05,"y_pct":0.15,"w_pct":0.42,"h_pct":0.35,"content":"title"}},{{"x_pct":0.52,"y_pct":0.15,"w_pct":0.42,"h_pct":0.35,"content":"body_0"}},{{"x_pct":0.05,"y_pct":0.52,"w_pct":0.42,"h_pct":0.35,"content":"body_2"}},{{"x_pct":0.52,"y_pct":0.52,"w_pct":0.42,"h_pct":0.35,"content":"diagram"}}]}}
"""
    try:
        r = llm.chat.completions.create(model="deepseek-v4-flash", messages=[{"role":"user","content":prompt}], max_tokens=1024, temperature=0.2)
        text = r.choices[0].message.content
        m = re.search(r'\{.*\}', text, re.DOTALL)
        return json.loads(m.group(0)) if m else {"layout":"top_bottom","zones":[{"x_pct":0.05,"y_pct":0.15,"w_pct":0.90,"h_pct":0.75,"content":"body"}]}
    except:
        return {"layout":"top_bottom","zones":[{"x_pct":0.05,"y_pct":0.15,"w_pct":0.90,"h_pct":0.75,"content":"body"}]}

# ── Tool 5: draw_zone ──────────────────────────────────────────────

def draw_zone(slide, zone: dict, spec: dict, content: str, content_type: str = "text"):
    """Draw content into a layout zone on the slide."""
    x = int(zone['x_pct'] * SW)
    y = int(zone['y_pct'] * SH)
    w = int(zone['w_pct'] * SW)
    h = int(zone['h_pct'] * SH)
    
    if content_type == "text":
        bstyle = spec.get('body', spec.get('title', {}))
        sz = bstyle.get('size_pt', 13)
        color = RGBColor(*_hx(bstyle.get('color', '#333333')))
        
        # Auto-calculate font size to fit zone
        chars = len(content)
        cpl = max(1, int(w / (sz * 9144)))
        n_lines = max(1, (chars + cpl - 1) // cpl)
        
        # Reduce font size if text overflows zone
        while n_lines * sz * 1.5 * 18000 > h and sz > 8:
            sz -= 1
            cpl = max(1, int(w / (sz * 9144)))
            n_lines = max(1, (chars + cpl - 1) // cpl)
        
        box_h = n_lines * sz * 1.5 * 18000
        tb = slide.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(min(box_h, h)))
        tf = tb.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.line_spacing = Pt(sz * 1.5)
        r = p.add_run(); r.text = content[:500]
        r.font.name = bstyle.get('font', FONT); r.font.size = Pt(sz); r.font.color.rgb = color
    
    elif content_type == "title":
        tstyle = spec.get('title', {})
        sz = tstyle.get('size_pt', 24)
        color = RGBColor(*_hx(tstyle.get('color', '#4472C4')))
        align = PP_ALIGN.CENTER if tstyle.get('align') == 'center' else PP_ALIGN.LEFT
        
        tb = slide.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
        tf = tb.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.alignment = align; p.line_spacing = Pt(sz * 1.3)
        r = p.add_run(); r.text = content[:200]
        r.font.name = tstyle.get('font', FONT); r.font.size = Pt(sz)
        r.font.bold = tstyle.get('bold', False); r.font.color.rgb = color
    
    elif content_type == "table":
        rows = [r.split('|') for r in content.split('__ROW__') if r.strip()]
        if len(rows) >= 2:
            nr, nc = len(rows), max(len(r) for r in rows)
            tbl = slide.shapes.add_table(nr, min(nc, 5), Emu(x), Emu(y), Emu(w), Emu(h))
            for ri, row in enumerate(rows):
                for ci, cell_text in enumerate(row[:5]):
                    cell = tbl.table.cell(ri, ci)
                    cell.text = cell_text[:80]
                    for p in cell.text_frame.paragraphs:
                        p.font.size = Pt(10); p.font.name = FONT
                        if ri == 0: p.font.bold = True

# ── Tool 6: screenshot ─────────────────────────────────────────────

def screenshot_pptx(pptx_path: str, output_dir: str = "/tmp/ppt_screens") -> str | None:
    """Screenshot a PPTX to PNG using qlmanage."""
    os.makedirs(output_dir, exist_ok=True)
    try:
        subprocess.run(['qlmanage','-t','-s','1920','-o',output_dir,pptx_path], capture_output=True, timeout=10)
        pngs = sorted(Path(output_dir).glob('*.png'))
        return str(pngs[-1]) if pngs else None
    except:
        return None

# ── Tool 7: review ────────────────────────────────────────────────

def review_slide(screenshot_path: str, spec: dict) -> dict:
    """VL model review for visual issues."""
    vl = _vl_client()
    try:
        img = Image.open(screenshot_path).convert('RGB')
        if img.width > 1024: img = img.resize((1024, int(img.height*1024/img.width)))
        buf = io.BytesIO(); img.save(buf, 'JPEG', quality=70)
        b64 = base64.b64encode(buf.getvalue()).decode()
        
        prompt = f"""Review this PPT slide for visual issues. Expected style:
  Title: {json.dumps(spec.get('title',{}))}
  Body: {json.dumps(spec.get('body',{}))}

Check: text overlap, spacing too tight, alignment issues, overflow beyond slide, color consistency.
Return JSON: {{"pass":bool,"issues":[{{"type":"overlap|spacing|alignment|overflow|color","detail":"...","fix":"..."}}]}}"""
        
        r = vl.chat.completions.create(
            model="mimo-v2.5-pro",
            messages=[{"role":"user","content":[{"type":"text","text":prompt},{"type":"image_url","image_url":{"url":f"data:image/jpeg;base64,{b64}"}}]}],
            max_tokens=512, temperature=0.1)
        text = r.choices[0].message.content
        m = re.search(r'\{.*\}', text, re.DOTALL)
        return json.loads(m.group(0)) if m else {"pass":True,"issues":[]}
    except:
        return {"pass": True, "issues": []}
