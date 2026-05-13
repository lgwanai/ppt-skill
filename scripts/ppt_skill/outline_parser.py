"""Outline parser — parse markdown course outline into page list."""

import re


def parse_outline(raw: str) -> list[dict]:
    """Parse outline markdown into page list.
    
    Returns list of page dicts:
      {index, type, title, body, chapter, ch_num}
    
    Types: cover, toc, end, content, transition
    """
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


def outlines_to_pages_file(outline_path: str, output_path: str) -> list[dict]:
    """Parse outline and write pages JSON for match_template consumption.

    Converts the parsed pages into the format expected by paginate.py:
      {page_index, title, content, page_type, item_count}
    """
    with open(outline_path, encoding='utf-8') as f:
        pages = parse_outline(f.read())

    output = []
    for p in pages:
        # Build content text from body items
        if p['type'] == 'cover':
            content = '\n'.join(p['body']) if p['body'] else p['title']
        elif p['type'] == 'toc':
            content = '\n'.join(f"{i+1}. {b}" for i, b in enumerate(p['body']))
        elif p['type'] == 'end':
            content = '\n'.join(p['body']) if p['body'] else ''
        else:
            content = '\n'.join(p['body']) if p['body'] else ''

        output.append({
            'page_index': p['index'],
            'title': p['title'],
            'content': content,
            'page_type': p['type'],
            'item_count': len(p['body']),
        })

    if output_path:
        import json
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(output, f, ensure_ascii=False, indent=2)

    return output
