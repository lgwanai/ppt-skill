#!/usr/bin/env python3
"""确定性分页（### 切割），LLM 分类+精确匹配（禁连续重复模板）"""
import json, sys, os, re, requests

from ppt_skill.config import load_config, get_llm_config
_cfg = load_config()
_llm = get_llm_config(_cfg)
LLM_MODEL = _llm["model"]
LLM_API_KEY = _llm["api_key"] or os.environ.get("LLM_API_KEY", "")
LLM_API_BASE = _llm["api_base"] or os.environ.get("LLM_API_BASE", "")

def call_llm(system, user):
    resp = requests.post(
        f"{LLM_API_BASE}/chat/completions",
        headers={"Content-Type": "application/json", "Authorization": f"Bearer {LLM_API_KEY}"},
        json={"model": LLM_MODEL, "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": user}],
            "temperature": 0.1, "max_tokens": 32000},
        timeout=180)
    resp.raise_for_status()
    raw = resp.json()["choices"][0]["message"]["content"]
    raw = re.sub(r"^```(?:json)?\s*", "", raw.strip())
    raw = re.sub(r"\s*```$", "", raw)
    return raw

def split_pages(text):
    pages = []
    lines = text.split('\n')
    cur = []
    for line in lines:
        s = line.strip()
        is_h1 = s.startswith('# ') and not s.startswith('## ')
        is_h2 = s.startswith('## ') and not s.startswith('### ')
        is_h3 = s.startswith('### ')
        if is_h2 or is_h3:
            if cur:
                pages.append('\n'.join(cur).strip())
                cur = []
        cur.append(line)
    if cur:
        pages.append('\n'.join(cur).strip())
    
    # Merge H1 into next page (it's the document title, not a standalone page)
    merged = []
    i = 0
    while i < len(pages):
        if pages[i].strip().startswith('# ') and not pages[i].strip().startswith('## '):
            if i + 1 < len(pages):
                merged.append(pages[i] + '\n\n' + pages[i+1])
                i += 2
                continue
        merged.append(pages[i])
        i += 1
    pages = merged
    return pages

def count_items(text):
    """Count list items: bullets, numbered, bold headers, table rows"""
    # Lines starting with - * • or digit.
    bullets = len(re.findall(r'^[ \t]*[-*•] ', text, re.MULTILINE))
    numbered = len(re.findall(r'^[ \t]*\d+[\.\)] ', text, re.MULTILINE))
    # Bold headers like **矛盾一：** as items
    bolds = len(re.findall(r'\*\*[^*]+[：:]\*\*', text))
    # Table rows (start with |)
    trows = len(re.findall(r'^\|.+\|', text, re.MULTILINE)) - 1  # minus header
    total = bullets + numbered + bolds + trows
    return max(total, 1)

def main(content_path, analysis_path):
    with open(content_path, 'r', encoding='utf-8') as f:
        content = f.read()
    with open(analysis_path, 'r', encoding='utf-8') as f:
        templates = json.load(f)
    
    # Step 1: 分页
    raw_pages = split_pages(content)
    pages = []
    for rp in raw_pages:
        title = rp.split('\n')[0].lstrip('# ').strip()
        pages.append({
            "title": title,
            "content": rp,
            "item_count": count_items(rp)
        })
    
    # Step 2: 分类 + 内容结构类型判断
    print(f"📄 {len(pages)} 页, 分类中...")
    
    pages_head = [{"idx": i, "title": p["title"], "item_count": p["item_count"],
                    "head": p["content"][:400]} for i, p in enumerate(pages)]
    
    system1 = """你是 PPT 分类专家。对每页判断 page_type 和 structure。
page_type: cover/目录/章节开头页/content/结束页
structure: {"type": "总分|并列|流程|循环|金字塔|无", "items": 数量或null}
规则：
- ## 封面/结尾页 → cover/结束页
- ## 目录 → 目录, items=列表项数
- ## 第X章/转场页 → 章节开头页
- ### 正文 → content, 类型和items根据内容判断（表格→并列, 步骤→流程, 矛盾→并列）
输出: [{"idx": 0, "page_type": "content", "structure": {"type": "并列", "items": 3}}, ...]"""

    raw1 = call_llm(system1, json.dumps(pages_head, ensure_ascii=False))
    try: results = json.loads(raw1)
    except: results = json.loads(re.search(r'\[[\s\S]*\]', raw1).group(0))
    
    rmap = {r["idx"]: r for r in results}
    for i, p in enumerate(pages):
        r = rmap.get(i, {})
        p["page_type"] = r.get("page_type", "content")
        p["structure"] = r.get("structure", {"type": "无", "items": None})
    
    # Step 3: 模板匹配（禁连续重复）
    print(f"🔗 匹配模板...")
    
    tpl_list = [{
        "page": t["page"], "type": t["page_type"],
        "structure": t.get("structure", {}),
        "tpl_items": t.get("structure", {}).get("items")
    } for t in templates if not t.get("error")]
    
    # Build page info for matching
    p_info = [{"idx": i, "page_type": p["page_type"], "title": p["title"],
               "structure": p["structure"]} for i, p in enumerate(pages)]
    
    system2 = """你是 PPT 模板匹配专家。为每页匹配一个 template_page。
规则（按优先级）：
1. page_type 必须匹配（cover↔cover, 目录↔目录, 章节开头页↔章节开头页, content↔content, 结束页↔结束页）
2. structure.type 优先匹配（并列↔并列, 总分↔总分, 流程↔流程）
3. items 尽量接近（差值最小优先）
4. **关键**：相邻两页（idx 相差1）不可用相同 template_page！
5. 没有合适模板时 template_page 填 null

输出: [{"idx": 0, "template_page": 7, "reason": "并列3项, 项数匹配"}, ...]"""

    raw2 = call_llm(system2,
        f"页面:\n{json.dumps(p_info, ensure_ascii=False)}\n\n模板:\n{json.dumps(tpl_list, ensure_ascii=False)}")
    try: matches = json.loads(raw2)
    except: matches = json.loads(re.search(r'\[[\s\S]*\]', raw2).group(0))
    
    # Step 4: 合并
    mmap = {m["idx"]: m for m in matches}
    for i, p in enumerate(pages):
        m = mmap.get(i, {})
        p["template_page"] = m.get("template_page")
        p["match_reason"] = m.get("reason", "")
    
    # Step 5: 验证无连续重复
    for i in range(1, len(pages)):
        if pages[i]["template_page"] and pages[i]["template_page"] == pages[i-1]["template_page"]:
            print(f"  ⚠️  P{i} 和 P{i+1} 重复模板 {pages[i]['template_page']}，需手动调整")
    
    out_path = content_path.replace('.md', '_pages.json').replace('.txt', '_pages.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(pages, f, ensure_ascii=False, indent=2)
    
    print(f"\n=== {len(pages)} 页 ===")
    for p in pages:
        s = p.get("structure", {})
        st = f"{s.get('type','?')}/{s.get('items','?')}"
        print(f"  [{p['template_page'] or ' -':>2}] {p['page_type']:12} | {st:10} | {p['title'][:45]}")
    print(f"\n保存至: {out_path}")

if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("用法: python match_template.py <内容.md> <分析.json>")
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])
