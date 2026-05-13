#!/usr/bin/env python3
"""使用 VL 模型解读每一页 PPT PNG，输出页面分析报告"""
import os, sys, json, base64, re, requests, glob, time

# ── 配置 ──
from ppt_skill.config import load_config, get_vl_config
_cfg = load_config()
_vl = get_vl_config(_cfg)
VL_PROVIDER = _vl.get("provider", "openai")
VL_MODEL = _vl["model"]
VL_API_KEY = _vl["api_key"] or os.environ.get("VL_API_KEY", "")
VL_API_BASE = _vl["api_base"] or os.environ.get("VL_API_BASE", "")
VL_MAX_TOKENS = int(_vl.get("max_tokens", "131072"))

SYSTEM_PROMPT = """你是一个 PPT 页面分析专家。分析每一页幻灯片截图，输出严格的 JSON 格式。

规则：
1. page_type: 页面类型，必须是以下之一：
   "cover" | "目录" | "章节开头页" | "内容页" | "结束页"

2. layout: 用 ASCII 字符描述页面布局，格式如：
   ┌─────────┬──────────┐
   │  标题    │  图片     │
   │  文字    │  文字     │
   └─────────┴──────────┘
   使用 │ ─ ┌ ┐ └ ┘ 等字符画出块状布局，标注每个块的内容类型（标题/图片/文字/图表/图标/空白）

3. structure: 页面适合表达的内容结构，格式如：
   {"type": "总分|流程|并列|循环|金字塔|无", "items": 数量或null}

   示例：
   - 总分结构，3个分项 → {"type": "总分", "items": 3}
   - 并列结构，4个 → {"type": "并列", "items": 4}
   - 流程结构（箭头/步骤）→ {"type": "流程", "items": 4}
   - 纯图或纯字无结构 → {"type": "无", "items": null}

输出 JSON 格式：
{
  "page_type": "内容页",
  "layout": "┌────┬─────────┐\\n│标题│  图片    │\\n│文字│  文字    │\\n└────┴─────────┘",
  "structure": {"type": "总分", "items": 3}
}

只输出 JSON，不要其他内容。"""

def encode_image(path):
    with open(path, 'rb') as f:
        return base64.b64encode(f.read()).decode('utf-8')

def analyze_page(image_path, max_retries=3):
    img_b64 = encode_image(image_path)
    ext = os.path.splitext(image_path)[1].lower().replace('.', '')
    if ext == 'jpg': ext = 'jpeg'
    
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {VL_API_KEY}"}
    payload = {
        "model": VL_MODEL,
        "messages": [{
            "role": "system", "content": SYSTEM_PROMPT
        }, {
            "role": "user",
            "content": [{
                "type": "image_url",
                "image_url": {"url": f"data:image/{ext};base64,{img_b64}"}
            }, {
                "type": "text",
                "text": "分析这张幻灯片"
            }]
        }],
        "temperature": 0.1,
        "max_tokens": VL_MAX_TOKENS
    }
    
    for attempt in range(1, max_retries + 1):
        print(f"  调用 VL 模型 ({VL_MODEL}) 尝试 {attempt}/{max_retries}...")
        try:
            resp = requests.post(f"{VL_API_BASE}/chat/completions",
                                headers=headers, json=payload, timeout=300)
            resp.raise_for_status()
            msg = resp.json()["choices"][0]["message"]
            content = msg.get("content", "") or msg.get("reasoning_content", "")
            content = re.sub(r"^```(?:json)?\s*", "", content.strip())
            content = re.sub(r"\s*```$", "", content)
            try:
                return json.loads(content)
            except:
                m = re.search(r'\{[\s\S]*\}', content)
                if m:
                    return json.loads(m.group(0))
                return {"error": content[:200]}
        except requests.Timeout:
            if attempt < max_retries:
                print(f"    超时, {3*attempt}s 后重试...")
                time.sleep(3 * attempt)
            else:
                raise
        except requests.HTTPError as e:
            if e.response.status_code >= 500 and attempt < max_retries:
                print(f"    服务器错误, {3*attempt}s 后重试...")
                time.sleep(3 * attempt)
            else:
                raise
        except (requests.ConnectionError, requests.RequestException) as e:
            if attempt < max_retries:
                print(f"    连接错误, {3*attempt}s 后重试...")
                time.sleep(3 * attempt)
            else:
                raise

def main(png_dir, output_file="page_analysis.json"):
    png_files = sorted(glob.glob(os.path.join(png_dir, '*.png')))
    if not png_files:
        print(f'未找到 PNG 文件: {png_dir}')
        return
    
    print(f'共 {len(png_files)} 张图片')
    results = []
    
    for i, f in enumerate(png_files, 1):
        name = os.path.basename(f)
        print(f'\n[{i}/{len(png_files)}] {name}')
        try:
            analysis = analyze_page(f)
            analysis['page'] = i
            analysis['file'] = name
            results.append(analysis)
            print(f'   类型: {analysis.get("page_type","?")}')
            print(f'   结构: {json.dumps(analysis.get("structure",{}), ensure_ascii=False)}')
        except Exception as e:
            print(f'   ❌ 错误: {e}')
            results.append({"page": i, "file": name, "error": str(e)})
    
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f'\n分析结果已保存: {output_file}')

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('用法: python vl_analyze.py <png目录> [输出json]')
        sys.exit(1)
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "page_analysis.json")
