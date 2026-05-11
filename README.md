# ppt-skill

AI 驱动的 PPT 生成技能 — 从参考 PPTX 提取设计规范 → 生成可原生编辑的 .pptx 文件。

不同于传统 PPT 生成工具只做"内容填入模板"，ppt-skill 的核心理念是 **"先理解设计，再生成内容"**：通过 VL 模型（视觉语言模型）分析参考 PPT 中每个元素的角色和关系，提取完整的布局蓝图，然后用 Agent Loop 机制逐页验证生成结果是否匹配规范。

## 应用场景

| 场景 | 说明 |
|------|------|
| **公司模板统一** | 将企业 PPT 模板提取为 spec，后续所有 PPT 自动遵循同一设计规范 |
| **PPT 风格迁移** | 客户发来一个参考 PPT，需要仿照风格做新内容 — 提取 spec 后填新内容即可 |
| **批量 PPT 生成** | 固定模板 + 内容大纲，批量产出风格一致的 PPT |
| **PPT 设计数字化** | 将设计师的 PPT 作品转化为可复用的 JSON 规范，纳入设计资产库 |
| **跨语言 PPT** | 提取中文 PPT 规范后，填入英文内容生成风格一致的英文 PPT |

## 与其他 PPT 工具的核心区别

| 维度 | ppt-skill | 传统模板工具 | AI 直接生成 |
|------|-----------|------------|------------|
| **设计来源** | 从参考 PPT 学习 | 固定模板 | 无参考，靠模型想象 |
| **还原度** | 90%+（Agent Loop 验证） | 100% 但限于模板 | 不稳定 |
| **输出格式** | 原生 PowerPoint 形状 | 原生形状 | 图片版 PPT |
| **可编辑性** | ✅ 文字/颜色/形状全可编辑 | ✅ | ❌ |
| **VL 分析** | 比对图片理解元素角色 | ❌ | ❌ |
| **多布局** | 封面/目录/内容/尾页各自独立 | 有限 | 随机 |
| **素材复用** | 自动提取背景/Logo/装饰图 | 需手动 | ❌ |

## 安装

### 前置依赖

```bash
# 1. LibreOffice — 用于将 PPTX 转为 PDF（获取真实截图）
brew install --cask libreoffice

# 2. Poppler (pdf2image) — PDF 每页转 PNG
brew install poppler

# 3. Python 包
pip install python-pptx Pillow PyYAML lxml pdf2image

# 4. VL 模型支持（可选，用于元素语义分析）
pip install openai
```

### 安装验证

```bash
# 验证 LibreOffice 可用
/Applications/LibreOffice.app/Contents/MacOS/soffice --headless --convert-to pdf --outdir /tmp /path/to/test.pptx

# 验证 pdftoppm 可用
pdftoppm -v

# 验证 CLI 可用
python scripts/ppt_cli.py --help
```

## 配置

复制 `config.example.txt` 为 `config.txt`，配置 VL 模型：

```ini
# VL 模型用于 PPT 元素语义分析
VL_ENABLED=true
VL_PROVIDER=openai        # openai | anthropic | gemini | ollama
VL_MODEL=gpt-4o           # 或 doubao-seed-dream-1.5-pro, claude-3.5-sonnet 等
VL_API_KEY=sk-...
VL_API_BASE=              # 兼容 OpenAI API 格式的自定义地址
VL_MAX_TOKENS=4096

# LLM 用于内容大纲生成和布局设计
LLM_PROVIDER=openai
LLM_MODEL=gpt-4o
LLM_API_KEY=sk-...
LLM_API_BASE=
LLM_MAX_TOKENS=8192
```

如果不配置 VL，系统会自动回退到程序化模式，使用基于位置/字体规则的元素分类。

## 完整工作流程

### 第一步：提取设计规范（extract-vl-spec）

```bash
python scripts/ppt_cli.py extract-vl-spec reference.pptx --name my-style
```

这一步是 ppt-skill 的核心创新。流程如下：

```
PPTX 文件
  │
  ├─ 1. python-pptx 提取每个元素的完整属性
  │    ├─ 位置、大小（归一化到 0-1）
  │    ├─ 字体（family, size, weight, color, alignment）
  │    ├─ 填充/描边（颜色、宽度、圆角）
  │    ├─ 图片尺寸、embed_id
  │    ├─ Z-order、shape name
  │    └─ 表格：行列数、每个单元格的文字/字体/填充色
  │
  ├─ 2. 检查 layout/master 层
  │    ├─ 提取背景图（全屏 `<p:pic>`）
  │    ├─ 提取 Logo/装饰元素
  │    └─ 保存到 assets/ 目录
  │
  ├─ 3. LibreOffice → PDF → pdf2image → 每页 PNG 截图
  │
  ├─ 4. VL 模型分析（每页一次调用）
  │    ├─ 输入：PNG 截图 + 结构化元素属性
  │    ├─ VL 对照图片和属性，标记每个元素：
  │    │   "元素 1 是标题"、"元素 2 是正文"、"元素 3-6 组成流程图"
  │    └─ 输出：role, type_category, visual_weight, style_notes
  │
  ├─ 5. 布局去重
  │    ├─ cover → 1 个 JSON 文件
  │    ├─ end_page → 1 个 JSON 文件
  │    └─ 内容页按布局签名分组 → 每组 1 个 JSON 文件
  │
  └─ 6. 输出到 specs/<name>/
       ├─ spec.json              # 主配置（palette, fonts, canvas, layout 索引）
       ├─ cover.json             # 封面布局
       ├─ end_page.json          # 尾页布局
       ├─ content_*.json         # 每种内容布局
       ├─ slides/                # 每页 PNG 截图
       ├─ assets/                # 背景图、Logo、装饰元素
       └─ <source>.pptx          # 源文件副本
```

#### 输出 JSON 示例

```json
{
  "spec_name": "my-style",
  "layouts": [
    {"file": "cover.json", "page_type": "cover", "slide_indices": [0]},
    {"file": "content_left_right.json", "page_type": "content", "slide_indices": [1,2,3,5]},
    {"file": "end_page.json", "page_type": "end_page", "slide_indices": [12]}
  ],
  "palette": {
    "dk1": "#000000", "lt1": "#FFFFFF", "accent1": "#4472C4", ...
  },
  "typography": {
    "heading_family": "Calibri Light", "body_family": "Calibri"
  }
}
```

每个 layout JSON 包含完整元素数据：
```json
{
  "elements": [
    {"id": 0, "element_type": "text", "semantic_role": "title",
     "position": {"x": 0.05, "y": 0.05, "w": 0.9, "h": 0.12},
     "text": "关于2026年预算申请",
     "text_style": {"font_family": "微软雅黑", "font_size_pt": 24, "font_weight": "bold"}},
    {"id": 1, "element_type": "image", "semantic_role": "decoration",
     "from_layout": "slide_layout",
     "image": {"saved_path": "assets/dec_00_00.png"}},
    {"id": 3, "element_type": "table",
     "table": {"rows": 8, "cols": 5, "col_widths": [0.14, 0.19, ...],
       "cells": [
         {"row": 0, "col": 0, "text": "价格段", "fill_color": "#5B9BD5",
          "text_style": {"font_family": "微软雅黑", "font_size_pt": 14}}
       ]
     }
    }
  ]
}
```

### 第二步：准备内容大纲

```bash
# 直接提供内容大纲 YAML
python scripts/ppt_cli.py outline --input content.txt
```

内容大纲格式：
```yaml
slides:
  - title: "2024 年度总结"
    body: ["营收增长 35%", "新客户 200+", "产品线扩展至 3 条"]
  - title: "市场分析"
    body: ["市场规模达 500 亿", "竞争对手分析", "我们的优势"]
```

### 第三步：生成 PPT

```bash
python scripts/ppt_cli.py generate-pptx \
    --spec specs/my-style \
    --outline outlines/content.yaml \
    -o output.pptx \
    --workers 4
```

生成过程有两个验证阶段：

```
Phase 1: Spec Matching（每页并行）
  ┌──────────────────────────────────────────┐
  │  Slide → 匹配 layout → 生成 SVG          │
  │  → 评估(颜色/字体/布局/背景/密度)        │
  │  → < 90% ？添加修复指令 → 重新生成       │
  │  → 最多 5 次迭代，取最高分               │
  └──────────────────────────────────────────┘
                    ↓
Phase 2: Layout Design（逐页）
  ┌──────────────────────────────────────────┐
  │  WPS 层级优化：W导航/P核心/S支撑          │
  │  排版规则：sans-serif/1.5x行距/留白       │
  │  图标建议：Bootstrap Icons               │
  └──────────────────────────────────────────┘
                    ↓
           SVG → DrawingML → .pptx
```

## AI 实现原理

### Agent Loop 机制

每一页 PPT 的生成都是一个"生成→评估→修复"的循环：

1. **生成**: 将 slide content + spec style guidance 组装成 prompt，调用 LLM 生成 SVG
2. **评估**: StyleEvaluator 从 5 个维度打分：
   - 颜色（30%）：SVG 中使用的颜色是否在 spec palette 内
   - 排版（20%）：字体 family 和大小层级是否匹配
   - 布局（30%）：元素位置 IoU 是否与 spec 对齐
   - 背景（15%）：背景色/图片是否正确
   - 密度（5%）：内容密度是否匹配（呼吸/密集/锚点）
3. **修复**: 综合得分 < 90% 时，将评估中发现的问题作为 fix instructions 输入下一轮生成
4. **迭代**: 最多 5 轮，取最高分 SVG

### VL 视觉分析

VL 模型接收两个输入：
1. **截图**: LibreOffice 渲染的精确 PNG（不是 Pillow 示意图）
2. **元素数据**: python-pptx 提取的每个元素的结构化属性列表

VL 的任务是**比对图片和属性数据**，回答：
- 这个元素在页面上看起来是什么角色？（标题/正文/背景/装饰/表格）
- 多个元素如何组合成视觉组？（"元素 4,5,6 组成一个卡片组件"）
- 整体布局模式是什么？（"两栏文本布局，左侧图标点缀"）

### 布局去重（Layout Dedup）

提取 spec 时，不是每页存一个 JSON，而是按布局签名分组：

```
签名 = PT:cover|LT:full_width|T:text2_image1|R:title1_subtitle1|G:...
      ^page_type  ^sub_type     ^元素类型统计 ^角色统计     ^位置网格
```

签名相同的页面合并为一个 layout JSON，减少冗余、提升后续生成时的匹配效率。

### 资源提取（Asset Extraction）

自动分类和保存可复用素材：

| 分类 | 条件 | 处理 |
|------|------|------|
| **背景** (background) | 面积 > 50% 幻灯片 | 保存到 `assets/bg_*.png` |
| **图标** (icon) | 面积 < 3% 且命名含 icon/logo | 保存到 `assets/ico_*.png` |
| **装饰** (decoration) | 面积 < 15% 或来自 layout 层 | 保存到 `assets/dec_*.png` |
| **内容图** (content_image) | 内容区域的中等大小图片 | 丢弃 |
| **图表** (chart) | chart 形状 | 丢弃 |

## CLI 命令速查

```bash
# ── Spec 提取 ──
python scripts/ppt_cli.py extract-spec reference.pptx --name my-style
    # 标准提取（保留文字内容，输出 YAML）
python scripts/ppt_cli.py extract-vl-spec reference.pptx --name my-style
    # VL 驱动提取（JSON + 截图 + 元素角色分析，去重）

# ── Spec 管理 ──
python scripts/ppt_cli.py list-specs
    # 列出所有已提取的 spec
python scripts/ppt_cli.py select-spec my-style
    # 设置默认 spec

# ── 内容准备 ──
python scripts/ppt_cli.py gather-content "2024 年度总结"
    # 交互式内容收集 → outlines/<name>.yaml
python scripts/ppt_cli.py outline --input content.txt
    # 从文本生成内容大纲
python scripts/ppt_cli.py list-outlines
    # 列出已保存的大纲

# ── PPT 生成 ──
python scripts/ppt_cli.py generate-pptx \
    --spec specs/my-style \
    --outline outlines/content.yaml \
    -o output.pptx \
    --workers 4

# ── 格式转换 ──
python scripts/ppt_cli.py convert slide1.svg slide2.svg -o output.pptx
    # SVG → 原生形状 PPTX
```

## 项目结构

```
ppt-skill/
├── scripts/
│   ├── ppt_cli.py                        # CLI 入口
│   └── ppt_skill/
│       ├── cli/                          # 命令实现
│       │   ├── spec_commands.py          # extract-spec, extract-vl-spec
│       │   └── generate_commands.py      # generate-pptx
│       ├── spec/                         # Spec 提取核心
│       │   ├── vl_spec_extractor.py      # VL 驱动提取管线
│       │   ├── vl_element_analyzer.py    # VL 模型调用 + prompt
│       │   ├── element_extractor.py      # python-pptx 元素提取
│       │   ├── asset_extractor.py        # 素材分类/保存
│       │   ├── theme.py                  # 主题色/字体提取
│       │   └── spec_model.py             # 数据类型定义
│       ├── ppt_generator.py              # 多线程生成器
│       ├── slide_generator.py            # 单页 Agent Loop
│       ├── style_evaluator.py            # 5 维风格评估
│       ├── layout/designer.py            # 布局设计 Agent
│       ├── config.py                     # 配置加载
│       └── pipeline.py                   # SVG→PPTX 转换
├── specs/                                # 提取的 spec（按名称目录）
│   └── <name>/
│       ├── spec.json                     # 主配置
│       ├── cover.json                    # 封面布局
│       ├── content_*.json                # 内容布局
│       ├── end_page.json                 # 尾页布局
│       ├── slides/                       # 截图
│       └── assets/                       # 可复用素材
├── outlines/                             # 保存的内容大纲
├── assets/
│   ├── Illustration/                     # 400+ 矢量插图
│   └── bootstrap-icons-1.13.1/           # 2000+ SVG 图标
├── references/
│   ├── spec-format.md                    # Spec 格式文档
│   ├── prompt-ppt-content.md             # 内容生成原则（WPS）
│   └── prompt-ppt-layout.md              # 布局设计原则
├── config.example.txt                    # 配置模板
├── config.txt                            # 实际配置（gitignore）
└── requirements.txt                      # Python 依赖
```

## 依赖清单

| 依赖 | 用途 | 安装方式 |
|------|------|---------|
| Python 3.10+ | 运行环境 | — |
| LibreOffice | PPTX→PDF 截图 | `brew install --cask libreoffice` |
| poppler (pdftoppm) | PDF→PNG | `brew install poppler` |
| python-pptx | PPTX 读写 | `pip install python-pptx` |
| Pillow | 图片处理 | `pip install Pillow` |
| pdf2image | PDF 分页转 PNG | `pip install pdf2image` |
| PyYAML | YAML 解析 | `pip install PyYAML` |
| lxml | XML 解析 | `pip install lxml` |
| openai (可选) | VL 模型调用 | `pip install openai` |

## License

MIT
