# Phase 3: Content Gathering - Research

**Researched:** 2026-05-07
**Domain:** AI/LLM-driven interactive content gathering — structured questioning, sufficiency heuristics, content outline generation
**Confidence:** HIGH (core technique is prompt engineering with structured dataclass contracts; no external library risk; verified against existing Phase 2 codebase conventions)

## Summary

Phase 3 has **zero external library dependencies**. The "technology" is Claude's native reasoning coupled with structured Python data classes and YAML serialization — the same stack already in the project. The phase implements three behaviors: (1) **sufficiency detection** — a prompt-engineering heuristic that evaluates whether user input has enough structure, topic breadth, and content depth to skip questioning; (2) **smart hybrid questioning** — section-level overview questions first (to establish structure), then per-section gap-fill questions (to complete detail), capped at 8 total; and (3) **content outline generation** — producing a slide-by-slide outline (title, body content, layout type per slide) as a structured dataclass that Phase 4 consumes directly.

The critical design risk is **LLM consistency** — the same input can produce different sufficiency assessments or question quality across runs. Mitigations: use explicit scoring rubrics in sufficiency prompts, deterministic question framing templates, and a two-pass approach (assess → ask → assess again → generate). The data contract between Phase 3 and Phase 4 is the most important artifact — the content outline structure must carry enough information for Phase 4 to select templates, apply styles, and compose slides without ambiguity.

**Primary recommendation:** Build a `ContentGatherer` class with three modes: `auto` (full questioning pipeline), `skip_questions` (direct outline generation from sufficient input), and `outline_only` (generate outline from structured content without any interaction). Use the same dataclass + YAML pattern as Phase 2 for the content outline artifact, ensuring Phase 4 can consume it identically to how Phase 4 consumes DesignSpec.

## Standard Stack

### Core
| Library | Version | Purpose | Why Standard |
|---------|---------|---------|--------------|
| (stdlib only) | 3.10+ | dataclasses, enum, typing, pathlib, yaml | No new deps; Phase 2 already added PyYAML >=6.0 |
| PyYAML | >=6.0 | Content outline serialization to `outlines/<name>.yaml` | Already in requirements.txt from Phase 2; same serialization pattern as DesignSpec |
| python-pptx | 1.0.2 | N/A — Phase 3 does NOT generate PPTX; it produces a structured outline consumed by Phase 4 | Already in project |
| lxml | 6.0.2 | N/A — no XML parsing needed in Phase 3 | Already in project |

### Supporting
| Library | Version | Purpose | When to Use |
|---------|---------|---------|-------------|
| (none) | — | — | This phase is pure AI reasoning + structured data. No runtime libraries needed beyond stdlib + PyYAML. |

### Alternatives Considered
| Instead of | Could Use | Tradeoff |
|------------|-----------|----------|
| Dataclass outline model | Pydantic BaseModel | Pydantic adds dependency; project decision (Phase 2) is dataclasses for the Phase 2–4 contract |
| YAML outline storage | JSON | YAML is human-readable/git-diffable (same rationale as Phase 2 spec files); JSON loses comment/edit ergonomics |
| LLM-only (no data model) | Free-text outline | Free-text is ambiguous for Phase 4 consumption; structured dataclass is parseable and testable |

**Installation:**
```bash
# No new dependencies. PyYAML already in requirements.txt from Phase 2.
```

## Architecture Patterns

### Recommended Project Structure
```
src/ppt_skill/
├── content/                       # NEW: Content gathering module
│   ├── __init__.py
│   ├── model.py                   # ContentOutline, SlideEntry dataclasses
│   ├── sufficiency.py             # Input sufficiency scoring + skip-question logic
│   ├── questioning.py             # Question generation (section-level + gap-fill)
│   └── gatherer.py                # ContentGatherer orchestrator (entry point)
├── cli/
│   ├── __init__.py
│   ├── spec_commands.py           # EXISTING from Phase 2
│   └── content_commands.py        # NEW: gather-content, generate-outline commands
├── ...
outlines/                          # NEW: Project-local outline storage
└── <outline-name>.yaml
```

### Pattern 1: Two-Pass Sufficiency → Generation

**What:** First pass evaluates whether user input is sufficient to skip questioning. Second pass generates the outline (either directly or after questioning fills gaps).

**When to use:** Every content gathering session. This separates the decision logic from the generation logic, making both testable independently.

**Example:**
```python
# Source: Design pattern from existing Phase 2 extractor.py patterns
@dataclass
class SufficiencyResult:
    sufficient: bool
    confidence: float          # 0.0–1.0
    missing_dimensions: list[str]  # ["structure", "detail", "audience", ...]
    section_count: int
    estimated_slide_count: int

def assess_sufficiency(user_input: str) -> SufficiencyResult:
    """
    Evaluate whether user input has enough detail to skip questioning.
    
    Checks for:
    - Topic breadth: Are multiple distinct sections/topics present?
    - Content depth: Per-topic word count, bullet points, supporting details
    - Structural signals: Numbered sections, slide-like formatting, explicit outline
    - Audience/purpose: Is target audience and presentation purpose stated?
    
    Returns SufficiencyResult with confidence score.
    The caller decides: confidence >= 0.7 → skip questions; otherwise → ask.
    """
    ...
```

### Pattern 2: Question Buffer + Cap Enforcement

**What:** Maintain a running counter of questions asked. Section-level questions come first (one per section), then gap-fill questions fill remaining budget. Hard stop at 8.

**When to use:** Every interactive questioning session.

**Example:**
```python
@dataclass
class Question:
    id: int
    category: str          # "structure", "detail", "audience", "storytelling"
    text: str
    target_section: str | None  # None = section-level overview question
    context: str            # Why this question is being asked

@dataclass
class QuestionSession:
    questions_asked: list[Question] = field(default_factory=list)
    budget_remaining: int = 8
    sections_identified: list[str] = field(default_factory=list)
    gaps_per_section: dict[str, list[str]] = field(default_factory=dict)

    def can_ask(self) -> bool:
        return self.budget_remaining > 0

    def allocate_section_questions(self, sections: list[str]) -> list[Question]:
        """Phase 1: Ask one overview question per section (consumes budget)."""
        ...

    def allocate_gap_questions(self) -> list[Question]:
        """Phase 2: Ask targeted questions for sections with identified gaps."""
        ...
```

### Pattern 3: Dataclass-First Content Outline (Phase 4 Contract)

**What:** Define the content outline schema as dataclasses BEFORE writing any gathering logic. The schema drives what gets collected and what Phase 4 can consume.

**When to use:** Always. This is the contract between Phase 3 (content gathering) and Phase 4 (PPT generation).

**Example:**
```python
from dataclasses import dataclass, field
from enum import Enum
from datetime import datetime

class OutlineLayoutType(str, Enum):
    """Layout types the content outline can recommend per slide.
    
    Mapped to SlideType enum from spec_model.py:
      TITLE         → SlideType.TITLE
      CONTENT       → SlideType.CONTENT  
      TWO_COLUMN    → SlideType.CONTENT (or IMAGE_TEXT depending on content)
      SECTION_DIVIDER → SlideType.SECTION_DIVIDER
      IMAGE_TEXT    → SlideType.IMAGE_TEXT
      DATA          → SlideType.DATA
    """
    TITLE = "title"
    CONTENT = "content"
    TWO_COLUMN = "two_column"
    SECTION_DIVIDER = "section_divider"
    IMAGE_TEXT = "image_text"
    DATA = "data"

@dataclass
class SlideEntry:
    """One slide in the content outline."""
    slide_number: int
    title: str
    body: list[str]                   # Bullet points or paragraphs
    layout_type: OutlineLayoutType
    notes: str = ""                   # Speaker notes (optional)
    image_hint: str = ""              # Suggested image topic/keyword (optional)
    section_name: str = ""            # Which section this slide belongs to

@dataclass  
class ContentOutline:
    """Full slide-by-slide content outline — output of Phase 3, input to Phase 4."""
    metadata: dict[str, object]       # name, created_at, input_summary, question_count
    presentation_title: str
    presentation_subtitle: str = ""
    target_audience: str = ""
    presentation_purpose: str = ""
    sections: list[str] = field(default_factory=list)  # Ordered section names
    slides: list[SlideEntry] = field(default_factory=list)
    spec_name: str = ""               # Active spec to use for generation (from Phase 2)
```

### Anti-Patterns to Avoid

- **Free-text outline output:** Phase 4 needs structured slide data (title, body, layout type). A plain text or markdown outline forces Phase 4 to re-parse and re-interpret content, introducing ambiguity and errors. Always output structured dataclass + YAML.
- **Single-pass questioning:** Asking all 8 questions at once in a batch defeats the purpose of adaptive questioning — later questions should depend on earlier answers. Section-level questions first, then gap-fill.
- **Questioning without context:** Asking "What should slide 3 contain?" without providing the section structure the user has already described. Every question should reference what's already known.
- **Hardcoded question count without budget tracking:** A `while budget > 0` loop without tracking which gaps are being filled leads to redundant questions. Track gap status per section.
- **Mixing questioning and generation in one function:** These are separate concerns. Sufficiency assessment → questioning (if needed) → outline generation. Each step should be independently callable (especially "skip questions and generate" for GEN-03).

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| YAML serialization | Custom YAML writer | PyYAML `yaml.dump()` (already in project) | Handles multi-line strings, special characters, anchors; same pattern as SpecExtractor.save() |
| Content outline data model | Free-form dicts or plain text | Dataclasses with typed fields | Phase 4 consumes structured data; dataclasses provide field validation, IDE autocomplete, and testability |
| CLI argument parsing | Manual argv parsing | Phase 5 will wire argparse; Phase 3 CLIs can use simple `sys.argv` or direct function calls like Phase 2's spec_commands.py | Phase 2 CLI functions are already callable from Python without argparse; Phase 3 follows same pattern |
| Question prioritization | Random question selection | Explicit category system (structure > detail > audience > storytelling) | Ensures consistent question quality; prevents asking about fine details before establishing structure |
| Slide content suggestion | Content generation without spec context | Reference loaded spec's SlideType distribution and rhythm as guidance | The spec's slide type distribution informs how many title/content/section slides to suggest |

**Key insight:** The core implementation is ~200 lines of Python dataclasses and orchestrator logic, plus prompt templates for Claude's questioning behavior. The real design work is in the **sufficiency rubric**, **question category system**, and **content outline schema** — these determine output quality more than any code complexity.

## Common Pitfalls

### Pitfall 1: Inconsistent Sufficiency Assessment
**What goes wrong:** The same user input gets classified as "sufficient" in one run and "insufficient" in another, causing non-deterministic behavior.
**Why it happens:** LLM-based assessment without explicit scoring criteria. Vague prompts produce varying interpretations.
**How to avoid:** Use a concrete scoring rubric with explicit thresholds:
  - Structure score: Has numbered sections? (+1), Has slide-like formatting? (+1)
  - Detail score: Average words per topic > 30? (+1), Has bullet points? (+1), Has supporting examples/data? (+1)
  - Audience score: Target audience stated? (+1), Purpose stated? (+1)
  - Total >= 5 → sufficient (confidence >= 0.7)
**Warning signs:** Same input triggers different questioning paths; users report "sometimes it asks, sometimes it doesn't."

### Pitfall 2: Question Fatigue from Redundant Questions
**What goes wrong:** The tool asks about information already present in the user's input, wasting the 8-question budget and frustrating the user.
**Why it happens:** Question generation doesn't check against already-gathered content. Gap detection is naive or missing.
**How to avoid:** Before generating each question, diff against the current knowledge state. Tag each piece of gathered information with its source (user input, section Q answer, gap-fill Q answer). Only ask about sections/topics with explicit gaps.
**Warning signs:** User says "I already told you about that" or provides one-word answers to redundant questions.

### Pitfall 3: Overly Granular Questioning
**What goes wrong:** After establishing 4 sections, the tool asks detailed questions about slide 3's third bullet point instead of checking whether section 2 is adequately covered.
**Why it happens:** The questioning strategy doesn't prioritize breadth-first coverage. It goes deep on the first section while others remain empty.
**How to avoid:** Phase 1 questions ask ONE question per identified section (overview level). Only after all sections have baseline coverage do Phase 2 gap-fill questions target specific sections. The algorithm: sort sections by gap severity, ask one question per section in round-robin order until budget exhausted.
**Warning signs:** Detailed questions about slide content before the high-level story arc is confirmed.

### Pitfall 4: Outline vs. Phase 4 Template Mismatch
**What goes wrong:** The outline suggests layout types that Phase 4 doesn't know how to render (e.g., "timeline", "comparison_table") or can't map to SlideType values.
**Why it happens:** The outline model allows arbitrary layout strings without validation against available template types.
**How to avoid:** Constrain `OutlineLayoutType` enum to values that map to `SlideType` + known Phase 4 rendering capabilities. For v1: title, content, two_column, section_divider, image_text, data. Future chart types can be added when Phase 4 supports them.
**Warning signs:** Phase 4 errors on unrecognized `layout_type` values; outline suggests layouts that don't exist.

### Pitfall 5: Empty or Vague Outline Body Content
**What goes wrong:** Generated outlines have slide titles but body content is "TBD", "Content here", or single-word bullets — insufficient for Phase 4 to generate meaningful slides.
**Why it happens:** The outline generation prompt doesn't enforce minimum content requirements. LLM takes shortcuts.
**How to avoid:** Validate outline post-generation: every slide must have a non-empty title, at least one body bullet with >20 characters, and a valid layout_type. Reject and regenerate outlines that fail validation.
**Warning signs:** Slides with generic placeholder text; Phase 4 generates slides with "Add content here" text boxes.

## Code Examples

Verified patterns from existing codebase conventions:

### Content Outline Data Model
```python
# Source: Following same pattern as spec_model.py (Phase 2)
from dataclasses import dataclass, field
from enum import Enum
from datetime import datetime

class OutlineLayoutType(str, Enum):
    TITLE = "title"
    CONTENT = "content"
    TWO_COLUMN = "two_column"
    SECTION_DIVIDER = "section_divider"
    IMAGE_TEXT = "image_text"
    DATA = "data"

@dataclass
class SlideEntry:
    slide_number: int
    title: str
    body: list[str] = field(default_factory=list)
    layout_type: OutlineLayoutType = OutlineLayoutType.CONTENT
    notes: str = ""
    image_hint: str = ""
    section_name: str = ""

@dataclass
class ContentOutline:
    metadata: dict = field(default_factory=dict)
    presentation_title: str = ""
    presentation_subtitle: str = ""
    target_audience: str = ""
    presentation_purpose: str = ""
    sections: list[str] = field(default_factory=list)
    slides: list[SlideEntry] = field(default_factory=list)
    spec_name: str = ""

    def to_dict(self) -> dict:
        """Serialize to plain dict for YAML output, same pattern as Phase 2."""
        return _dataclass_to_dict(self)

    @classmethod
    def from_dict(cls, data: dict) -> "ContentOutline":
        """Deserialize from YAML-loaded dict."""
        slides_data = data.pop("slides", [])
        outline = cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})
        for sd in slides_data:
            sd["layout_type"] = OutlineLayoutType(sd.get("layout_type", "content"))
            outline.slides.append(SlideEntry(**{k: v for k, v in sd.items() if k in SlideEntry.__dataclass_fields__}))
        return outline
```

### Sufficiency Assessment (Prompt Template Pattern)
```python
# Source: Design — rubric-based heuristics for consistent LLM decisions
# This is a PROMPT TEMPLATE, not runtime code. Claude calls this internally.
SUFFICIENCY_PROMPT = """
Evaluate the following content input for a presentation. Score each dimension 0-2:

1. STRUCTURE: Are topics/sections clearly delineated?
   0=no structure, 1=some grouping, 2=numbered/headered sections

2. DETAIL: Does each section have supporting content?
   0=topics only, 1=some bullets/details, 2=rich content with examples/data

3. AUDIENCE: Is target audience and purpose clear?
   0=not stated, 1=implied, 2=explicitly stated

4. SCOPE: Is the presentation scope bounded?
   0=open-ended/unbounded, 1=rough scope, 2=well-defined scope

Input: {user_input}

Return JSON: {"scores": {"structure": N, "detail": N, "audience": N, "scope": N}, "total": N, "sufficient": bool, "rationale": "..."}

Threshold: total >= 5 AND structure >= 1 AND detail >= 1 → sufficient
"""

def assess_sufficiency(user_input: str) -> dict:
    """Run sufficiency assessment. Returns dict with 'sufficient' flag + dimensions."""
    # In practice, Claude processes this internally via the prompt template.
    # The implementation wraps this in a function returning SufficiencyResult.
    ...
```

### Question Generation (Section-First Strategy)
```python
# Source: Design — section-level overview first, gap-fill second
QUESTION_CATEGORY_PRIORITY = [
    "structure",     # What are the main sections?
    "detail",        # What content goes in each section?
    "audience",      # Who is this for, what's the takeaway?
    "storytelling",  # What's the narrative arc?
]

def generate_section_questions(sections: list[str]) -> list[dict]:
    """Phase 1: One overview question per identified section."""
    templates = {
        "structure": "What key points should the '{section}' section cover?",
        "detail": "What supporting data, examples, or evidence should appear in '{section}'?",
        "audience": "What should the audience understand or do after the '{section}' section?",
        "storytelling": "How should '{section}' transition from the previous section?",
    }
    questions = []
    for i, section in enumerate(sections):
        # First section → structure, subsequent → detail
        category = "structure" if i == 0 else "detail"
        q = templates[category].format(section=section)
        questions.append({"category": category, "text": q, "target_section": section})
    return questions

def generate_gap_questions(
    gaps: dict[str, list[str]], 
    budget: int
) -> list[dict]:
    """Phase 2: Targeted questions for sections with identified gaps.
    
    gaps: {"section_name": ["missing_detail", "no_examples", ...], ...}
    budget: remaining question budget (1–8)
    
    Returns up to `budget` questions, prioritizing sections with most gaps.
    """
    gap_templates = {
        "missing_detail": "Can you provide 2-3 specific points for '{section}'?",
        "no_examples": "What concrete example or case study illustrates '{section}'?",
        "no_takeaway": "What's the single key takeaway for the '{section}' section?",
        "missing_transition": "How does '{section}' connect to the next topic?",
    }
    sorted_sections = sorted(gaps.items(), key=lambda x: len(x[1]), reverse=True)
    questions = []
    for section, gap_types in sorted_sections:
        if len(questions) >= budget:
            break
        gap_type = gap_types[0]
        template = gap_templates.get(gap_type, gap_templates["missing_detail"])
        questions.append({
            "category": "detail",
            "text": template.format(section=section),
            "target_section": section,
        })
    return questions
```

### Content Outline Generation (Structure Enforcement)
```python
# Source: Design — post-generation validation against minimum content requirements
OUTLINE_GENERATION_PROMPT = """
Generate a slide-by-slide content outline for the following presentation.

Context:
- Active spec: {spec_name}
- Available layout types: {layout_types}
- Target audience: {audience}
- Purpose: {purpose}

Content gathered:
{content_summary}

Rules:
1. First slide must be a TITLE slide with presentation title and subtitle
2. Each section starts with a SECTION_DIVIDER slide (section name only)
3. Content slides use CONTENT layout type
4. Use TWO_COLUMN for comparison/contrast content
5. Use IMAGE_TEXT for slides that would benefit from imagery
6. Every slide must have a descriptive title AND at least 2 body points
7. Total slides: follow the gathered content scope (don't pad or truncate)

Output JSON array of slides: [{"title": "...", "body": ["...", "..."], "layout_type": "...", ...}]
"""

def validate_outline(outline: ContentOutline) -> list[str]:
    """Post-generation validation. Returns list of issues (empty = valid)."""
    issues = []
    if not outline.presentation_title:
        issues.append("Missing presentation title")
    if not outline.slides:
        issues.append("No slides generated")
    for slide in outline.slides:
        if not slide.title.strip():
            issues.append(f"Slide {slide.slide_number}: empty title")
        if not slide.body or all(len(b.strip()) < 10 for b in slide.body):
            issues.append(f"Slide {slide.slide_number}: body too short or empty")
        if slide.layout_type not in OutlineLayoutType:
            issues.append(f"Slide {slide.slide_number}: invalid layout_type '{slide.layout_type}'")
    return issues
```

### YAML Serialization (Reusing Phase 2 Pattern)
```python
# Source: Follows spec_commands.py / extractor.py pattern exactly
import yaml
from pathlib import Path
from ppt_skill.spec.extractor import _dataclass_to_dict  # Reuse Phase 2 serialization

def save_outline(outline: ContentOutline, outlines_dir: str = "outlines") -> Path:
    """Save content outline as YAML file in outlines/ directory."""
    out_dir = Path(outlines_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    
    outline.metadata.setdefault("saved_at", datetime.now().isoformat())
    outline_data = _dataclass_to_dict(outline)
    out_path = out_dir / f"{outline.metadata.get('name', 'outline')}.yaml"
    
    with open(out_path, "w", encoding="utf-8") as f:
        yaml.dump(outline_data, f, default_flow_style=False, sort_keys=False, allow_unicode=True)
    
    return out_path

def load_outline(name: str, outlines_dir: str = "outlines") -> ContentOutline:
    """Load content outline from YAML file."""
    filepath = Path(outlines_dir) / f"{name}.yaml"
    with open(filepath, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)
    return ContentOutline.from_dict(data)
```

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|--------------|--------|
| Single-pass "give me all content" prompt | Two-pass: sufficiency check → questioning (if needed) → outline | Design choice for Phase 3 | Avoids generating outlines from insufficient input; prevents user frustration from 1-shot failures |
| Fixed-question questionnaire (8 predefined questions) | Adaptive questioning with gap detection | Design choice for Phase 3 | Questions are relevant to actual gaps, not generic; prevents redundant questioning |
| Free-text/markdown outline output | Structured dataclass → YAML outline | Design choice (following Phase 2 pattern) | Phase 4 can consume programmatically; outlines are machine-validatable |
| Layout type as free-text string | Constrained `OutlineLayoutType` enum | Design choice for Phase 3 | Prevents Phase 4 from receiving unknown layout types; validation at outline generation time |

**Deprecated/outdated:**
- N/A — Phase 3 is the first implementation of content gathering. No prior version exists.

## Open Questions

1. **Should the outline include suggested chart types when content is data-heavy?**
   - What we know: Phase 4 has 70+ chart SVG templates available (bar_chart, pie_chart, line_chart, etc.). The content outline could hint at chart types without requiring full data specification.
   - What's unclear: Whether Phase 4's generation pipeline accepts chart type hints or requires complete data specifications for chart generation.
   - Recommendation: Add an optional `chart_hint` field to `SlideEntry` (e.g., `chart_hint: "bar_chart"`) but make it advisory only. Phase 4 can ignore hints it can't fulfill. Don't block Phase 3 on chart integration — Phase 4 may handle chart generation in a separate plan.

2. **How should the outline reference the active spec?**
   - What we know: Phase 2's `spec_commands.py` implements `get_active_spec()` which reads `specs/.active`. Phase 3 can call this directly.
   - What's unclear: Whether Phase 4 reads the outline's `spec_name` field to auto-load the spec, or whether Phase 4 independently calls `get_active_spec()`. Redundancy could cause drift if outline was saved with spec A but spec B is now active.
   - Recommendation: Store `spec_name` in the outline metadata as the "frozen" spec reference at creation time. Phase 4 should use the outline's `spec_name` field (not current active spec) when generating from a saved outline. This prevents spec drift between outline creation and generation.

3. **What constitutes "sufficiently detailed" input for GEN-03 bypass?**
   - What we know: The rubric in the sufficiency prompt (structure + detail + audience + scope >= 5) is the initial design. Real-world inputs vary widely.
   - What's unclear: Whether the 5-point threshold is too high or too low. A user providing "Make a 10-slide deck about our Q3 results" has scope but no detail — should it trigger questioning?
   - Recommendation: Use the rubric as the default. Make the threshold configurable (but not exposed to users in v1). Collect data during testing to calibrate. The rubric is implemented as a prompt template, so it can be tuned without code changes.

4. **Should questioning support multi-turn conversation or single-pass Q&A?**
   - What we know: The success criteria say "section-level overview questions first, then targeted follow-ups." This implies sequential interaction (ask → answer → ask → answer).
   - What's unclear: Whether the implementation should batch questions (ask 2-3 at a time for efficiency) or ask one at a time (maximally adaptive but slower). YOLO mode means no user checkpoints mid-flow.
   - Recommendation: Batch 2-3 related questions per turn to balance adaptiveness with speed. Section-level questions can be batched (e.g., ask about 3 sections at once). Gap-fill questions should be 1-2 at a time since answers may affect subsequent gaps.

5. **What's the error behavior when the user provides unusable input?**
   - What we know: GEN-01 says the tool "asks section-level overview questions first when input lacks detail." This implies the tool should always make forward progress.
   - What's unclear: What happens when the user provides completely irrelevant input ("tell me a joke") or refuses to answer questions ("skip", "idk", empty responses)?
   - Recommendation: Implement a minimum-content gate — if after 8 questions the gathered content is still insufficient for a meaningful outline, generate a minimal outline (title slide + "please provide more detail" slide) and flag it. Don't crash or loop. The outline metadata can include a `quality: minimal` flag.

## Sources

### Primary (HIGH confidence)
- Existing codebase (`src/ppt_skill/spec/spec_model.py`, `src/ppt_skill/spec/extractor.py`, `src/ppt_skill/cli/spec_commands.py`) — verified dataclass patterns, YAML serialization, CLI function signatures, and project conventions
- `.planning/REQUIREMENTS.md` — confirmed Phase 3 requirements (GEN-01, GEN-02, GEN-03) and success criteria (5 items)
- `.planning/ROADMAP.md` — confirmed Phase 3 scope: "smart hybrid questioning and slide-by-slide outline generation", template layout types: title, content, two-column, section header, image+text
- `.planning/STATE.md` — confirmed Phase 2 completion, project conventions (dataclass NOT Pydantic, YAML artifacts, stateless CLI functions, YOLO mode)

### Secondary (MEDIUM confidence)
- Phase 2 RESEARCH.md and PLAN.md files — verified research and planning patterns for consistency in Phase 3 documentation
- SlideType enum in `spec_model.py` — confirmed 5 types: TITLE, CONTENT, SECTION_DIVIDER, IMAGE_TEXT, DATA; extra "two_column" type needed for content outlines

### Tertiary (LOW confidence)
- N/A — this phase has no external library dependencies; all patterns are derived from existing project conventions and requirements

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH — zero new dependencies; PyYAML already in project; all code is stdlib + dataclasses + existing imports
- Architecture: HIGH — patterns derived from existing Phase 2 codebase (SpecExtractor orchestrator, spec_commands CLI pattern, dataclass → YAML serialization). Three-part structure (sufficiency → questioning → generation) maps cleanly to requirements.
- Pitfalls: HIGH — LLM consistency pitfalls are well-understood (rubric-based assessment, gap-tracking, validation); Phase 2 pitfalls (pptx bugs, lxml complexity) are not relevant to Phase 3.
- Outline schema: MEDIUM — schema design is sound but exact Phase 4 contract needs may require adjustment during Phase 4 planning. The `OutlineLayoutType` enum and `SlideEntry` fields are a reasonable v1 contract.

**Research date:** 2026-05-07
**Valid until:** 2026-06-07 (no external library dependencies; schema stable until Phase 4 implementation reveals contract gaps)

<phase_requirements>
## Phase Requirements

| ID | Description | Research Support |
|----|-------------|-----------------|
| GEN-01 | Smart hybrid questioning — section-level overview questions first, then gap-fills per section with targeted follow-ups, capping at 8 total questions | §Pattern 2 (Question Buffer + Cap Enforcement) provides the budget-tracking algorithm. §Question Generation code shows section-first + gap-fill templates with round-robin allocation. §Pitfall 2 (Question Fatigue) and §Pitfall 3 (Overly Granular) address common failure modes. |
| GEN-02 | Generate a detailed content outline (title, body content, suggested layout type per slide) for user review and approval before any PPT generation begins | §Pattern 3 (Dataclass-First Content Outline) defines the `ContentOutline` + `SlideEntry` schema. §Outline Generation code provides the structured prompt template with layout type constraints. §Pitfall 5 (Empty Body Content) covers outline validation against minimum content requirements. YAML serialization (§Code Examples) follows Phase 2 `extractor.py` pattern for Phase 4 compatibility. |
| GEN-03 | User can skip questioning entirely when input is sufficiently detailed, going directly to content outline generation | §Pattern 1 (Two-Pass Sufficiency → Generation) defines the assessment rubric with explicit scoring (structure + detail + audience + scope). §Open Question 3 addresses threshold calibration. The `content_commands.py` CLI pattern supports `gather-content` (full pipeline) and `gather-content --skip-questions` (bypass questioning). |
</phase_requirements>
