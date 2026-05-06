"""Content gathering package — data models and logic for adaptive content gathering.

This package provides the dataclass schemas that form the contract between
Phase 3 (Content Gathering) and Phase 4 (PPT Generation), plus the internal
types for sufficiency assessment and adaptive questioning.

The data flow:
  1. Raw user input → Sufficiency assessment (is it enough?)
  2. Not sufficient → Adaptive questioning (max 8 questions)
  3. Enriched input → Content Outline generation
  4. ContentOutline → Phase 4 PPT Generation
"""

from ppt_skill.content.model import (
    ContentOutline,
    OutlineLayoutType,
    SlideEntry,
)

__all__ = [
    "ContentOutline",
    "OutlineLayoutType",
    "SlideEntry",
]
