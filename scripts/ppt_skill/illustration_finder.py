#!/usr/bin/env python3
"""Illustration finder for PPT skill.

Search and retrieve illustrations from assets/Illustration/ directory.

Usage:
    python3 scripts/ppt_skill/illustration_finder.py search "ai chat"
    python3 scripts/ppt_skill/illustration_finder.py list --category ai
    python3 scripts/ppt_skill/illustration_finder.py random --count 5
"""

from __future__ import annotations

import argparse
import json
import os
import random
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional

SCRIPT_DIR = Path(__file__).parent
ASSETS_DIR = SCRIPT_DIR.parent.parent / "assets"
ILLUSTRATION_DIR = ASSETS_DIR / "Illustration"

# Category mappings for search
CATEGORIES = {
    "ai": ["ai-", "artificial", "machine", "neural", "robot"],
    "data": ["data", "analytics", "chart", "graph", "database", "stats"],
    "team": ["team", "collaboration", "meeting", "together", "followers", "group"],
    "success": ["success", "accomplish", "complete", "goal", "achieve", "finished"],
    "business": ["business", "market", "customer", "contract", "agreement", "finance"],
    "learning": ["learn", "education", "exam", "study", "course", "teaching"],
    "document": ["document", "file", "paper", "report", "doc"],
    "security": ["security", "auth", "protect", "lock", "secure"],
    "mobile": ["mobile", "phone", "device", "sync", "app"],
    "creative": ["design", "creative", "art", "color", "palette", "create"],
}

# Keywords mapping for semantic search
KEYWORDS = {
    "chat": ["chat", "conversation", "message", "talk"],
    "robot": ["ai", "robot", "assistant", "bot"],
    "growth": ["growth", "up", "increase", "rise", "chart-up"],
    "coding": ["code", "developer", "programming", "dev", "cli"],
    "analysis": ["analyze", "analytics", "data", "stats", "metrics"],
    "team": ["team", "group", "people", "collaboration", "meeting"],
    "success": ["success", "complete", "done", "achieve", "finish"],
    "time": ["time", "clock", "schedule", "alarm", "calendar"],
    "document": ["document", "file", "paper", "report", "note"],
    "finance": ["money", "wallet", "crypto", "finance", "portfolio"],
    "celebration": ["celebrate", "party", "happy", "joy", "birthday"],
}


@dataclass
class Illustration:
    """Single illustration entry."""

    filename: str
    path: Path
    theme: str
    keywords: List[str] = field(default_factory=list)

    @property
    def name(self) -> str:
        """Human-readable name."""
        # undraw_ai-chat_ljb9.svg -> ai-chat
        name = self.filename.replace("undraw_", "").rsplit("_", 1)[0]
        return name.replace("-", " ").replace("_", " ")

    @property
    def svg_content(self) -> str:
        """Read SVG content."""
        return self.path.read_text(encoding="utf-8")

    def to_dict(self) -> dict:
        return {
            "filename": self.filename,
            "path": str(self.path),
            "name": self.name,
            "theme": self.theme,
            "keywords": self.keywords,
        }


class IllustrationFinder:
    """Find and retrieve illustrations."""

    def __init__(self, illustration_dir: Path | None = None):
        self.illustration_dir = illustration_dir or ILLUSTRATION_DIR
        self._illustrations: List[Illustration] | None = None

    def _load_illustrations(self) -> List[Illustration]:
        """Load all illustrations from directory."""
        if self._illustrations is not None:
            return self._illustrations

        illustrations: List[Illustration] = []
        for svg_file in sorted(self.illustration_dir.glob("*.svg")):
            filename = svg_file.name
            # Skip duplicates (files with " (1)" suffix)
            if " (" in filename:
                continue

            theme = self._extract_theme(filename)
            keywords = self._extract_keywords(filename)

            illustrations.append(Illustration(
                filename=filename,
                path=svg_file,
                theme=theme,
                keywords=keywords,
            ))

        self._illustrations = illustrations
        return illustrations

    def _extract_theme(self, filename: str) -> str:
        """Extract primary theme from filename."""
        # undraw_ai-chat_ljb9.svg -> ai-chat
        name = filename.replace("undraw_", "").rsplit("_", 1)[0]
        theme = name.split("-")[0]
        return theme

    def _extract_keywords(self, filename: str) -> List[str]:
        """Extract keywords from filename."""
        # undraw_ai-chat_ljb9.svg -> ["ai", "chat"]
        name = filename.replace("undraw_", "").rsplit("_", 1)[0]
        return name.split("-")

    def search(self, query: str, limit: int = 20) -> List[Illustration]:
        """Search illustrations by query."""
        query_lower = query.lower().strip()
        illustrations = self._load_illustrations()

        # Check if query matches a category
        for category, patterns in CATEGORIES.items():
            if category in query_lower:
                query_lower = "|".join(patterns)
                break

        # Check keyword mappings
        for keyword, synonyms in KEYWORDS.items():
            if keyword in query_lower:
                query_lower = f"{query_lower}|{'|'.join(synonyms)}"
                break

        results: List[Illustration] = []
        for ill in illustrations:
            # Search in name and keywords
            text = f"{ill.name} {' '.join(ill.keywords)}".lower()
            if any(q in text for q in query_lower.split("|")):
                results.append(ill)
                if len(results) >= limit:
                    break

        return results

    def list_by_category(self, category: str) -> List[Illustration]:
        """List illustrations by category."""
        illustrations = self._load_illustrations()

        if category not in CATEGORIES:
            return []

        patterns = CATEGORIES[category]
        results: List[Illustration] = []

        for ill in illustrations:
            # Check both filename and name
            filename_lower = ill.filename.lower()
            text = f"{ill.name} {' '.join(ill.keywords)}".lower()

            # Match pattern in filename (undraw_ai-chat_...) or text (ai chat)
            for p in patterns:
                if p in filename_lower or p in text:
                    results.append(ill)
                    break

        return results

    def random(self, count: int = 5) -> List[Illustration]:
        """Get random illustrations."""
        illustrations = self._load_illustrations()
        return random.sample(illustrations, min(count, len(illustrations)))

    def get_by_name(self, name: str) -> Optional[Illustration]:
        """Get illustration by name or filename."""
        illustrations = self._load_illustrations()

        # Try exact filename match first
        for ill in illustrations:
            if ill.filename == name or ill.filename.startswith(name):
                return ill

        # Try name match
        name_lower = name.lower().replace(" ", "-")
        for ill in illustrations:
            if ill.name.lower() == name_lower:
                return ill

        return None

    def list_categories(self) -> List[str]:
        """List available categories."""
        return list(CATEGORIES.keys())

    def count(self) -> int:
        """Total count of illustrations."""
        return len(self._load_illustrations())


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Illustration finder for PPT skill"
    )
    sub = parser.add_subparsers(dest="command", required=True)

    # search
    p_search = sub.add_parser("search", help="Search illustrations")
    p_search.add_argument("query", help="Search query")
    p_search.add_argument("--limit", "-l", type=int, default=20)

    # list
    p_list = sub.add_parser("list", help="List by category")
    p_list.add_argument("--category", "-c", help="Category name")
    p_list.add_argument("--categories", action="store_true", help="List categories")

    # random
    p_random = sub.add_parser("random", help="Get random illustrations")
    p_random.add_argument("--count", "-n", type=int, default=5)

    # get
    p_get = sub.add_parser("get", help="Get specific illustration")
    p_get.add_argument("name", help="Illustration name or filename")

    args = parser.parse_args()

    finder = IllustrationFinder()

    if args.command == "search":
        results = finder.search(args.query, args.limit)
        for ill in results:
            print(f"  {ill.filename:50} # {ill.name}")

    elif args.command == "list":
        if args.categories:
            print("Categories:")
            for cat in finder.list_categories():
                count = len(finder.list_by_category(cat))
                print(f"  {cat:15} ({count} illustrations)")
        elif args.category:
            results = finder.list_by_category(args.category)
            for ill in results:
                print(f"  {ill.filename:50} # {ill.name}")
        else:
            print(f"Total: {finder.count()} illustrations")
            print("Use --category to filter or --categories to list all")

    elif args.command == "random":
        results = finder.random(args.count)
        for ill in results:
            print(f"  {ill.filename:50} # {ill.name}")

    elif args.command == "get":
        ill = finder.get_by_name(args.name)
        if ill:
            print(f"Name: {ill.name}")
            print(f"File: {ill.filename}")
            print(f"Path: {ill.path}")
        else:
            print(f"Not found: {args.name}", file=sys.stderr)
            return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
