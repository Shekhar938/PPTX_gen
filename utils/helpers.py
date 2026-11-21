"""
helpers.py - Utility functions for text processing, asset management, and environment setup.

This module provides common helpers used across the PPTX generator:
- sanitize_text: Cleans text of excessive whitespace and unprintable control characters.
- truncate_text: Smart truncation at sentence boundaries with word-level fallback.
- extract_key_numbers: Extracts numerical metrics (currencies, percentages, multipliers, quantities).
- get_project_root: Resolves the absolute path to the project root directory.
- get_asset_path: Resolves paths to assets (icons, characters, images).
"""

from __future__ import annotations

import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional


def get_project_root() -> Path:
    """Return the absolute path to the project root directory.

    Resolves based on the location of this file or the PROJECT_ROOT environment
    variable if explicitly defined.

    Returns:
        Path: Absolute path to the PPTX_gen project root.
    """
    env_root = os.environ.get("PROJECT_ROOT")
    if env_root:
        return Path(env_root).resolve()

    # helpers.py is in <root>/utils/helpers.py -> parent.parent is <root>
    candidate = Path(__file__).resolve().parent.parent
    return candidate


def get_asset_path(category: str, name: str) -> Path:
    """Return the resolved path to an asset file.

    Looks inside ``<project_root>/assets/<category>/``.
    If ``name`` does not contain a file extension, checks for existing files
    matching common extensions (.svg, .png, .jpg, .jpeg, .webp).
    If no file exists on disk, defaults to ``.svg`` extension.

    Args:
        category: Subfolder under assets (e.g. 'icons', 'characters').
        name: Name of the asset (e.g. 'brain', 'brain.svg', 'thinking').

    Returns:
        Path: Path object pointing to the asset file.

    Examples:
        >>> get_asset_path('icons', 'brain')
        PosixPath('/home/shekhar/Desktop/PPTX_gen/assets/icons/brain.svg')
        >>> get_asset_path('characters', 'thinking.png')
        PosixPath('/home/shekhar/Desktop/PPTX_gen/assets/characters/thinking.png')
    """
    category_dir = get_project_root() / "assets" / category

    # If extension is already provided in the name, return directly
    has_extension = bool(re.search(r"\.[a-zA-Z0-9]+$", name))
    if has_extension:
        return category_dir / name

    # Search for an existing file with standard asset extensions
    for ext in (".svg", ".png", ".jpg", ".jpeg", ".webp"):
        candidate = category_dir / f"{name}{ext}"
        if candidate.is_file():
            return candidate

    # Default to .svg if no existing file was found
    return category_dir / f"{name}.svg"


def sanitize_text(text: str) -> str:
    """Clean text of excessive whitespace, control characters, and encoding artifacts.

    Normalizes newlines, removes non-printable control characters, strips
    zero-width characters and non-breaking spaces, collapses multiple
    spaces to single spaces, and limits consecutive newlines to two.

    Args:
        text: Input string to sanitize.

    Returns:
        str: Cleaned, normalized string. Returns empty string if input is empty or None.
    """
    if not text:
        return ""

    if not isinstance(text, str):
        text = str(text)

    # Normalize special unicode spaces and zero-width artifacts
    text = (
        text.replace("\u00a0", " ")   # non-breaking space
        .replace("\u200b", "")        # zero-width space
        .replace("\ufeff", "")        # byte order mark (BOM)
        .replace("\u200e", "")        # left-to-right mark
        .replace("\u200f", "")        # right-to-left mark
    )

    # Normalize carriage returns and line endings
    text = text.replace("\r\n", "\n").replace("\r", "\n")

    # Strip unprintable control characters while preserving \n and \t
    # ASCII 0-8, 11-12, 14-31, and 127-159 (DEL & C1 controls)
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]", "", text)

    # Normalize horizontal whitespace per line (tabs and spaces to single space)
    lines: List[str] = []
    for line in text.split("\n"):
        cleaned_line = re.sub(r"[ \t]+", " ", line).strip()
        lines.append(cleaned_line)

    # Rejoin lines and collapse 3+ newlines to max 2 (\n\n for paragraph separation)
    normalized = "\n".join(lines)
    normalized = re.sub(r"\n{3,}", "\n\n", normalized)

    return normalized.strip()


def truncate_text(text: str, max_chars: int = 500) -> str:
    """Smartly truncate text at a sentence boundary before max_chars.

    Attempts to find the last clean sentence boundary (ending in '.', '!', or '?')
    within the character limit. If no sentence boundary exists or if truncating
    there would discard too much content (less than 30% of max_chars), falls
    back to the last word boundary and appends an ellipsis.

    Args:
        text: The string to truncate.
        max_chars: Maximum character length allowed. Must be a positive integer.

    Returns:
        str: Truncated text within max_chars length.
    """
    if not text:
        return ""

    text = sanitize_text(text)
    if len(text) <= max_chars:
        return text

    if max_chars <= 3:
        return text[:max_chars]

    # Candidate slice up to max_chars
    candidate = text[:max_chars]

    # Look for sentence boundaries ending with '.', '!', or '?' followed by whitespace or quote
    sentence_matches = list(re.finditer(r"([.!?])(?:\s+|\Z)", candidate))
    if sentence_matches:
        # Check from the latest match backward
        for match in reversed(sentence_matches):
            end_pos = match.end()
            # Accept if it preserves a reasonable amount of context (>= 30% of max_chars)
            if end_pos >= max_chars * 0.3:
                return candidate[:end_pos].strip()

    # Fallback: truncate at the last word boundary within (max_chars - 3) to accommodate "..."
    cut_limit = max_chars - 3
    slice_candidate = candidate[:cut_limit]
    last_space = slice_candidate.rfind(" ")

    if last_space > 0:
        return slice_candidate[:last_space].rstrip(".,;:!? \t") + "..."

    # If no space was found (e.g. single long token), hard cut
    return candidate[:cut_limit].rstrip(" \t") + "..."


def extract_key_numbers(text: str) -> List[Dict[str, str]]:
    """Extract statistics and numeric metrics from text using regex patterns.

    Identifies key metrics commonly highlighted in presentations:
    - Currency amounts: e.g. '$45.2M', '€120,000', '£1.5B', '500 USD'
    - Percentages: e.g. '35%', '+18.5%', '-4.2%', '25 percent'
    - Multipliers: e.g. '10x', '2.5X', '5-fold'
    - Quantities: e.g. '5 million', '2.5 billion', '100k', '50,000 users'

    Args:
        text: Input text containing data, metrics, or statistics.

    Returns:
        List[Dict[str, str]]: List of extracted metrics with keys:
            - 'value': The exact matched metric string.
            - 'type': 'currency' | 'percentage' | 'multiplier' | 'quantity'.
            - 'context': The sentence containing the metric.
    """
    if not text:
        return []

    # Regex definitions ordered by specificity
    metric_patterns: List[tuple[str, re.Pattern[str]]] = [
        (
            "currency",
            re.compile(
                r"([\$€£¥₹]\s*\d+(?:,\d{3})*(?:\.\d+)?\s*(?:[kKmMbBtT]|thousand|million|billion|trillion)?|"
                r"\b\d+(?:,\d{3})*(?:\.\d+)?\s*(?:USD|EUR|GBP|INR)\b)",
                re.IGNORECASE,
            ),
        ),
        (
            "percentage",
            re.compile(
                r"([+-]?\b\d+(?:\.\d+)?\s*(?:%|percent(?:age)?\b))",
                re.IGNORECASE,
            ),
        ),
        (
            "multiplier",
            re.compile(
                r"(\b\d+(?:\.\d+)?\s*(?:[xX]|\-fold)\b)",
            ),
        ),
        (
            "quantity",
            re.compile(
                r"(\b\d+(?:,\d{3})*(?:\.\d+)?\s*(?:[kKmMbBtT]|thousand|million|billion|trillion)\b)",
                re.IGNORECASE,
            ),
        ),
    ]

    results: List[Dict[str, str]] = []

    # Split text into sentences or paragraphs for contextual grouping
    sentences = re.split(r"(?<=[.!?\n])\s+", text)

    for sentence in sentences:
        s_clean = sentence.strip()
        if not s_clean:
            continue

        # Track spans within the current sentence to avoid duplicate or overlapping matches
        matched_spans: List[tuple[int, int]] = []

        for m_type, regex in metric_patterns:
            for match in regex.finditer(s_clean):
                start, end = match.start(), match.end()

                # Check for overlap with already extracted spans in this sentence
                overlap = any(
                    not (end <= s_start or start >= s_end)
                    for s_start, s_end in matched_spans
                )
                if overlap:
                    continue

                matched_spans.append((start, end))
                matched_val = match.group(0).strip()

                results.append({
                    "value": matched_val,
                    "type": m_type,
                    "context": s_clean,
                })

    return results
