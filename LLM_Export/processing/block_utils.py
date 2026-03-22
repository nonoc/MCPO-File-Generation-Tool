"""
Block processing utilities for content structure handling.
"""

import re
import logging
from typing import Any, List, Dict, Tuple, Optional

log = logging.getLogger(__name__)

# Try relative import for constants, fall back to absolute
try:
    from ..shared.constants import TABLE_SEPARATOR_RE
except ImportError:
    from shared.constants import TABLE_SEPARATOR_RE

# Paragraph-like block types
_PARAGRAPH_TYPES = {"paragraph", "text", "body", "description", "summary", "title"}


def _expand_paragraph_block(block: Dict) -> List[Dict]:
    """
    Expand a paragraph block with segments into semantic blocks.
    
    Args:
        block: The paragraph block to expand
        
    Returns:
        List of semantic block dictionaries
    """
    segments = block.get("segments") or []
    child_blocks = block.get("children")
    if not segments:
        fallback = dict(block)
        fallback.pop("segments", None)
        return [fallback]

    semantic_blocks: List[Dict] = []
    bullet_buffer: List[Dict] = []

    def flush_bullets() -> None:
        nonlocal bullet_buffer
        if not bullet_buffer:
            return
        semantic_blocks.append({"type": "bullet_list", "items": [dict(item) for item in bullet_buffer]})
        bullet_buffer = []

    for segment in segments:
        seg_type = segment.get("type")
        if seg_type == "bullet":
            text = segment.get("text") or ""
            formatted = segment.get("formatted") or text
            if text:
                bullet_buffer.append({"text": text, "formatted": formatted})
            continue
        flush_bullets()
        if seg_type == "label":
            label_block: Dict = {
                "type": "label_paragraph",
                "label": segment.get("label"),
                "formatted_label": segment.get("formatted_label"),
            }
            rest_text = segment.get("text")
            if rest_text:
                label_block["text"] = rest_text
            rest_formatted = segment.get("formatted_text")
            if rest_formatted:
                label_block["formatted"] = rest_formatted
            semantic_blocks.append(label_block)
            continue
        paragraph_text = segment.get("text")
        if paragraph_text:
            paragraph_block: Dict = {"type": "paragraph", "text": paragraph_text}
            paragraph_formatted = segment.get("formatted")
            if paragraph_formatted:
                paragraph_block["formatted"] = paragraph_formatted
            semantic_blocks.append(paragraph_block)
    flush_bullets()

    if not semantic_blocks:
        fallback = dict(block)
        fallback.pop("segments", None)
        return [fallback]

    if child_blocks and semantic_blocks:
        semantic_blocks[-1]["children"] = child_blocks

    return semantic_blocks


def _finalize_normalized_block(block: Dict) -> List[Dict]:
    """
    Finalize a normalized block, handling special cases like tables.
    
    Args:
        block: The normalized block to finalize
        
    Returns:
        List of finalized block dictionaries
    """
    candidate = dict(block)
    block_type = (candidate.get("type") or "").strip().lower()
    if block_type == "table":  # protect tables unconditionally first
        candidate.pop("segments", None)
        return [candidate]
    if block_type in _PARAGRAPH_TYPES:
        return _expand_paragraph_block(candidate)
    candidate.pop("segments", None)
    return [candidate]


def _convert_markdown_to_structured(markdown_content: str) -> List[Dict]:
    """
    Convert Markdown content into a structured format.
    
    Args:
        markdown_content: Markdown content
        
    Returns:
        List of structured block dictionaries
    """
    if not markdown_content or not isinstance(markdown_content, str):
        return []
    
    lines = markdown_content.splitlines()
    structured = []
    i = 0
    
    # Use TABLE_SEPARATOR_RE from shared module
    from ..shared.constants import TABLE_SEPARATOR_RE

    def parse_table_row(row_line: str) -> List[str]:
        cleaned = row_line.strip().strip("|")
        return [cell.strip() for cell in cleaned.split("|")] if cleaned else []

    while i < len(lines):
        raw_line = lines[i]
        line = raw_line.strip()
        if not line:
            i += 1
            continue

        if line.startswith('|') and i + 1 < len(lines) and TABLE_SEPARATOR_RE.match(lines[i + 1].strip()):
            rows = [parse_table_row(line)]
            i += 2
            while i < len(lines):
                next_line = lines[i].strip()
                if not next_line or '|' not in next_line:
                    break
                rows.append(parse_table_row(next_line))
                i += 1
            structured.append({"type": "table", "data": rows})
            continue

        if line.startswith('# '):
            structured.append({"text": line[2:].strip(), "type": "title"})
        elif line.startswith('## '):
            structured.append({"text": line[3:].strip(), "type": "heading"})
        elif line.startswith('### '):
            structured.append({"text": line[4:].strip(), "type": "subheading"})
        elif line.startswith('#### '):
            structured.append({"text": line[5:].strip(), "type": "subheading"})
        elif line.startswith('- '):
            structured.append({"text": line[2:].strip(), "type": "bullet"})
        elif line.startswith('* '):
            structured.append({"text": line[2:].strip(), "type": "bullet"})
        elif line.startswith('**') and line.endswith('**'):
            structured.append({"text": line[2:-2].strip(), "type": "bold"})
        else:
            structured.append({"text": line, "type": "paragraph"})
        i += 1
    
    return structured


def flatten_structured_blocks(blocks: List[Any] | None, depth: int = 1) -> List[Tuple[Any, int]]:
    """
    Flatten structured blocks with depth information.
    
    Args:
        blocks: The structured blocks to flatten
        depth: Starting depth level
        
    Returns:
        List of (block, depth) tuples
    """
    if not blocks:
        return []
    flattened: List[Tuple[Any, int]] = []
    for block in blocks:
        flattened.append((block, depth))
        if isinstance(block, dict):
            child_depth = depth + 1 if (block.get("type") or "").strip().lower() == "section" else depth
            children = block.get("children")
            if isinstance(children, list) and children:
                flattened.extend(flatten_structured_blocks(children, child_depth))
    return flattened