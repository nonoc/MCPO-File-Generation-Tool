"""
Processing utilities for content normalization and text handling.
"""

import re
import logging
from typing import Any, List, Dict, Tuple, Optional
try:
    from bs4 import BeautifulSoup
    from bs4.element import NavigableString
except ImportError:
    from bs4 import BeautifulSoup, NavigableString
import markdown2

log = logging.getLogger(__name__)

# Paragraph-like block types
_PARAGRAPH_TYPES = {"paragraph", "text", "body", "description", "summary", "title"}


def _normalize_markup_text(value: Any, inline_only: bool = False) -> Tuple[str, str]:
    """
    Normalize markup text by converting markdown to HTML and extracting plain/formatted text.
    
    Args:
        value: The text value to normalize
        inline_only: If True, strip block-level tags before processing
        
    Returns:
        Tuple of (plain_text, formatted_html)
    """
    if value is None:
        return "", ""
    text = str(value)
    html = markdown2.markdown(text, extras=["fenced-code-blocks"])
    soup = BeautifulSoup(html, "html.parser")
    if inline_only:
        # Strip block-level tags before the inline-only pass so that
        # list/paragraph HTML never reaches _apply_formatted_html_to_paragraph.
        # unwrap() keeps the tag's children in place, so text is preserved.
        for tag in soup.find_all(["p", "ul", "ol", "li", "blockquote", "pre", "div", "h1", "h2", "h3", "h4", "h5", "h6"]):
            tag.unwrap()
    allowed = {"strong", "b", "em", "i", "br"}
    for tag in soup.find_all():
        if tag.name not in allowed and tag.name != "body":
            tag.unwrap()
    body = soup.body
    formatted = ""
    if body:
        formatted = "".join(str(child) for child in body.children)
    else:
        formatted = str(soup)
    plain = soup.get_text(" ", strip=True)
    return plain, formatted


def _strip_wrapping_paragraph_tags(html: str) -> str:
    """Remove wrapping <p> tags from HTML content."""
    if not html:
        return ""
    trimmed = html.strip()
    stripped = re.sub(r"(?i)^<p>(.*?)</p>$", r"\1", trimmed, flags=re.S)
    return stripped.strip()


def _parse_paragraph_segments(raw_text: str) -> List[Dict]:
    """
    Parse raw text into structured segments (bullets, labels, paragraphs).
    
    Args:
        raw_text: Raw text to parse
        
    Returns:
        List of segment dictionaries
    """
    segments = []
    if not raw_text:
        return segments
    bullet_re = re.compile(r"^(?:[•\*-])\s+(.*)")
    label_re = re.compile(r"^\*\*(.+?):\*\*\s*(.*)")
    lines = [line.strip() for line in raw_text.splitlines() if line.strip()]
    for line in lines:
        bullet_match = bullet_re.match(line)
        if bullet_match:
            inner = bullet_match.group(1)
            plain, formatted = _normalize_markup_text(inner, inline_only=True)
            formatted = _strip_wrapping_paragraph_tags(formatted)
            if plain:
                segments.append({"type": "bullet", "text": plain, "formatted": formatted})
            continue
        label_match = label_re.match(line)
        if label_match:
            label_text = label_match.group(1).strip()
            label_value = f"{label_text}:"
            label_source = f"**{label_value}**"
            plain_label, formatted_label = _normalize_markup_text(label_source, inline_only=True)
            formatted_label = _strip_wrapping_paragraph_tags(formatted_label)
            rest = label_match.group(2).strip()
            rest_plain, rest_formatted = _normalize_markup_text(rest, inline_only=True)
            rest_formatted = _strip_wrapping_paragraph_tags(rest_formatted)
            if plain_label:
                segments.append({
                    "type": "label",
                    "label": label_value,
                    "formatted_label": formatted_label,
                    "text": rest_plain,
                    "formatted_text": rest_formatted,
                })
            elif rest_plain:
                segments.append({"type": "paragraph", "text": rest_plain, "formatted": rest_formatted})
            continue
        plain, formatted = _normalize_markup_text(line, inline_only=True)
        formatted = _strip_wrapping_paragraph_tags(formatted)
        if plain:
            segments.append({"type": "paragraph", "text": plain, "formatted": formatted})
    return segments


def _extract_block_text(block: Any) -> str:
    """
    Extract text content from a block dictionary.
    
    Args:
        block: Block dictionary or other type
        
    Returns:
        Extracted text string
    """
    if isinstance(block, str):
        return block.strip()
    if not isinstance(block, dict):
        return ""
    
    block_type = (block.get("type") or "").strip().lower()
    if block_type == "table":
        return ""
    
    for key in ("text", "title", "description", "label", "name"):
        value = block.get(key)
        if value:
            return str(value).strip()
    return ""


def _collect_nested_child_blocks(item: Dict) -> List[Any]:
    """Collect nested child blocks from various potential keys."""
    nested: List[Any] = []
    for key in ("children", "content", "blocks"):
        value = item.get(key)
        if isinstance(value, dict):
            nested.append(value)
        elif isinstance(value, list):
            nested.extend(value)
    return nested


def _normalize_table_rows(data: Any) -> List[List[str]]:
    """
    Normalize table data into a list of row lists.
    
    Args:
        data: Table data in various formats
        
    Returns:
        List of row lists containing strings
    """
    rows: List[List[str]] = []
    if not isinstance(data, list):
        return rows
    for row in data:
        if isinstance(row, dict):
            cell_values = row.get("cells") or row.get("data") or row.get("row") or []
        elif isinstance(row, list):
            cell_values = row
        else:
            cell_values = [row]
        cells: List[str] = []
        for cell in cell_values:
            plain, _ = _normalize_markup_text(cell)
            cells.append(plain)
        if cells:
            rows.append(cells)
    return rows