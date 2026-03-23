"""
Normalization utilities for content processing.
"""

import logging
from typing import Any, List, Dict, Tuple, Optional

from reportlab.platypus import Paragraph

try:
    from .text_utils import (
        _normalize_markup_text,
        _parse_paragraph_segments,
        _normalize_table_rows,
        _collect_nested_child_blocks,
        _extract_block_text,
    )
    from .block_utils import (
        _finalize_normalized_block,
        flatten_structured_blocks,
    )
except ImportError:
    from processing.text_utils import (
        _normalize_markup_text,
        _parse_paragraph_segments,
        _normalize_table_rows,
        _collect_nested_child_blocks,
        _extract_block_text,
    )
    from processing.block_utils import (
        _finalize_normalized_block,
        flatten_structured_blocks,
    )

log = logging.getLogger(__name__)

# Paragraph-like block types
_PARAGRAPH_TYPES = {"paragraph", "text", "body", "description", "summary", "title"}


def normalize_content_for_export(content: Any) -> List[Dict]:
    """
    Normalize content for export to various formats.
    
    Args:
        content: Content in various formats (string, dict, list)
        
    Returns:
        List of normalized block dictionaries
    """
    def normalize_item(item: Any) -> List[Dict]:
        if item is None:
            return []
        if isinstance(item, str):
            try:
                from .block_utils import _convert_markdown_to_structured
            except ImportError:
                from processing.block_utils import _convert_markdown_to_structured
            structured = _convert_markdown_to_structured(item)
            normalized_blocks: List[Dict] = []
            for block in structured:
                normalized_block = dict(block)
                if normalized_block.get("text"):
                    raw_text = normalized_block["text"]
                    plain, formatted = _normalize_markup_text(raw_text)
                    normalized_block["text"] = plain
                    normalized_block["formatted"] = formatted
                if normalized_block["type"] in _PARAGRAPH_TYPES:
                    normalized_block["segments"] = _parse_paragraph_segments(raw_text)
                if normalized_block.get("type") == "table":
                    table_data = normalized_block.get("data") or []
                    normalized_block["data"] = _normalize_table_rows(table_data)
                normalized_blocks.extend(_finalize_normalized_block(normalized_block))
            return normalized_blocks
        if isinstance(item, dict):
            normalized: Dict = {}
            item_type = (item.get("type") or "").strip().lower()

            if not item_type:
                if item.get("items"):
                    item_type = "list"
                elif item.get("cells") or item.get("rows") or item.get("table"):
                    item_type = "table"
                else:
                    item_type = "paragraph"

            normalized["type"] = item_type

            child_sources = _collect_nested_child_blocks(item)
            text_value = _extract_block_text(item)
            if item_type in _PARAGRAPH_TYPES and not text_value and child_sources:
                return normalize_list(child_sources)

            if item_type == "table":
                data = (
                    item.get("data")
                    or item.get("rows")
                    or item.get("cells")
                    or item.get("content")
                )
                normalized["data"] = _normalize_table_rows(data)
            else:
                raw_source = item.get("text") or item.get("title") or item.get("content") or ""
                normalized_plain, normalized_formatted = _normalize_markup_text(raw_source)
                if normalized_plain:
                    normalized["text"] = normalized_plain
                    normalized["formatted"] = normalized_formatted
                    normalized["raw_text"] = raw_source
                    if item_type in _PARAGRAPH_TYPES:
                        normalized["segments"] = _parse_paragraph_segments(raw_source)

            if child_sources:
                nested_children = normalize_list(child_sources)
                if nested_children:
                    normalized["children"] = nested_children

            if item_type == "list" and item.get("items"):
                normalized["items"] = []
                for list_entry in item.get("items") or []:
                    entry_plain, entry_formatted = _normalize_markup_text(list_entry)
                    if entry_plain:
                        normalized["items"].append({"text": entry_plain, "formatted": entry_formatted})
            return _finalize_normalized_block(normalized)
        if isinstance(item, list):
            return normalize_list(item)
        plain, formatted = _normalize_markup_text(item)
        if not plain:
            return []
        fallback_block = {"type": "paragraph", "text": plain, "formatted": formatted}
        fallback_block["segments"] = _parse_paragraph_segments(plain)
        return _finalize_normalized_block(fallback_block)

    def normalize_list(items: List[Any] | None) -> List[Dict]:
        if not isinstance(items, list):
            return []
        result: List[Dict] = []
        for child in items:
            result.extend(normalize_item(child))
        return result

    return normalize_item(content)


def _render_structured_content(structured_blocks: List[Any]) -> List:
    """Render structured blocks into a story/list for PDF generation."""
    from reportlab.platypus import Paragraph
    from ..shared.styles import styles, get_list_flowable
    from ..shared.utils import render_text_with_emojis
    
    renderer = _StructuredContentRenderer(styles, render_text_with_emojis, get_list_flowable)
    return renderer.build(structured_blocks)


class _StructuredContentRenderer:
    """Renderer for structured content blocks."""
    
    def __init__(self, styles, render_func, list_flowable_func):
        self.story: List[Any] = []
        self.section_counters: List[int] = []
        self.styles = styles
        self.render_func = render_func
        self.list_flowable_func = list_flowable_func

    def build(self, blocks: List[Any]) -> List[Any]:
        """Build the story from structured blocks."""
        for block, depth in flatten_structured_blocks(blocks):
            self._render_block(block, depth)
        if not self.story:
            self.story.append(Paragraph("Empty Content", self.styles["CustomNormal"]))
        return self.story

    def _render_block(self, block: Any, section_depth: int) -> None:
        """Render a single block."""
        if isinstance(block, str):
            self._append_paragraph(block)
            return
        if not isinstance(block, dict):
            return

        block_type = (block.get("type") or "").strip().lower()

        if block_type == "title":
            self._render_title(block)
            return

        if block_type == "section":
            self._render_section(block, section_depth)
            return

        if block_type == "table":
            self._render_table(block)
            return

        if block_type == "label_paragraph":
            self._render_label_paragraph(block)
            return

        if block_type in {"paragraph", "text", "body", "description", "summary", "heading", "subheading"}:
            self._append_paragraph(block.get("text"), block.get("formatted"))
            return
        elif block_type in {"sources", "source", "references"}:
            self._render_sources(block)
            return
        elif "list" in block_type or "bullet" in block_type:
            list_flowable = self._build_list_flowable(block, depth=section_depth)
            if list_flowable is not None:
                self.story.append(list_flowable)
                self.story.append(Paragraph(" ", self.styles["CustomNormal"]))
            return
        else:
            text = self._extract_text(block)
            self._append_paragraph(text, self._extract_formatted(block))

    def _render_title(self, block: Dict) -> None:
        """Render a title block."""
        text = self._extract_text(block)
        if text:
            self.story.append(Paragraph(self.render_func(text), self.styles["StructuredDocumentTitle"]))

    def _render_section(self, block: Dict, depth: int) -> None:
        """Render a section block with numbering."""
        heading_text = self._extract_text(block, keys=("title", "text", "heading"))
        heading_text = heading_text or "Untitled Section"
        numbering = self._increment_section_counter(depth)
        display_text = f"{numbering} {heading_text}" if numbering else heading_text
        self.story.append(Paragraph(self.render_func(display_text), self._heading_style(depth)))
        self.story.append(Paragraph(" ", self.styles["CustomNormal"]))

    def _render_table(self, block: Dict) -> None:
        """Render a table block."""
        from ..shared.styles import get_table_style
        
        table_data = _normalize_table_rows(
            block.get("data") or block.get("rows") or block.get("cells") or block.get("content") or []
        )
        if table_data:
            table = self._create_table(table_data)
            table.setStyle(get_table_style())
            self.story.append(table)
            self.story.append(Paragraph(" ", self.styles["CustomNormal"]))

    def _render_sources(self, block: Dict) -> None:
        """Render sources/references section."""
        heading = block.get("title") or block.get("text") or "Sources"
        self.story.append(Paragraph(self.render_func(heading), self.styles["StructuredSourcesHeading"]))
        entries = self._extract_list_entries(block)
        if not entries:
            entries = block.get("children") or []
        for entry in entries:
            text = self._extract_text(entry)
            if text:
                self.story.append(Paragraph(self.render_func(text), self.styles["StructuredSourcesItem"]))
        self.story.append(Paragraph(" ", self.styles["CustomNormal"]))

    def _build_list_flowable(self, block: Dict, depth: int):
        """Build a list flowable from a block."""
        entries = self._extract_list_entries(block)
        if not entries:
            return None
        list_items = []
        for entry in entries:
            entry_flow = []
            entry_text = self._extract_text(entry)
            if entry_text:
                entry_formatted = self._extract_formatted(entry)
                entry_flow.append(Paragraph(self.render_func(entry_formatted or entry_text), self.styles["StructuredListItem"]))
            if isinstance(entry, dict):
                nested = self._extract_list_entries(entry)
                if nested:
                    nested_flowable = self._build_list_flowable({"items": nested, "ordered": entry.get("ordered")}, depth + 1)
                    if nested_flowable:
                        entry_flow.append(nested_flowable)
            if entry_flow:
                list_items.append(self._create_list_item(entry_flow))
        if not list_items:
            return None
        ordered = self._determine_ordering(block)
        return self.list_flowable_func(list_items, ordered=ordered, depth=depth)

    def _determine_ordering(self, block: Dict) -> bool:
        """Determine if a list is ordered."""
        ordered = block.get("ordered")
        if isinstance(ordered, str):
            return ordered.lower() in {"true", "1", "yes", "ordered", "ol"}
        if isinstance(ordered, bool):
            return ordered
        block_type = (block.get("type") or "").lower()
        return "ordered" in block_type or block_type in {"ol", "ordered_list"}

    def _extract_list_entries(self, block: Dict) -> List[Any]:
        """Extract list entries from a block."""
        entries = block.get("items")
        if isinstance(entries, list) and entries:
            return entries
        entries = block.get("children")
        if isinstance(entries, list) and entries:
            return entries
        entries = block.get("entries")
        if isinstance(entries, list) and entries:
            return entries
        block_type = (block.get("type") or "").lower()
        if block_type in {"bullet", "list_item"}:
            text = block.get("text")
            if text:
                return [text]
        return []

    def _increment_section_counter(self, depth: int) -> str:
        """Increment section counters for numbering."""
        if depth < 1:
            depth = 1
        while len(self.section_counters) < depth:
            self.section_counters.append(0)
        if len(self.section_counters) > depth:
            self.section_counters = self.section_counters[:depth]
        self.section_counters[depth - 1] += 1
        return ".".join(str(num) for num in self.section_counters if num > 0)

    def _heading_style(self, depth: int):
        """Get the appropriate heading style for a depth."""
        if depth <= 1:
            return self.styles["CustomHeading1"]
        if depth == 2:
            return self.styles["CustomHeading2"]
        return self.styles["CustomHeading3"]

    def _append_paragraph(self, text: str | None, formatted: str | None = None, style=None) -> None:
        """Append a paragraph to the story."""
        if not text and not formatted:
            return
        paragraph_style = style or self.styles["StructuredParagraph"]
        content = formatted or text or ""
        markup = self._prepare_inline_markup(content)
        self.story.append(Paragraph(self.render_func(markup), paragraph_style))

    def _prepare_inline_markup(self, content: str) -> str:
        """Prepare inline markup for paragraphs."""
        if not content:
            return ""
        return (
            content.replace("<strong>", "<b>")
            .replace("</strong>", "</b>")
            .replace("<em>", "<i>")
            .replace("</em>", "</i>")
        )

    def _render_label_paragraph(self, block: Dict) -> None:
        """Render a label paragraph block."""
        label = block.get("label")
        formatted_label = block.get("formatted_label")
        body = block.get("text")
        formatted_body = block.get("formatted")
        plain_parts = [str(part).strip() for part in (label, body) if part and str(part).strip()]
        formatted_parts = [str(part) for part in (formatted_label, formatted_body) if part]
        combined_text = " ".join(plain_parts) if plain_parts else None
        combined_formatted = " ".join(formatted_parts) if formatted_parts else None
        self._append_paragraph(combined_text, combined_formatted)

    def _extract_text(self, block: Any, keys=None) -> str:
        """Extract text from a block."""
        if keys is None:
            keys = ("text", "title", "content", "description", "label", "name")
        if isinstance(block, str):
            return block.strip()
        if not isinstance(block, dict):
            return ""
        for key in keys:
            value = block.get(key)
            if value:
                return str(value).strip()
        return ""

    def _extract_formatted(self, block: Any) -> str | None:
        """Extract formatted content from a block."""
        if isinstance(block, dict):
            value = block.get("formatted")
            if value is not None:
                return str(value)
        return None

    def _create_table(self, table_data):
        """Create a table object."""
        from reportlab.platypus import Table as ReportLabTable
        return ReportLabTable(table_data, repeatRows=1)

    def _create_list_item(self, content):
        """Create a list item object."""
        from reportlab.platypus import ListItem
        return ListItem(content)