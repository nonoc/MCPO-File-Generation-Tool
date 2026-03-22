"""
Processing package for content normalization and text handling.
"""

from .text_utils import (
    _normalize_markup_text,
    _strip_wrapping_paragraph_tags,
    _parse_paragraph_segments,
    _extract_block_text,
    _collect_nested_child_blocks,
    _normalize_table_rows,
)
from .block_utils import (
    _expand_paragraph_block,
    _finalize_normalized_block,
    _convert_markdown_to_structured,
    flatten_structured_blocks,
)
from .normalization import (
    normalize_content_for_export,
    _render_structured_content,
    _StructuredContentRenderer,
)

__all__ = [
    "_normalize_markup_text",
    "_strip_wrapping_paragraph_tags",
    "_parse_paragraph_segments",
    "_extract_block_text",
    "_collect_nested_child_blocks",
    "_normalize_table_rows",
    "_expand_paragraph_block",
    "_finalize_normalized_block",
    "_convert_markdown_to_structured",
    "flatten_structured_blocks",
    "normalize_content_for_export",
    "_render_structured_content",
    "_StructuredContentRenderer",
]