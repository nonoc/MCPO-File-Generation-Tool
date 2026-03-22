"""
Shared types and data structures for the file export module.
"""

from typing import Any, Optional, Dict, List, Tuple, Union

# Common type aliases
Block = Dict[str, Any]
ContentItem = Union[str, Dict[str, Any], List[Any]]
NormalizedBlock = Dict[str, Any]
TableData = List[List[str]]
SlideData = Dict[str, Any]

# Structured content block types
BLOCK_TITLE = "title"
BLOCK_SECTION = "section"
BLOCK_PARAGRAPH = "paragraph"
BLOCK_HEADING = "heading"
BLOCK_SUBHEADING = "subheading"
BLOCK_BULLET = "bullet"
BLOCK_LIST = "list"
BLOCK_TABLE = "table"
BLOCK_IMAGE = "image"
BLOCK_SOURCES = "sources"
BLOCK_SOURCES_ITEM = "sources_item"
BLOCK_LABEL_PARAGRAPH = "label_paragraph"

# Paragraph-like block types
PARAGRAPH_TYPES = {
    "paragraph", "text", "body", "description", "summary", "title", 
    "heading", "subheading", "bullet"
}

# Source/Reference types
SOURCE_TYPES = {"sources", "source", "references"}

# Export format types
FORMAT_PDF = "pdf"
FORMAT_DOCX = "docx"
FORMAT_XLSX = "xlsx"
FORMAT_PPTX = "pptx"
FORMAT_CSV = "csv"
FORMAT_TXT = "txt"
FORMAT_XML = "xml"

# Image position options
IMAGE_POS_LEFT = "left"
IMAGE_POS_RIGHT = "right"
IMAGE_POS_TOP = "top"
IMAGE_POS_BOTTOM = "bottom"

# Image size options
IMAGE_SIZE_SMALL = "small"
IMAGE_SIZE_MEDIUM = "medium"
IMAGE_SIZE_LARGE = "large"