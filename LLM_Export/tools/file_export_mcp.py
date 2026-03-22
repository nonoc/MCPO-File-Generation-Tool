"""
MCPO File Export Server - Main entry point.

This module provides the MCP server interface for file export functionality.
It imports the actual implementation from the new modular structure.
"""

import asyncio
import os
import re
from pathlib import Path

import markdown2
from bs4 import BeautifulSoup
from reportlab.platypus import Paragraph as RLParagraph, Table as RLTable

# Import from new modules (try relative first, then absolute)
try:
    from ..shared.constants import (
        EXPORT_DIR,
        BASE_URL,
        FILES_DELAY,
        PERSISTENT_FILES,
        TABLE_SEPARATOR_RE,
        IMAGE_SOURCE_UNSPLASH,
        IMAGE_SOURCE_PEXELS,
        IMAGE_SOURCE_LOCAL_SD,
        DOCS_TEMPLATE_PATH,
        PPTX_TEMPLATE,
        DOCX_TEMPLATE,
        XLSX_TEMPLATE,
        PPTX_TEMPLATE_PATH,
        DOCX_TEMPLATE_PATH,
        XLSX_TEMPLATE_PATH,
    )
    from ..shared.utils import (
        _public_url,
        _generate_unique_folder,
        _generate_filename,
        _cleanup_files,
        render_text_with_emojis,
        _extract_paragraph_style_info,
        _extract_cell_style_info,
        dynamic_font_size,
        _resolve_log_level,
    )
    from ..shared.images import search_image, search_local_sd, search_unsplash, search_pexels
    from ..processing.block_utils import _convert_markdown_to_structured, flatten_structured_blocks
    from ..processing.normalization import normalize_content_for_export, _StructuredContentRenderer, _render_structured_content
    from ..exporters import (
        _create_word,
        _create_pdf,
        _create_excel,
        _create_presentation,
        add_auto_sized_review_comment,
    )
    from ..export_dispatcher import (
        _extract_document_title,
        _create_csv,
        _create_raw_file,
    )
except ImportError:
    from shared.constants import (
        EXPORT_DIR,
        BASE_URL,
        FILES_DELAY,
        PERSISTENT_FILES,
        TABLE_SEPARATOR_RE,
        IMAGE_SOURCE_UNSPLASH,
        IMAGE_SOURCE_PEXELS,
        IMAGE_SOURCE_LOCAL_SD,
        DOCS_TEMPLATE_PATH,
        PPTX_TEMPLATE,
        DOCX_TEMPLATE,
        XLSX_TEMPLATE,
        PPTX_TEMPLATE_PATH,
        DOCX_TEMPLATE_PATH,
        XLSX_TEMPLATE_PATH,
    )
    from shared.utils import (
        _public_url,
        _generate_unique_folder,
        _generate_filename,
        _cleanup_files,
        render_text_with_emojis,
        _extract_paragraph_style_info,
        _extract_cell_style_info,
        dynamic_font_size,
        _resolve_log_level,
    )
    from shared.images import search_image, search_local_sd, search_unsplash, search_pexels
    from processing.block_utils import _convert_markdown_to_structured, flatten_structured_blocks
    from processing.normalization import normalize_content_for_export, _StructuredContentRenderer, _render_structured_content
    from exporters import (
        _create_word,
        _create_pdf,
        _create_excel,
        _create_presentation,
        add_auto_sized_review_comment,
    )
    from export_dispatcher import (
        _extract_document_title,
        _create_csv,
        _create_raw_file,
    )

# Re-export functions for backward compatibility
__all__ = [
    # Constants
    "EXPORT_DIR",
    "BASE_URL",
    "FILES_DELAY",
    "PERSISTENT_FILES",
    "TABLE_SEPARATOR_RE",
    "IMAGE_SOURCE_UNSPLASH",
    "IMAGE_SOURCE_PEXELS",
    "IMAGE_SOURCE_LOCAL_SD",
    "DOCS_TEMPLATE_PATH",
    "PPTX_TEMPLATE",
    "DOCX_TEMPLATE",
    "XLSX_TEMPLATE",
    "PPTX_TEMPLATE_PATH",
    "DOCX_TEMPLATE_PATH",
    "XLSX_TEMPLATE_PATH",
    # Utils
    "_public_url",
    "_generate_unique_folder",
    "_generate_filename",
    "_cleanup_files",
    "render_text_with_emojis",
    "_extract_paragraph_style_info",
    "_extract_cell_style_info",
    "dynamic_font_size",
    "_resolve_log_level",
    # Images
    "search_image",
    "search_local_sd",
    "search_unsplash",
    "search_pexels",
    # Processing
    "_convert_markdown_to_structured",
    "flatten_structured_blocks",
    "normalize_content_for_export",
    "_StructuredContentRenderer",
    # Exporters
    "_create_word",
    "_create_pdf",
    "_create_excel",
    "_create_presentation",
    "add_auto_sized_review_comment",
    # Export dispatcher
    "_extract_document_title",
    "_create_csv",
    "_create_raw_file",
    # Additional functions for backward compatibility
    "render_html_elements",
    "create_file",
    "render_structured_content",
]


def render_html_elements(soup):
    """Render HTML elements into a PDF story."""
    try:
        from ..shared.styles import styles
    except ImportError:
        from shared.styles import styles
    from reportlab.platypus import ListFlowable
    
    story = []
    for elem in soup.find_all(["h1", "h2", "h3", "p", "ul", "ol", "li", "table"]):
        if elem.name in ("h1", "h2", "h3"):
            level = int(elem.name[1])
            style_name = f"Heading {level}"
            if style_name in styles:
                story.append(RLParagraph(elem.get_text(), styles[style_name]))
            else:
                story.append(RLParagraph(elem.get_text(), styles["Normal"]))
        elif elem.name == "p":
            story.append(RLParagraph(elem.get_text(), styles["Normal"]))
        elif elem.name in ("ul", "ol"):
            list_items = []
            for li in elem.find_all("li"):
                list_items.append(RLParagraph(li.get_text(), styles["Normal"]))
            if elem.name == "ul":
                story.append(ListFlowable(list_items, bulletType="bullet"))
            else:
                story.append(ListFlowable(list_items, bulletType="1"))
        elif elem.name == "table":
            rows = []
            for tr in elem.find_all("tr"):
                row = []
                for td in tr.find_all(["td", "th"]):
                    row.append(td.get_text())
                if row:
                    rows.append(row)
            if rows:
                story.append(RLTable(rows))
    return story


async def create_file(data: dict, persistent: bool = PERSISTENT_FILES) -> dict:
    """Create a file from data dictionary."""
    format_type = data.get("format", "pdf").lower()
    filename = data.get("filename", "export")
    content = data.get("content", "")
    title = data.get("title")
    
    folder_path = _generate_unique_folder()
    
    if format_type == "pdf":
        result = _create_pdf(content, filename, folder_path=folder_path)
    elif format_type == "docx":
        result = _create_word(content, filename, folder_path=folder_path)
    elif format_type == "xlsx":
        result = _create_excel(content, filename, folder_path=folder_path)
    elif format_type == "pptx":
        # Convert content to slides format
        slides_data = [{"title": "Slide", "content": [content]}]
        result = _create_presentation(slides_data, filename, folder_path=folder_path)
    elif format_type == "csv":
        result = _create_csv(content, filename, folder_path=folder_path)
    else:
        result = _create_raw_file(content, filename, folder_path=folder_path)
    
    # Handle persistent files
    if not persistent:
        import threading
        import time
        delay = int(os.environ.get("FILES_DELAY", "60"))
        
        def cleanup():
            time.sleep(delay)
            try:
                _cleanup_files(folder_path, delay)
            except Exception:
                pass
        
        thread = threading.Thread(target=cleanup)
        thread.daemon = True
        thread.start()
    
    return result


def render_structured_content(structured_blocks):
    """Render structured blocks into a PDF story."""
    return _render_structured_content(structured_blocks)