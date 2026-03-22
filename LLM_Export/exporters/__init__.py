"""
Exporters package for document format-specific export functions.
"""

try:
    from .docx_exporter import (
        _create_word,
        _normalize_content_for_export,
        _render_structured_docx,
        _add_docx_paragraph,
        _apply_formatted_html_to_paragraph,
        _add_docx_heading,
        _extract_block_text,
    )
    from .pdf_exporter import (
        _create_pdf,
        _build_markdown_story,
        _render_html_elements,
        process_list_items,
        search_image,
    )
    from .xlsx_exporter import (
        _create_excel,
        add_auto_sized_review_comment,
    )
    from .pptx_exporter import (
        _create_presentation,
        dynamic_font_size,
        search_image,
    )
except ImportError:
    from exporters.docx_exporter import (
        _create_word,
        _normalize_content_for_export,
        _render_structured_docx,
        _add_docx_paragraph,
        _apply_formatted_html_to_paragraph,
        _add_docx_heading,
        _extract_block_text,
    )
    from exporters.pdf_exporter import (
        _create_pdf,
        _build_markdown_story,
        _render_html_elements,
        process_list_items,
        search_image,
    )
    from exporters.xlsx_exporter import (
        _create_excel,
        add_auto_sized_review_comment,
    )
    from exporters.pptx_exporter import (
        _create_presentation,
        dynamic_font_size,
        search_image,
    )

__all__ = [
    "_create_word",
    "_normalize_content_for_export",
    "_render_structured_docx",
    "_add_docx_paragraph",
    "_apply_formatted_html_to_paragraph",
    "_add_docx_heading",
    "_extract_block_text",
    "_create_pdf",
    "_build_markdown_story",
    "_render_html_elements",
    "process_list_items",
    "search_image",
    "_create_excel",
    "add_auto_sized_review_comment",
    "_create_presentation",
    "dynamic_font_size",
]