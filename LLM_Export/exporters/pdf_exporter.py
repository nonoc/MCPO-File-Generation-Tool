"""
PDF exporter module using ReportLab.
"""

import os
import logging
from typing import Any, List, Optional

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, ListFlowable, ListItem, Image as ReportLabImage, Table as ReportLabTable, TableStyle
from reportlab.lib import colors
from reportlab.lib.units import mm

# Import from new modules (try relative first)
try:
    from ..shared.styles import styles, get_table_style
    from ..shared.utils import _generate_unique_folder, _public_url, render_text_with_emojis
    from ..processing.normalization import normalize_content_for_export, _StructuredContentRenderer
except ImportError:
    from shared.styles import styles, get_table_style
    from shared.utils import _generate_unique_folder, _public_url, render_text_with_emojis
    from processing.normalization import normalize_content_for_export, _StructuredContentRenderer

log = logging.getLogger(__name__)


def _create_pdf(content: Any, filename: str, folder_path: Optional[str] = None) -> dict:
    """
    Create a PDF file from content.
    
    Args:
        content: Content to render (string, dict, or list)
        filename: Output filename
        folder_path: Optional folder path for output
        
    Returns:
        Dict with url and path to generated file
    """
    log.debug("Creating PDF file")
    if folder_path is None:
        folder_path = _generate_unique_folder()
    
    if filename:
        filepath = os.path.join(folder_path, filename)
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        fname = filename
    else:
        try:
            from ..shared.utils import _generate_filename
        except ImportError:
            from shared.utils import _generate_filename
        filepath, fname = _generate_filename(folder_path, "pdf")

    try:
        normalized_blocks = normalize_content_for_export(content)
        if normalized_blocks:
            story = _StructuredContentRenderer(
                styles=styles, 
                render_func=render_text_with_emojis,
                list_flowable_func=ListFlowable
            ).build(normalized_blocks)
        else:
            story = _build_markdown_story(content)
    except Exception as e:
        log.error(f"Error processing content: {e}")
        story = [Paragraph("Error in PDF generation", styles["CustomNormal"])]

    doc = SimpleDocTemplate(filepath, topMargin=54, bottomMargin=54, leftMargin=54, rightMargin=54)
    try:
        doc.build(story)
    except Exception as e:
        log.error(f"Error building PDF {fname}: {e}", exc_info=True)
        fallback = SimpleDocTemplate(filepath)
        fallback.build([Paragraph("Error in PDF generation", styles["CustomNormal"])])

    return {"url": _public_url(folder_path, fname), "path": filepath}


def _build_markdown_story(source: Any) -> List:
    """Build a story from markdown source."""
    import re
    import markdown2
    from bs4 import BeautifulSoup
    from reportlab.platypus import Spacer
    
    md_parts: List[str] = []

    def append_paragraph(value: Any) -> None:
        normalized = (value or "").strip()
        if normalized:
            md_parts.append(normalized)

    items = source if isinstance(source, list) else [source]
    for entry in items:
        if isinstance(entry, str):
            append_paragraph(entry)
            continue
        if isinstance(entry, dict):
            t = (entry.get("type") or "").lower()
            text_value = entry.get("text") or entry.get("title") or ""
            if t == "title":
                append_paragraph(f"# {text_value}")
            elif t == "subtitle":
                append_paragraph(f"## {text_value}")
            elif t == "paragraph":
                append_paragraph(entry.get("text", ""))
            elif t == "list":
                append_paragraph("\n".join([f"- {x}" for x in entry.get("items", [])]))
            elif t in ("image", "image_query"):
                query = entry.get("query", "")
                if query:
                    append_paragraph(f"![Image](image_query: {query})")
            else:
                append_paragraph(text_value)
        else:
            append_paragraph(str(entry))

    md_text = "\n\n".join(md_parts) or ""

    def replace_image_query(match):
        query = match.group(1).strip()
        from ..shared.images import search_image
        image_url = search_image(query)
        if image_url:
            return f'\n\n<img src="{image_url}" alt="Image: {query}" />\n\n'
        return ""

    md_text = re.sub(r'!\[[^\]]*\]\(\s*image_query:\s*([^)]+)\)', replace_image_query, md_text)
    html = markdown2.markdown(md_text, extras=['fenced-code-blocks','tables','break-on-newline','cuddled-lists'])
    soup = BeautifulSoup(html, "html.parser")
    
    from ..shared.utils import render_text_with_emojis
    story = _render_html_elements(soup) or [Paragraph("Empty Content", styles["CustomNormal"])]
    return story


def _render_html_elements(soup):
    """Render HTML elements to PDF story."""
    from reportlab.platypus import Spacer
    from reportlab.platypus import Image as ReportLabImage
    import requests
    from io import BytesIO
    
    story = []
    element_count = 0
    for elem in soup.children:
        element_count += 1
        if isinstance(elem, str):
            text = str(elem).strip()
            if text:
                story.append(Paragraph(render_text_with_emojis(text), styles["CustomNormal"]))
                story.append(Spacer(1, 6))
            continue
            
        tag_name = getattr(elem, 'name', None)
        if not tag_name:
            continue
            
        if tag_name == "h1":
            text = render_text_with_emojis(elem.get_text().strip())
            story.append(Paragraph(text, styles["CustomHeading1"]))
            story.append(Spacer(1, 10))
        elif tag_name == "h2":
            text = render_text_with_emojis(elem.get_text().strip())
            story.append(Paragraph(text, styles["CustomHeading2"]))
            story.append(Spacer(1, 8))
        elif tag_name == "h3":
            text = render_text_with_emojis(elem.get_text().strip())
            story.append(Paragraph(text, styles["CustomHeading3"]))
            story.append(Spacer(1, 6))
        elif tag_name == "p":
            imgs = elem.find_all("img")
            if imgs:
                for img_tag in imgs:
                    src = img_tag.get("src")
                    alt = img_tag.get("alt", "[Image]")
                    try:
                        if src and src.startswith("http"):
                            response = requests.get(src)
                            response.raise_for_status()
                            img_data = BytesIO(response.content)
                            img = ReportLabImage(img_data, 200, 150)
                        else:
                            img = ReportLabImage(src, 200, 150)
                        story.append(img)
                        story.append(Spacer(1, 10))
                    except Exception as e:
                        log.error(f"Error loading image {src}: {e}")
                        story.append(Paragraph(f"[Image: {alt}]", styles["CustomNormal"]))
                        story.append(Spacer(1, 6))
            else:
                text = render_text_with_emojis(elem.get_text().strip())
                if text:
                    story.append(Paragraph(text, styles["CustomNormal"]))
                    story.append(Spacer(1, 6))
        elif tag_name in ["ul", "ol"]:
            is_ordered = tag_name == "ol"
            items = process_list_items(elem, is_ordered)
            if items:
                story.append(ListFlowable(
                    items,
                    bulletType='1' if is_ordered else 'bullet',
                    leftIndent=10 * mm,
                    bulletIndent=5 * mm,
                    spaceBefore=6,
                    spaceAfter=10
                ))
        elif tag_name == "blockquote":
            text = render_text_with_emojis(elem.get_text().strip())
            if text:
                story.append(Paragraph(text, styles["CustomNormal"]))
                story.append(Spacer(1, 8))
        elif tag_name in ["code", "pre"]:
            text = elem.get_text().strip()
            if text:
                story.append(Paragraph(text, styles["CustomCode"]))
                story.append(Spacer(1, 6 if tag_name == "code" else 8))
        elif tag_name == "img":
            src = elem.get("src")
            alt = elem.get("alt", "[Image]")
            if src is not None:
                try:
                    if src.startswith("image_query:"):
                        query = src.replace("image_query:", "").strip()
                        image_url = search_image(query)
                        if image_url:
                            response = requests.get(image_url)
                            response.raise_for_status()
                            img_data = BytesIO(response.content)
                            img = ReportLabImage(img_data, 200, 150)
                            story.append(img)
                            story.append(Spacer(1, 10))
                        else:
                            story.append(Paragraph(f"[Image not found for: {query}]", styles["CustomNormal"]))
                            story.append(Spacer(1, 6))
                    elif src.startswith("http"):
                        response = requests.get(src)
                        response.raise_for_status()
                        img_data = BytesIO(response.content)
                        img = ReportLabImage(img_data, 200, 150)
                        story.append(img)
                        story.append(Spacer(1, 10))
                    else:
                        if os.path.exists(src):
                            img = ReportLabImage(src, 200, 150)
                            story.append(img)
                            story.append(Spacer(1, 10))
                        else:
                            story.append(Paragraph(f"[Image not found: {src}]", styles["CustomNormal"]))
                            story.append(Spacer(1, 6))
                except Exception as e:
                    log.error(f"Error processing image {src}: {e}")
                    story.append(Paragraph(f"[Image: {alt}]", styles["CustomNormal"]))
                    story.append(Spacer(1, 6))
            else:
                story.append(Paragraph(f"[Image: {alt} (missing source)]", styles["CustomNormal"]))
                story.append(Spacer(1, 6))
        elif tag_name == "br":
            story.append(Spacer(1, 6))
        else:
            text = elem.get_text().strip()
            if text:
                story.append(Paragraph(render_text_with_emojis(text), styles["CustomNormal"]))
                story.append(Spacer(1, 6))
    
    return story


def process_list_items(ul_or_ol_element, is_ordered=False):
    """Process list items for PDF generation."""
    items = []
    bullet_type = '1' if is_ordered else 'bullet'
    for li in ul_or_ol_element.find_all('li', recursive=False):
        li_text_parts = []
        for content in li.contents:
            if isinstance(content, str):
                li_text_parts.append(str(content))
            elif content.name not in ['ul', 'ol']:
                li_text_parts.append(content.get_text())
        li_text = ''.join(li_text_parts).strip()
        list_item_paragraph = None
        if li_text:
            rendered_text = render_text_with_emojis(li_text)
            list_item_paragraph = Paragraph(rendered_text, styles["CustomListItem"])
        sub_lists = li.find_all(['ul', 'ol'], recursive=False)
        sub_flowables = []
        if list_item_paragraph:
            sub_flowables.append(list_item_paragraph)
        for sub_list in sub_lists:
            is_sub_ordered = sub_list.name == 'ol'
            nested_items = process_list_items(sub_list, is_sub_ordered)
            if nested_items:
                nested_list_flowable = ListFlowable(
                    nested_items,
                    bulletType='1' if is_sub_ordered else 'bullet',
                    leftIndent=10 * mm,
                    bulletIndent=5 * mm,
                    spaceBefore=2,
                    spaceAfter=2
                )
                sub_flowables.append(nested_list_flowable)
        if sub_flowables:
            items.append(ListItem(sub_flowables))
    return items


def search_image(query):
    """Search for an image using configured source."""
    from ..shared.images import search_image as img_search
    return img_search(query)