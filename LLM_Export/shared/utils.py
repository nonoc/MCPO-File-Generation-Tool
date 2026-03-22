"""
Shared utility functions for the file export module.
"""

import re
import os
import uuid
import datetime
import logging
from pathlib import Path
from typing import Any, List, Optional, Tuple
from pptx.util import Pt as PptPt

log = logging.getLogger(__name__)

# Default configuration values
DEFAULT_FILES_DELAY = 60

# Import shared constants (try relative first, then absolute)
try:
    from .constants import (
        EXPORT_DIR,
        BASE_URL,
        FILES_DELAY,
        TABLE_SEPARATOR_RE,
        IMAGE_SOURCE_UNSPLASH,
        IMAGE_SOURCE_PEXELS,
        IMAGE_SOURCE_LOCAL_SD,
    )
except ImportError:
    from shared.constants import (
        EXPORT_DIR,
        BASE_URL,
        FILES_DELAY,
        TABLE_SEPARATOR_RE,
        IMAGE_SOURCE_UNSPLASH,
        IMAGE_SOURCE_PEXELS,
        IMAGE_SOURCE_LOCAL_SD,
    )


def _public_url(folder_path: str, filename: str) -> str:
    """Build a stable public URL for a generated file."""
    folder = os.path.basename(folder_path).lstrip("/")
    name = filename.lstrip("/")
    return f"{BASE_URL}/{folder}/{name}"


def _generate_unique_folder() -> str:
    """Generate a unique folder name for file exports."""
    folder_name = f"export_{uuid.uuid4().hex[:10]}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
    folder_path = os.path.join(EXPORT_DIR, folder_name)
    os.makedirs(folder_path, exist_ok=True)
    return folder_path


def _generate_filename(folder_path: str, ext: str, filename: Optional[str] = None) -> Tuple[str, str]:
    """Generate a unique filename in the given folder."""
    if not filename:
        filename = f"export_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.{ext}"
    base, ext_part = os.path.splitext(filename)
    filepath = os.path.join(folder_path, filename)
    counter = 1
    while os.path.exists(filepath):
        filename = f"{base}_{counter}{ext_part}"
        filepath = os.path.join(folder_path, filename)
        counter += 1
    return filepath, filename


def _cleanup_files(folder_path: str, delay_minutes: int = DEFAULT_FILES_DELAY):
    """Schedule file cleanup after a delay."""
    def delete_files():
        import time
        import shutil
        time.sleep(delay_minutes * 60)
        try:
            shutil.rmtree(folder_path) 
            log.debug(f"Folder {folder_path} deleted.")
        except Exception as e:
            logging.error(f"Error deleting files: {e}")
    
    import threading
    thread = threading.Thread(target=delete_files)
    thread.start()


# Re-export for convenience
__all__ = [
    "_public_url",
    "_generate_unique_folder", 
    "_generate_filename",
    "_cleanup_files",
    "render_text_with_emojis",
    "_extract_paragraph_style_info",
    "_extract_cell_style_info",
    "dynamic_font_size",
    "_resolve_log_level",
]


def render_text_with_emojis(text: str) -> str:
    """Convert emoji aliases to actual emojis."""
    if not text:
        return ""
    try:
        import emoji
        converted = emoji.emojize(text, language="alias")
        return converted
    except Exception as e:
        log.error(f"Error in emoji conversion: {e}")
        return text


def _extract_paragraph_style_info(para):
    """Extract detailed style information from a paragraph."""
    if not para.runs:
        return {}
    
    first_run = para.runs[0]
    return {
        "font_name": first_run.font.name,
        "font_size": first_run.font.size,
        "bold": first_run.font.bold,
        "italic": first_run.font.italic,
        "underline": first_run.font.underline,
        "color": first_run.font.color.rgb if first_run.font.color else None
    }


def _extract_cell_style_info(cell):
    """Extract style information from a cell."""
    return {
        "style": cell.style.name if hasattr(cell, 'style') else None,
        "text_alignment": cell.paragraphs[0].alignment if cell.paragraphs else None
    }


def dynamic_font_size(content_list, max_chars=400, base_size=28, min_size=12):
    """
    Calculate dynamic font size based on content length.
    
    Args:
        content_list: List of content items
        max_chars: Maximum characters for base size
        base_size: Base font size
        min_size: Minimum font size
        
    Returns:
        PptPt font size object
    """
    total_chars = sum(len(str(line)) for line in content_list)
    ratio = total_chars / max_chars if max_chars > 0 else 1
    if ratio <= 1:
        return PptPt(base_size)
    else:
        new_size = int(base_size / ratio)
        return PptPt(max(min_size, new_size))


def _resolve_log_level(val: str | None) -> int:
    """
    Resolve log level from string or int.
    
    Args:
        val: Log level as string or int
        
    Returns:
        Logging level constant
    """
    if not val:
        return logging.INFO
    v = val.strip()
    if v.isdigit():
        try:
            return int(v)
        except ValueError:
            return logging.INFO
    return getattr(logging, v.upper(), logging.INFO)