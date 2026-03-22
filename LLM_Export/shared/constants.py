"""
Shared constants and configuration for the file export module.
"""

# Environment-based configuration
import os
import logging
import re

# Default paths
DEFAULT_PATH_ENV = os.getenv("PYTHONPATH", r"").rstrip("/")
EXPORT_DIR_ENV = os.getenv("FILE_EXPORT_DIR")
EXPORT_DIR = (EXPORT_DIR_ENV or os.path.join(DEFAULT_PATH_ENV, "output")).rstrip("/")

# Ensure export directory exists
os.makedirs(EXPORT_DIR, exist_ok=True)

BASE_URL_ENV = os.getenv("FILE_EXPORT_BASE_URL")
BASE_URL = (BASE_URL_ENV or "http://localhost:9003/files").rstrip("/")

PERSISTENT_FILES = os.getenv("PERSISTENT_FILES", "false")
FILES_DELAY = int(os.getenv("FILES_DELAY", 60))

# Template paths
DOCS_TEMPLATE_DIR_ENV = os.getenv("DOCS_TEMPLATE_DIR")
DOCS_TEMPLATE_PATH = ((DOCS_TEMPLATE_DIR_ENV or os.path.join(DEFAULT_PATH_ENV, "templates")).rstrip("/"))
os.makedirs(DOCS_TEMPLATE_PATH, exist_ok=True)

# Template state (loaded later)
PPTX_TEMPLATE = None
DOCX_TEMPLATE = None
XLSX_TEMPLATE = None
PPTX_TEMPLATE_PATH = None
DOCX_TEMPLATE_PATH = None
XLSX_TEMPLATE_PATH = None

# Logging configuration
LOG_LEVEL_ENV = os.getenv("LOG_LEVEL")
LOG_FORMAT_ENV = os.getenv(
    "LOG_FORMAT", "%(asctime)s %(levelname)s %(name)s - %(message)s"
)

def _resolve_log_level(val: str | None) -> int:
    if not val:
        return logging.INFO
    v = val.strip()
    if v.isdigit():
        try:
            return int(v)
        except ValueError:
            return logging.INFO
    return getattr(logging, v.upper(), logging.INFO)

# Table separator regex
TABLE_SEPARATOR_RE = re.compile(r"^\s*\|?(?:\s*:?-+:?\s*\|)+\s*$")

# Image source types
IMAGE_SOURCE_UNSPLASH = "unsplash"
IMAGE_SOURCE_PEXELS = "pexels"
IMAGE_SOURCE_LOCAL_SD = "local_sd"

# File format types
FORMAT_PDF = "pdf"
FORMAT_DOCX = "docx"
FORMAT_XLSX = "xlsx"
FORMAT_PPTX = "pptx"
FORMAT_CSV = "csv"
FORMAT_TXT = "txt"
FORMAT_XML = "xml"