# AGENTS.md - Development Guidelines for MCPO File Generation Tool

## Overview

This document provides coding standards, build/test commands, and style guidelines for agentic coding agents working on the MCPO File Generation Tool repository.

## Project Structure

```
MCPO-File-Generation-Tool/
├── LLM_Export/                    # Main Python source code
│   ├── tools/                     # MCP server implementation
│   │   ├── file_export_mcp.py    # Main MCP server (FastMCP)
│   │   ├── file_export_server.py # Static file serving (FastAPI)
│   │   └── __init__.py
│   ├── shared/                    # Shared utilities and constants
│   │   ├── constants.py          # Environment-based configuration
│   │   ├── utils.py              # Common utility functions
│   │   └── styles.py             # PDF styling definitions
│   ├── exporters/                 # Format-specific exporters
│   │   ├── docx_exporter.py      # Word document generation
│   │   ├── pdf_exporter.py       # PDF generation (ReportLab)
│   │   ├── xlsx_exporter.py      # Excel generation
│   │   └── pptx_exporter.py      # PowerPoint generation
│   ├── processing/                # Content processing pipeline
│   │   ├── block_utils.py        # Markdown block parsing
│   │   ├── normalization.py      # Content normalization
│   │   └── text_utils.py         # Text processing helpers
│   ├── functions/                 # Filter functions for Open WebUI
│   │   └── files_metadata_injector.py
│   ├── templates/                 # Document templates (.docx, .pptx, .xlsx)
│   └── output/                    # Generated files directory
├── tests/                         # Test suite
│   └── test_export_content_regression.py
├── Documentation/                 # User documentation
├── config.json                    # MCP server configuration template
└── .github/workflows/containers.yaml  # CI/CD pipeline
```

## Build Commands

### Install Dependencies

```bash
pip install -r LLM_Export/requirements.txt
```

**Required packages:** openpyxl, reportlab, mcp, py7zr, fastapi, uvicorn, python-multipart, markdown2, beautifulsoup4, emoji, python-pptx, python-docx, requests, lxml, PIL/Pillow

### Start File Export Server

```bash
export FILE_EXPORT_DIR=/path/to/output
python LLM_Export/tools/file_export_server.py
```

Or:
```bash
FILE_EXPORT_DIR=/path/to/output uvicorn LLM_Export.tools.file_export_server:app --host 0.0.0.0 --port 9003
```

### Start MCPO Server

```bash
cd LLM_Export && python -m tools.file_export_mcp
```

For development with auto-reload:
```bash
python -m uvicorn tools.file_export_mcp:mcp.app --host 0.0.0.0 --port 8000 --reload
```

## Test Commands

### Run All Tests

```bash
cd LLM_Export && pytest ../tests/ -v
```

### Run Single Test

```bash
# Run by test name
pytest ../tests/ -v -k "test_a_pdf_export"

# Run specific test
pytest ../tests/test_export_content_regression.py::test_a_pdf_export_with_string_markdown -v

# Run with debug logging
pytest ../tests/ -v -s --log-cli-level=DEBUG
```

### Test Coverage

```bash
pytest ../tests/ --cov=../LLM_Export --cov-report=term-missing -v
```

## Code Style Guidelines

### Python Style (PEP 8)

- **Indentation**: 4 spaces per level
- **Line length**: 120 characters maximum
- **Imports**: Standard library first, then third-party, then local imports (relative preferred)
- **Naming**: Functions/variables `snake_case`, Classes `PascalCase`, Constants `UPPER_SNAKE_CASE`, Private `_leading_underscore`
- **Internal exports**: Define `__all__` list in modules

### Import Organization

```python
# 1. Standard library
import os
import re
import logging
from pathlib import Path
from typing import Any, Optional, List, Dict

# 2. Third-party
import requests
from mcp.server.fastmcp import FastMCP
from fastapi import FastAPI
from docx import Document

# 3. Local imports (relative first)
try:
    from ..shared.constants import EXPORT_DIR, BASE_URL
    from ..shared.utils import _generate_unique_folder
except ImportError:
    from shared.constants import EXPORT_DIR, BASE_URL
    from shared.utils import _generate_unique_folder
```

### Logging

- Use module-level logger: `log = logging.getLogger(__name__)`
- Log at appropriate levels: `log.debug()` for diagnostics, `log.info()` for operations, `log.warning()` for issues, `log.error()` for errors with `exc_info=True`

### Error Handling

- Use try/except for network operations (requests)
- Log errors with context: `log.error(f"Error doing X: {e}", exc_info=True)`
- Return None or default values where appropriate
- Use specific exception types: `requests.exceptions.RequestException`

### Type Hints

```python
from typing import Any, Optional, Dict, List

def create_excel(
    data: list[list[str]], 
    filename: str, 
    folder_path: str | None = None, 
    title: str | None = None
) -> dict:
    ...
```

### Configuration

- Environment variables are primary configuration method
- Default values provided for all env vars
- Configuration keys in `config.json` use snake_case
- All constants defined in `shared/constants.py`

### File Naming

- Python modules: `snake_case.py`
- Templates: `Default_Template.{ext}` (docx, pptx, xlsx)
- Generated files: `export_{datetime}.{ext}` or user-provided filename
- Test files: `test_*.py`

### HTML/Markdown Processing

- Use `markdown2` for Markdown to HTML conversion with extras
- Use `BeautifulSoup` for HTML parsing
- Render emojis with `emoji.emojize(text, language="alias")`

### PDF Generation

- Use ReportLab with custom styles from `shared/styles.py`
- Define custom styles in `styles` dict
- Use `SimpleDocTemplate` for document layout
- Handle images with `ReportLabImage`
- Use `Table` and `TableStyle` for tabular data

### Word/PPTX Generation

- Use python-docx for .docx files with template support
- Use python-pptx for .pptx files
- Support template-based generation
- Dynamic font sizing for content using `dynamic_font_size`

## Docker Guidelines

### Environment Variables

**Required:** `FILE_EXPORT_BASE_URL` (default: `http://localhost:9003/files`), `FILE_EXPORT_DIR` (default: `PYTHONPATH/output`)

**Optional:** `PERSISTENT_FILES` (default: `false`), `FILES_DELAY` (default: `60`), `LOG_LEVEL` (default: `INFO`), `IMAGE_SOURCE` (default: `unsplash`)

### Image Sources

- **Unsplash**: Set `UNSPLASH_ACCESS_KEY`
- **Pexels**: Set `PEXELS_ACCESS_KEY`
- **Local Stable Diffusion**: Configure `LOCAL_SD_*` variables

## Branching Strategy

| Branch | Purpose | Docker Tag |
|--------|---------|------------|
| `dev` | Active development | `dev-latest` |
| `alpha` | Post-approval testing | `alpha-latest` |
| `beta` | Optimization & testing | `beta-latest` |
| `release-candidate` | Pre-production validation | `rc-latest` |
| `main` | Production-ready code | `latest` |

## CI/CD

- **Workflow**: `.github/workflows/containers.yaml`
- **Platforms**: linux/amd64, linux/arm64
- **Registry**: ghcr.io

## Common Tasks

### Adding a New File Type

1. Create `_create_{ext}` function in `exporters/{ext}_exporter.py`
2. Implement error handling with try/except and logging
3. Use `_generate_unique_folder()` for output directory
4. Return `{"url": ..., "path": ...}` dict
5. Update `exporters/__init__.py` to re-export the function
6. Add test in `tests/test_export_content_regression.py`

### Modifying Templates

1. Place templates in `LLM_Export/templates/`
2. Template names: `Default_Template.{ext}`
3. Templates are auto-loaded on startup
4. Gracefully handle template load failures with try/except

## Performance Considerations

- Image downloads: 30-second timeout
- File cleanup runs in background threads
- Use BytesIO for in-memory image processing
- Generate unique folder names with UUID + timestamp
- Use relative imports within LLM_Export package

## Context Usage Rules

- Do NOT treat test inputs, sample outputs, or user prompts as source-of-truth for implementation.
- Test data is for validation only and must not be embedded into logic.
- When implementing changes, focus only on:
  - source code files
  - relevant modules
  - minimal required context

- If test data is present in the conversation:
  - use it only to understand the issue
  - do NOT copy or depend on it in the implementation

- Prefer minimal context loading:
  - do not scan unrelated files
  - do not re-read large test payloads unless strictly necessary

## Security

- JWT tokens for OWUI authentication (optional)
- MCPO API key for server authentication (optional)
- Never commit API keys or secrets
- Use environment variables for sensitive data
- Validate all user-provided file paths

## Documentation

- User docs in `Documentation/` directory
- Code comments for complex logic
- Log messages for debugging context
- Docstrings for all public functions

## Notes for Agentic Agents

1. **Always test changes manually** - use pytest with `-k` flag for single tests
2. **Respect existing patterns** - follow established conventions exactly
3. **Use environment variables** - never hardcode configuration
4. **Log appropriately** - add debug logs for new functionality
5. **Handle edge cases** - empty data, missing files, network errors
6. **Docker-first** - most deployments use Docker Compose
7. **Relative imports** - prefer `from ..module import func` over absolute
8. **Test-first** - add regression tests before fixing bugs