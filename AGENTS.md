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
│   ├── functions/                 # Filter functions for Open WebUI
│   │   └── files_metadata_injector.py
│   ├── templates/                 # Document templates (.docx, .pptx, .xlsx)
│   └── output/                    # Generated files directory
├── Documentation/                 # User documentation
├── config.json                    # MCP server configuration template
└── .github/workflows/containers.yaml  # CI/CD pipeline
```

## Build Commands

### Install Dependencies

```bash
pip install -r LLM_Export/requirements.txt
```

**Required packages:**
- openpyxl (Excel file generation)
- reportlab (PDF generation)
- mcp (Model Context Protocol)
- py7zr (7z archive support)
- fastapi (HTTP server framework)
- uvicorn (ASGI server)
- python-multipart (multipart/form-data handling)
- markdown2 (Markdown to HTML conversion)
- beautifulsoup4 (HTML parsing)
- emoji (Emoji support)
- python-pptx (PowerPoint generation)
- python-docx (Word document generation)
- requests (HTTP client)
- lxml (XML processing)
- PIL/Pillow (Image processing)

### Start File Export Server

```bash
export FILE_EXPORT_DIR=/path/to/output
python LLM_Export/tools/file_export_server.py
```

Or with environment variables:
```bash
FILE_EXPORT_DIR=/path/to/output uvicorn LLM_Export.tools.file_export_server:app --host 0.0.0.0 --port 9003
```

### Start MCPO Server

```bash
cd LLM_Export
python -m tools.file_export_mcp
```

For development with auto-reload:
```bash
python -m uvicorn tools.file_export_mcp:mcp.app --host 0.0.0.0 --port 8000 --reload
```

## Test Commands

**No formal test suite exists.** The project is currently production-only without automated tests.

### Manual Testing

1. **Start the file server:**
   ```bash
   FILE_EXPORT_DIR=/tmp/test-output python LLM_Export/tools/file_export_server.py
   ```

2. **Start the MCPO server (in another terminal):**
   ```bash
   cd LLM_Export
   FILE_EXPORT_BASE_URL=http://localhost:9003/files FILE_EXPORT_DIR=/tmp/test-output python -m tools.file_export_mcp
   ```

3. **Verify file generation** by using Open WebUI or calling the MCP server endpoints.

### Docker Testing

Build and run with docker-compose:
```bash
docker-compose -f LLM_Export/Example_docker-compose.yaml up -d
```

## Code Style Guidelines

### Python Style (PEP 8)

- **Indentation**: 4 spaces per level
- **Line length**: 120 characters maximum
- **Imports**: Standard library first, then third-party, then local imports
- **Naming**:
  - Functions/variables: `snake_case`
  - Classes: `PascalCase`
  - Constants: `UPPER_SNAKE_CASE`
  - Private: `_leading_underscore`

### Import Organization

```python
# 1. Standard library
import re
import os
import json
import uuid
import datetime
import logging
from pathlib import Path

# 2. Third-party
import requests
from mcp.server.fastmcp import FastMCP
from fastapi import FastAPI

# 3. Local imports
from tools.file_export_mcp import some_function
```

### Logging

- Use module-level logger: `log = logging.getLogger(__name__)`
- Log levels: DEBUG for development, INFO for production
- Always log at appropriate levels:
  - `log.debug()` for detailed diagnostic info
  - `log.info()` for normal operations
  - `log.warning()` for recoverable issues
  - `log.error()` for errors with stack traces

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

### File Naming

- Python modules: `snake_case.py`
- Templates: `Default_Template.{ext}` (docx, pptx, xlsx)
- Generated files: `export_{datetime}.{ext}` or user-provided filename

### HTML/Markdown Processing

- Use `markdown2` for Markdown to HTML conversion
- Use `BeautifulSoup` for HTML parsing
- Render emojis with `emoji.emojize(text, language="alias")`

### PDF Generation

- Use ReportLab with custom styles
- Define custom styles in `styles` dict
- Use `SimpleDocTemplate` for document layout
- Handle images with `ReportLabImage`

### Word/PPTX Generation

- Use python-docx for .docx files
- Use python-pptx for .pptx files
- Support template-based generation
- Dynamic font sizing for content

## Docker Guidelines

### Environment Variables

**Required for all deployments:**
- `FILE_EXPORT_BASE_URL`: URL to file export server (default: `http://localhost:9003/files`)
- `FILE_EXPORT_DIR`: Output directory (default: `PYTHONPATH/output`)

**Optional:**
- `PERSISTENT_FILES`: Keep files after download (default: `false`)
- `FILES_DELAY`: Cleanup delay in minutes (default: `60`)
- `LOG_LEVEL`: DEBUG, INFO, WARNING, ERROR, CRITICAL (default: `INFO`)
- `IMAGE_SOURCE`: "pexels", "unsplash", or "local_sd" (default: `unsplash`)

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

1. Create `_create_{ext}` function in `file_export_mcp.py`
2. Implement error handling with try/except
3. Use `_generate_unique_folder()` for output directory
4. Return `{"url": ..., "path": ...}` dict
5. Update `_public_url()` if URL construction is custom

### Modifying Templates

1. Place templates in `LLM_Export/templates/`
2. Template names: `Default_Template.{ext}`
3. Templates are auto-loaded on startup
4. Gracefully handle template load failures

## Performance Considerations

- Image downloads: 30-second timeout
- File cleanup runs in background threads
- Use BytesIO for in-memory image processing
- Generate unique folder names with UUID + timestamp

## Security

- JWT tokens for OWUI authentication (optional)
- MCPO API key for server authentication (optional)
- Never commit API keys or secrets
- Use environment variables for sensitive data

## Documentation

- User docs in `Documentation/` directory
- Code comments for complex logic
- Log messages for debugging context

## Notes for Agentic Agents

1. **Always test changes manually** - no automated tests exist
2. **Respect existing patterns** - follow established conventions exactly
3. **Use environment variables** - never hardcode configuration
4. **Log appropriately** - add debug logs for new functionality
5. **Handle edge cases** - empty data, missing files, network errors
6. **Docker-first** - most deployments use Docker Compose