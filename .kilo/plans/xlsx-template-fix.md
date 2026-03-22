# XLSX Export Template Fix Plan

## Problem Statement

The XLSX exporter receives valid 2D matrix input (e.g., `[["Header1", "Header2"], ["Val1", "Val2"]]`) but the generated workbook behaves like a template/report layout instead of a clean spreadsheet grid.

## Key Observations

### 1. Template Loading Behavior
```python
# xlsx_exporter.py lines 170-181
if XLSX_TEMPLATE and XLSX_TEMPLATE_PATH:
    try:
        log.debug("Loading XLSX template...")
        wb = load_workbook(XLSX_TEMPLATE_PATH)
    except Exception as e:
        log.warning(f"Failed to load XLSX template: {e}")
        wb = Workbook()
else:
    wb = Workbook()
```

**Current Issue:**
- Template is loaded whenever `XLSX_TEMPLATE_PATH` is configured
- Template has pre-defined layout (title cell at B3, auto-filter at B5, empty rows)
- This layout interferes with raw 2D matrix data

### 2. Template Layout Analysis
```
Row 1: [empty, empty, empty, ...]
Row 2: [empty, empty, empty, ...]
Row 3: [empty, "Title", empty, ...]  ← Title placeholder
Row 4: [empty, empty, empty, ...]
Row 5: [empty, "Title", "Title", ...] ← Auto-filter starts here
```

### 3. Data Flow for 2D Matrix
```
User Input: [["Healthy Workplace Habits Report"], ["Title", "Value"], ...]
  ↓
_preprocess_excel_content() detects 2D list, returns as-is
  ↓
_create_excel() loads XLSX_TEMPLATE (if configured)
  ↓
Data written starting at start_row=1, start_col=1 (or auto_filter position)
  ↓
Template layout overwrites/interferes with matrix structure
```

## Root Cause

The XLSX exporter **always** loads the template when `XLSX_TEMPLATE_PATH` is set, regardless of whether the content is:
- A clean 2D matrix (spreadsheet mode - should NOT use template)
- Structured blocks (report mode - SHOULD use template)

## Solution Strategy

### Mode Detection
When content is already a valid 2D matrix:
- Detect: `isinstance(data, list) and data and isinstance(data[0], list)`
- Action: Skip template, create fresh workbook

When content is structured blocks:
- Detect: `isinstance(data, list) and data and isinstance(data[0], dict)`
- Action: Use template for report-style formatting

### Implementation Plan

**File:** `LLM_Export/exporters/xlsx_exporter.py`

**Changes:**

1. **After preprocessing (line 158)**, add mode detection:
```python
# Preprocess content to normalize to 2D list format
data = _preprocess_excel_content(data)

# Detect if this is raw 2D matrix data
is_raw_matrix = (
    isinstance(data, list) 
    and data 
    and isinstance(data[0], list)
    and all(isinstance(cell, (str, int, float, bool, type(None))) 
            for row in data[:5] for cell in row[:10])  # Sample check
)

# Load template or create new workbook
if is_raw_matrix:
    # Clean spreadsheet mode - always fresh workbook
    log.debug("Raw 2D matrix detected, creating clean workbook")
    wb = Workbook()
else:
    # Template/report mode - use template if available
    if XLSX_TEMPLATE and XLSX_TEMPLATE_PATH:
        try:
            log.debug("Loading XLSX template...")
            wb = load_workbook(XLSX_TEMPLATE_PATH)
        except Exception as e:
            log.warning(f"Failed to load XLSX template: {e}")
            wb = Workbook()
    else:
        log.debug("No XLSX template available, creating new workbook")
        wb = Workbook()
```

2. **Adjust title handling** (lines 185-198):
   - Skip title cell replacement for raw matrices (no "title" placeholder in clean data)
   - Or: Set sheet title only, don't write title into cell A1

3. **Simplify write loop for raw matrices** (lines 222-251):
   - Start at A1 (start_row=1, start_col=1)
   - No auto_filter detection needed
   - Write directly without template border logic

## Implementation Options

### Option A: Template Skipping (RECOMMENDED)
- Detect 2D matrix → skip template entirely
- Pros: Clean, minimal changes, correct behavior for spreadsheet data
- Cons: May affect users expecting template behavior for all XLSX

### Option B: Explicit Flag
- Add `use_template` parameter to `_create_excel()`
- Default: `True` (backward compatible)
- Pros: Explicit control, no ambiguity
- Cons: Requires caller changes

## Files to Modify

1. **`LLM_Export/exporters/xlsx_exporter.py`** - Add mode detection and conditional template usage

## Regression Risks

1. **Users relying on template for XLSX reports** - May need to migrate to structured blocks or use a different format
2. **Existing templates with specific formatting** - Will be ignored for 2D matrix input

## Mitigation

- Document the behavior: "Raw 2D matrices skip template; use structured blocks for report-style XLSX"
- Template is still used for non-matrix content (structured blocks, markdown strings with tables)
- Backward compatible for non-matrix content

## Test Cases

1. **2D Matrix Input:**
   ```python
   data = [["Header1", "Header2"], ["Val1", "Val2"]]
   # Expected: Clean workbook, data at A1:B2
   ```

2. **Structured Blocks Input:**
   ```python
   data = [{"type": "table", "data": [...]}]
   # Expected: Template applied, formatted output
   ```

3. **Markdown String Input:**
   ```python
   data = "| A | B |\n|---|---|\n|1|2|"
   # Expected: Template applied after parsing
   ```

## Secondary Consideration

**Title parameter handling** (line 207 in create_file):
```python
result = _create_excel(content, filename, folder_path=folder_path)
# Title is NOT passed to _create_excel()
```

Should be:
```python
result = _create_excel(content, filename, folder_path=folder_path, title=title)
```

This is a minor improvement but not required for the core fix.