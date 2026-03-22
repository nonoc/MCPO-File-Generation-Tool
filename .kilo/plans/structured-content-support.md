# XLSX Export Structured Content Support Plan

## Problem Statement

The XLSX exporter currently only exports table data from structured content. When input like:
```python
[
    {"type": "title", "text": "Habit Tracker"},
    {"type": "paragraph", "text": "Tracking habits"},
    {"type": "table", "data": [["Habit", "Benefit", "Frequency"], ["Exercise", "Energy", "Daily"]]}
]
```
is provided, only the table is exported. The title and paragraph content is lost.

## Expected Behavior

Structured content should be exported with:
1. **Title** written first (row 1, bold)
2. **Paragraphs** following (each to new row)
3. **List items** rendered (each to new row)
4. **Table** written last as grid (columns consecutive: A, B, C, ...)

## Current Behavior

Only the table data is extracted and exported. Title and paragraphs are completely lost.

## Root Cause Analysis

### Location: `LLM_Export/exporters/xlsx_exporter.py`

**Function: `_preprocess_excel_content()` (lines 79-139)**

For structured blocks (list of dicts), the function:
1. Calls `_find_table_in_blocks()` which recursively searches for table blocks
2. Returns only the table data as a 2D list
3. **Loses all other content** (title, paragraphs, list items)

**Example:**
```python
structured = [{"type": "title", "text": "Habit Tracker"}, {"type": "table", "data": [...]}]
result = _preprocess_excel_content(structured)
# result = [["Habit", "Benefit", "Frequency"], ...]  # Title is lost!
```

### Location: `LLM_Export/exporters/xlsx_exporter.py`

**Function: `_create_excel()` (lines 170-181)**

After preprocessing, raw matrix detection:
```python
is_raw_matrix = (
    isinstance(data, list)
    and data
    and isinstance(data[0], list)
    and all(isinstance(cell, (str, int, float, bool, type(None))) for row in data[:5] for cell in row[:10])
)
```

For structured content:
1. `_preprocess_excel_content()` returns 2D list of primitives
2. Raw matrix detection sees this and sets `is_raw_matrix = True`
3. Clean workbook is created (template skipped)
4. Only table is exported

## Solution Strategy

### Option A: Full Structured Content Rendering (RECOMMENDED)

**File:** `LLM_Export/exporters/xlsx_exporter.py`

**Changes:**

1. **Add structured content detection** before preprocessing:
   ```python
   # Detect if input is structured blocks (list of dicts)
   is_structured = (
       isinstance(data, list)
       and data
       and isinstance(data[0], dict)
   )
   ```

2. **For structured content:**
   - Create clean workbook (skip template)
   - Render each block type in order:
     - Title: Write to first row, bold
     - Paragraphs: Write each to new row
     - List items: Write each to new row
     - Table: Write as grid starting after other content

3. **For raw 2D matrix:**
   - Keep existing behavior (clean workbook, matrix at A1)

4. **Simplified template handling:**
   - Never use template for structured content (clean spreadsheet mode)
   - Template is only for report-style layouts (if needed)

### Option B: Skip Template for Structured Content

If structured content is detected:
1. Skip template (create fresh workbook)
2. Extract and export only table data
3. Write table at A1 (columns consecutive)

## Implementation Plan

### Phase 1: Add Structured Content Detection

**File:** `LLM_Export/exporters/xlsx_exporter.py`

Add detection before preprocessing:
```python
# Detect structured content (list of dicts with type field)
is_structured = (
    isinstance(data, list)
    and data
    and isinstance(data[0], dict)
)

# Detect raw 2D matrix (list of lists of primitives)
is_raw_matrix = (
    isinstance(data, list)
    and data
    and isinstance(data[0], list)
    and all(isinstance(cell, (str, int, float, bool, type(None))) 
            for row in data[:5] for cell in row[:10])
)
```

### Phase 2: Add Structured Content Rendering

**File:** `LLM_Export/exporters/xlsx_exporter.py`

Add rendering function for structured blocks:
```python
def _render_structured_content_to_excel(ws, data, start_row=1, start_col=1):
    """Render structured blocks to Excel worksheet."""
    current_row = start_row
    
    for block in data:
        block_type = (block.get("type") or "").strip().lower()
        
        if block_type == "title":
            # Write title to first row, bold
            title_text = block.get("text") or block.get("title") or ""
            cell = ws.cell(row=current_row, column=start_col)
            cell.value = title_text
            cell.font = Font(bold=True)
            current_row += 1
            
        elif block_type == "paragraph":
            # Write paragraph to new row
            text = block.get("text") or ""
            cell = ws.cell(row=current_row, column=start_col)
            cell.value = text
            current_row += 1
            
        elif block_type in ("bullet", "list"):
            # Write list items
            items = block.get("items") or []
            for item in items:
                text = item.get("text") if isinstance(item, dict) else item
                cell = ws.cell(row=current_row, column=start_col)
                cell.value = text
                current_row += 1
                
        elif block_type == "table":
            # Write table as grid
            table_data = _normalize_table_rows(
                block.get("data") or block.get("rows") or block.get("cells") or []
            )
            for row_idx, row in enumerate(table_data):
                for col_idx, cell_value in enumerate(row):
                    ws.cell(
                        row=current_row + row_idx,
                        column=start_col + col_idx
                    ).value = cell_value
            current_row += len(table_data)
    
    return current_row
```

### Phase 3: Update Write Logic

**File:** `LLM_Export/exporters/xlsx_exporter.py`

Update write logic to handle both structured and raw matrix:
```python
if is_raw_matrix:
    # Existing raw matrix logic (clean workbook, matrix at A1)
    wb = Workbook()
    ws = wb.active
    # ... existing matrix write logic ...
else:
    # New structured content logic
    wb = Workbook()
    ws = wb.active
    _render_structured_content_to_excel(ws, data)
```

## Files to Modify

1. **`LLM_Export/exporters/xlsx_exporter.py`**
   - Add structured content detection
   - Add `_render_structured_content_to_excel()` function
   - Update write logic to handle structured content

## Regression Risks

1. **Existing raw 2D matrix users**: No change - existing behavior preserved
2. **Existing structured content users**: Behavior changes from "only table" to "full structured content"

## Test Cases

1. **Raw 2D Matrix:**
   ```python
   data = [["A", "B"], ["1", "2"]]
   # Expected: Clean workbook, data at A1:B2
   ```

2. **Structured Content with Title:**
   ```python
   data = [{"type": "title", "text": "Report"}, {"type": "table", "data": [...]}]
   # Expected: Title at row 1, table at row 2+
   ```

3. **Structured Content with Paragraph:**
   ```python
   data = [{"type": "paragraph", "text": "Description"}, {"type": "table", "data": [...]}]
   # Expected: Paragraph at row 1, table at row 2+
   ```

4. **Structured Content with All Block Types:**
   ```python
   data = [
       {"type": "title", "text": "Report"},
       {"type": "paragraph", "text": "Description"},
       {"type": "bullet", "text": "Item 1"},
       {"type": "table", "data": [...]}
   ]
   # Expected: All blocks rendered in order
   ```

## Implementation Checklist

- [ ] Add structured content detection
- [ ] Add `_render_structured_content_to_excel()` function
- [ ] Update write logic to handle structured content
- [ ] Test with raw 2D matrix (regression)
- [ ] Test with structured content (new feature)
- [ ] Run all existing tests