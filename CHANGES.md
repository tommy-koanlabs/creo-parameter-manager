# Changes Summary

## New Features

### Enhanced Conditional Formatting System (Refined)
**Feature**: Comprehensive color-coded formatting system for standard and additional parameters with automatic detection of user-added data.

**Color Scheme** (all pastel colors have matching saturation):

*Standard fields (PTC_WM_NAME, CAGE_CODE, PART_NUMBER, DESCRIPTION_1, DESCRIPTION_2):*
- Pastel green (RGB 204, 255, 204): Filled
- Pastel red (RGB 255, 204, 204): Blank
- Light grey (RGB 240, 240, 240): Missing from object
- Note: PTC_WM_NAME has dark grey fill, no conditional formatting (always present)

*Additional fields - full presence:*
- Pastel green (RGB 204, 255, 204): Filled
- Pastel red (RGB 255, 204, 204): Blank

*Additional fields - partial presence:*
- Pastel blue (RGB 204, 204, 255): Original data from XML
- Light grey (RGB 240, 240, 240): Missing (empty cell)
- Light yellow (RGB 255, 255, 204): User-added data (auto-applied)

**Import Formatting**:
- First column: Bold text, dark grey fill (RGB 200, 200, 200)
- First row: Bold headers, dark grey fill (RGB 200, 200, 200)
- All sheet cells: White fill
- All data cells: Bordered
- Marker row (row 2): Hidden, contains "F" (full) or "P" (partial) markers

**Implementation**:
- Added marker row system to track field presence type
- Rewrote `FormatDataSheet()` with comprehensive conditional formatting rules
- Added `Workbook_SheetChange` event handler in `ThisWorkbook.cls` to detect user additions
- When user fills empty cell in partial field, automatically changes from grey to yellow
- Added `InStrInArray()` helper function to check if field is standard

**Refinements**:
- Adjusted all pastel colors to matching saturation levels (51 units from white RGB 255,255,255)
- Removed dark red/bold formatting for missing standard fields (simplified to light grey)
- PTC_WM_NAME column now has dark grey fill and no conditional formatting
- First row (headers) now has dark grey fill
- White fill applied to entire sheet, not just data cells
- Standard field detection no longer includes locked fields check

**Technical Details**:
- Data rows now start at row 3 (row 1 = headers, row 2 = markers)
- Marker row contains "F" for full fields, "P" for partial fields
- Direct formatting (borders, first column) applied before conditional formatting
- Conditional formatting uses formulas like `=LEN(TRIM(cell))>0`
- Event handler disables events during processing to prevent recursion

---

## Bug Fixes

### 1. Fixed List Refresh "Subscript out of range" Error
**Problem**: RefreshXMLFileList and RefreshSheetList were throwing "Subscript out of range" errors on subsequent refreshes after the first successful load.

**Root Cause**: The bubble sort algorithms were manipulating VBA Collections during sorting by removing and adding items. When you remove items from a Collection, the indices shift, causing subsequent accesses to `xmlFiles(j-1)` or `dataSheets(j-1)` to be out of range.

**Solution**: Replaced Collection-based sorting with array-based sorting:
- Collect items into dynamic arrays instead of Collections
- Perform bubble sort on arrays (stable indices during swaps)
- Populate ListBox from sorted arrays

**Files Changed**:
- `RefreshXMLFileList()` in modParamManager.bas:117-203
- `RefreshSheetList()` in modParamManager.bas:205-266

### 2. Fixed Parameter Duplication Issue
**Problem**: Parameters were appearing duplicated - once in priority order, then again in alphabetical order.

**Root Cause**: In `OrderFieldsByPriority` function, priority fields were added to the `orderedFields` collection without a key (line 413). When `CollectionContains` later checked for these fields by key, it couldn't find them, causing all priority fields to be added again as "additional fields".

**Solution**: Changed line 413 from:
```vba
orderedFields.Add priorityArr(i)
```
to:
```vba
orderedFields.Add priorityArr(i), priorityArr(i)
```

This ensures fields are added with both value and key, allowing proper duplicate detection.

---

## New Features

### 2. Dynamic Locked Column Detection
**Previous Behavior**: Only `PTC_WM_NAME` was locked (hardcoded).

**New Behavior**: 
- During import, the code scans for `<Access>Locked</Access>` elements in the XML
- Any field with this attribute is locked and displayed with grey background
- All locked columns are protected, not just the first one

**Implementation**:
- Added `DetectLockedFields()` function to scan XML for locked fields
- Updated `FormatDataSheet()` to lock all detected locked columns
- Lock status is dynamically determined from XML, not hardcoded

### 3. Preserved Locked Status on Export
**Previous Behavior**: Only `PTC_WM_NAME` received `<Access>Locked</Access>` in exported XML (hardcoded).

**New Behavior**: 
- Export function reads which columns are locked in the Excel sheet
- All locked columns receive `<Access>Locked</Access>` in the exported XML
- Lock status is preserved across import/export cycle

**Implementation**:
- Updated `ExportXML()` to detect locked columns via `Columns(col).Locked`
- Changed conditional from `If fieldName = "PTC_WM_NAME"` to `If CollectionContains(lockedFields, fieldName)`

### 4. Conditional Formatting for Partial Field Presence
**New Feature**: When a field exists in some CAD objects but not others, cells are color-coded to show status.

**Behavior**:
- **Grey background**: Parameter missing from this object (empty cell)
- **Pale yellow background**: Parameter present in this object (either from original data or user-added)
- **No background**: Field exists in all objects (normal column)
- When user enters a value in a grey cell, it automatically turns pale yellow
- Visual status makes it easy to identify missing parameters and track additions

**Implementation**:
- Added logic in `FormatDataSheet()` to detect columns with empty cells (partial field presence)
- Applies two conditional formatting rules per affected column:
  - Rule 1 (Priority 1): `=LEN(TRIM(cell))>0` → Pale yellow (RGB 255, 255, 204)
  - Rule 2 (Priority 2): `=LEN(TRIM(cell))=0` → Light grey (RGB 240, 240, 240)
- Rules apply only to columns where some objects have the field and others don't

---

## Documentation Updates

Updated `CLAUDE.md` to reflect:
- Clarified that "additional fields" excludes priority fields already added
- Documented dynamic locked column behavior
- Documented conditional formatting for partial field presence (grey/yellow indicators)
- Updated validation rules to mention preserved lock status
- Added `DetectLockedFields` to VBA module structure table

## Edge Cases Handled

1. **Missing standard fields**: Code correctly handles situations where priority fields (CAGE_CODE, PART_NUMBER, etc.) are missing from the XML entirely
   - `OrderFieldsByPriority` only adds priority fields if they exist in `fieldNames`
   - No errors or blank columns for missing standard fields

2. **Partial field presence**: Handles parameters that exist in some objects but not others
   - Empty cells for missing parameters are left blank during import
   - Conditional formatting provides visual indicators (grey = missing, yellow = present/added)
   - Users can add parameters to individual objects by filling in grey cells

---

## Files Modified

- `vba/modParamManager.bas`: Core logic changes
- `CLAUDE.md`: Documentation updates

## Compatibility

All changes are backward compatible:
- Existing XML files work unchanged
- `PTC_WM_NAME` will still be locked (it has `<Access>Locked</Access>` in example.xml)
- Export behavior preserves lock status correctly
