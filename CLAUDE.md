# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Tools for managing Creo Parametric model parameters via a macro-enabled Excel workbook (VBA). Converts Creo parameter XML files to editable spreadsheets and back to importable XML.

**Terminology:** Parts (.prt) and assemblies (.asm) are referred to as "CAD objects."

## Repository Files

| File | Purpose |
|------|---------|
| `param_manager.xlsm` | Main macro-enabled workbook with VBA code for import/export |
| `example.xml` | Sample Creo parameter export for testing |
| `rp_config.xml` | Creo parameter dialog filter configuration |
| `vba/modParamManager.bas` | Main VBA module: import, export, list refresh logic |
| `vba/ThisWorkbook.cls` | Workbook events: auto-refresh on open/activate |
| `vba/Sheet1_Manager.cls` | Sheet events: button click handlers (optional) |
| `vba/INSTALL.txt` | Installation instructions for VBA modules |

## XML Structure (Critical)

Creo exports parameters in **flat sequential order** with no grouping tags. Parameters are grouped by CAD object using sort order only. The field list is **dynamic** — detect the group size by finding the first duplicate parameter name:

```
CAGE_CODE → DESCRIPTION_1 → ... → PTC_WM_NAME → CAGE_CODE (duplicate = end of group)
```

Each group of N consecutive `<Parameter>` elements belongs to one CAD object. The `PTC_WM_NAME` value identifies which object owns the preceding parameters.

**This order must be preserved exactly when writing back to XML** — there is no other mechanism to associate parameters with their CAD objects.

Example parameter structure:
```xml
<CreoParamSet>
    <Parameter Name="CAGE_CODE">
        <DataType>String</DataType>
        <Value>0AEX9</Value>
        <Description>*Design Activity CAGE Code</Description>
    </Parameter>
    <!-- DESCRIPTION_1, DESCRIPTION_2, PART_NUMBER follow -->
    <Parameter Name="PTC_WM_NAME">
        <DataType>String</DataType>
        <Value>ssp-j12ttf_brnch_tee_12.prt</Value>
        <Access>Locked</Access>
    </Parameter>
</CreoParamSet>
```

## Spreadsheet Column Order

Columns are ordered dynamically with a priority system:

1. **Priority fields** (in this order, if present): `PTC_WM_NAME`, `CAGE_CODE`, `PART_NUMBER`, `DESCRIPTION_1`, `DESCRIPTION_2`
2. **Additional fields**: Sorted alphabetically after priority fields (excludes priority fields already added)

Example: XML with fields `A`, `CAGE_CODE`, `E`, `DESCRIPTION_1`, `PTC_WM_NAME` becomes:
```
PTC_WM_NAME | CAGE_CODE | DESCRIPTION_1 | A | E
```

## Cell and Column Formatting

### Visual Structure
- **First column**: Bold text with grey fill (#C8C8C8) - identifies CAD objects
- **First row**: Bold text - column headers
- **Marker row** (row 2): Hidden row containing "F" (full field) or "P" (partial field) markers
- **Data rows**: Start at row 3, all cells have borders and white default fill
- **All cells**: Formatted as text to prevent Excel from misinterpreting values

### Color-Coded Field Status

**Standard fields** (PTC_WM_NAME, CAGE_CODE, PART_NUMBER, DESCRIPTION_1, DESCRIPTION_2):
- **Pastel green** (RGB 204, 255, 204): Filled - good status
- **Pastel red** (RGB 255, 204, 204): Blank - needs attention
- **Light grey** (RGB 240, 240, 240): Missing from object when present in others
- **Note**: PTC_WM_NAME column has dark grey fill and no conditional formatting (always present)

**Additional fields with full presence** (all objects have this parameter):
- **Pastel green** (RGB 204, 255, 204): Filled
- **Pastel red** (RGB 255, 204, 204): Blank

**Additional fields with partial presence** (some objects have this parameter, some don't):
- **Pastel blue** (RGB 204, 204, 255): Original data from XML
- **Light grey** (RGB 240, 240, 240): Parameter missing (empty cell)
- **Light yellow** (RGB 255, 255, 204): User-added data (automatically applied when user fills empty cell)

**Color saturation**: All pastel colors (green, red, blue, yellow) have matching saturation levels for visual consistency

### Locked Columns
- Any field with `<Access>Locked</Access>` in the XML is locked (read-only)
- Typically `PTC_WM_NAME` is locked
- Lock status is preserved during export

## Implementation Architecture

The `param_manager.xlsm` workbook uses:
- **Manager sheet** (first sheet): Contains readme, Import/Export buttons, and two list boxes
- **XML File ListBox**: Lists `.xml` files in the workbook's directory (newest first)
- **Sheet ListBox**: Lists data sheets (excluding Manager), newest first

**Import flow:** XML file → new worksheet named `<filename>-<yyyymmdd_hhmmss>.xml`

**Export flow:** Selected worksheet → XML file with same name as sheet

## Validation Rules

When exporting from spreadsheet back to XML:
- **Required fields:** CAGE_CODE, DESCRIPTION_1, PART_NUMBER, PTC_WM_NAME (warn if blank)
- **Optional fields:** DESCRIPTION_2 and any dynamically-detected fields (blanks allowed)
- **Blank handling:** Use `<Value></Value>` format
- **Row count:** Must match original XML parameter count
- **Locked fields:** Preserved from import — columns that were locked in Excel will have `<Access>Locked</Access>` in exported XML

## VBA Development Notes

- XML processing via MSXML2.DOMDocument60
- ListBox controls: ActiveX (`ListBox1` for XML files, `ListBox2` for sheets)
- **Refresh strategy:** Combined approach — `Workbook_Activate` event auto-refreshes XML file list, plus manual Refresh button for on-demand updates
- **Dynamic field detection:** Iterate XML parameter names until first duplicate to determine group size
- **Marker row system:** Hidden row 2 contains "F" or "P" to indicate full/partial field presence
- **Automatic color updates:** `Workbook_SheetChange` event detects user-added data in partial fields and applies yellow highlighting

### VBA Module Structure

| Module | Key Procedures |
|--------|----------------|
| `modParamManager` | `ImportXML`, `ExportXML`, `RefreshXMLFileList`, `RefreshSheetList`, `DetectLockedFields` |
| `ThisWorkbook` | `Workbook_Open`, `Workbook_Activate` events |

### Key Constants (in modParamManager)

```vba
PRIORITY_FIELDS = "PTC_WM_NAME,CAGE_CODE,PART_NUMBER,DESCRIPTION_1,DESCRIPTION_2"
REQUIRED_FIELDS = "CAGE_CODE,DESCRIPTION_1,PART_NUMBER,PTC_WM_NAME"
```
