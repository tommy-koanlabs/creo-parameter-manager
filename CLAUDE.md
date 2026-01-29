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

## XML Structure (Critical)

Creo exports parameters in **flat sequential order** with no grouping tags. Parameters are grouped by CAD object using sort order only:

```
CAGE_CODE → DESCRIPTION_1 → DESCRIPTION_2 → PART_NUMBER → PTC_WM_NAME
```

Each group of 5 consecutive `<Parameter>` elements belongs to one CAD object. The `PTC_WM_NAME` value identifies which object owns the preceding 4 parameters.

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

When displayed in the workbook, parameters are reordered to:
```
PTC_WM_NAME | CAGE_CODE | PART_NUMBER | DESCRIPTION_1 | DESCRIPTION_2
```

- Column A (PTC_WM_NAME): Read-only/locked — identifies the CAD object
- All cells: Formatted as text

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
- **Optional field:** DESCRIPTION_2 (allowed blank)
- **Blank handling:** Use `<Value></Value>` format
- **Row count:** Must match original XML parameter count
- **PTC_WM_NAME:** Must match original values (verify despite column lock)

## VBA Development Notes

- XML processing via MSXML2.DOMDocument
- ListBox controls: Evaluate Form vs ActiveX based on file system access needs
- File listing may require manual refresh button for detecting new external files
