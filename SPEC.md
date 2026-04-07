# WPS MCP Server Specification

## Overview
MCP Server for WPS Office operations - enables AI agents to create, edit, and manage WPS documents, spreadsheets, and presentations.

## Features

### Document Tools (Word)
1. `create_document` - Create a new WPS Word document
2. `add_heading` - Add heading to document
3. `add_paragraph` - Add paragraph text
4. `read_document` - Read document content

### Spreadsheet Tools (Excel)
5. `create_spreadsheet` - Create a new WPS Excel file
6. `write_cell` - Write data to a cell
7. `read_cell` - Read cell data
8. `add_formula` - Add formula to cell

### Presentation Tools (PowerPoint)
9. `create_presentation` - Create a new WPS PPT file
10. `add_slide` - Add a new slide
11. `add_text_to_slide` - Add text to slide
12. `set_slide_layout` - Set slide layout

## Technical Stack
- Python 3.10+
- python-docx (WPS/Word documents)
- openpyxl (WPS/Excel spreadsheets)
- python-pptx (WPS/PowerPoint presentations)
- FastMCP

## Note
WPS Office uses the same file formats as Microsoft Office (.docx, .xlsx, .pptx).
This MCP server manipulates these standard formats, making it compatible with both WPS and MS Office.

## Status
Implementation started - 2026-04-07
