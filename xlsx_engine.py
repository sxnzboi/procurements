"""
OpenXML XLSX Reader and Writer using Python standard library (zipfile, xml.etree.ElementTree).
Requires zero third-party dependencies. Generates 100% compliant .xlsx files readable by
Microsoft Excel, Google Sheets, LibreOffice Calc, Apple Numbers, and pandas/openpyxl.
Supports both inline strings (inlineStr) and sharedStrings for bulletproof compatibility.
"""

import os
import io
import re
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape as xml_escape

def _col_name(n):
    """Convert 0-indexed column number to Excel column letters (0 -> 'A', 25 -> 'Z', 26 -> 'AA')."""
    result = []
    n += 1
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result.append(chr(65 + rem))
    return "".join(reversed(result))

def _col_index(col_str):
    """Convert Excel column letters to 0-indexed column number ('A' -> 0, 'Z' -> 25, 'AA' -> 26)."""
    result = 0
    for char in col_str.upper():
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result - 1

def write_xlsx(filepath, sheets_data):
    """
    Write an Excel .xlsx workbook.
    sheets_data: dict of {sheet_name: [ [cell, cell, ...], [row2...], ... ]}
    or list of rows (if single sheet default named 'Sheet1').
    Cells can be int, float, str, or None.
    Uses inlineStr for strings so text values NEVER show up as numbers in Excel.
    """
    if isinstance(sheets_data, list):
        sheets_data = {"Sheet1": sheets_data}

    # Collect unique shared strings for full spec compliance
    shared_strings = []
    string_to_idx = {}
    
    def get_string_idx(s):
        s_str = str(s)
        if s_str not in string_to_idx:
            idx = len(shared_strings)
            string_to_idx[s_str] = idx
            shared_strings.append(s_str)
            return idx
        return string_to_idx[s_str]

    sheet_xmls = {}
    sheet_names = list(sheets_data.keys())

    for sheet_idx, (name, rows) in enumerate(sheets_data.items(), start=1):
        # Build worksheet XML
        ws_lines = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
            '  <sheetViews>',
            '    <sheetView tabSelected="' + ('1' if sheet_idx == 1 else '0') + '" workbookViewId="0">',
            '    </sheetView>',
            '  </sheetViews>',
            '  <sheetFormatPr defaultRowHeight="18"/>',
            '  <cols>'
        ]
        
        # Calculate max column width
        num_cols = max((len(r) for r in rows), default=0)
        col_widths = [12] * num_cols
        for row in rows:
            for c_idx, cell in enumerate(row):
                if cell is not None:
                    col_widths[c_idx] = min(max(col_widths[c_idx], len(str(cell)) + 4), 50)
        
        for c_idx, width in enumerate(col_widths, start=1):
            ws_lines.append(f'    <col min="{c_idx}" max="{c_idx}" width="{width}" customWidth="1"/>')
        ws_lines.append('  </cols>')
        ws_lines.append('  <sheetData>')

        for r_idx, row in enumerate(rows, start=1):
            is_header = (r_idx == 1)
            row_ht = '22' if is_header else '19'
            ws_lines.append(f'    <row r="{r_idx}" ht="{row_ht}" customHeight="1">')
            for c_idx, val in enumerate(row):
                col_letter = _col_name(c_idx)
                cell_ref = f"{col_letter}{r_idx}"
                s_attr = ' s="1"' if is_header else ''
                
                if val is None or val == "":
                    continue
                elif isinstance(val, (int, float)):
                    num_val = f"{val:.2f}" if isinstance(val, float) else str(val)
                    ws_lines.append(f'      <c r="{cell_ref}"{s_attr}><v>{num_val}</v></c>')
                elif isinstance(val, bool):
                    ws_lines.append(f'      <c r="{cell_ref}" t="b"{s_attr}><v>{1 if val else 0}</v></c>')
                else:
                    # Use inlineStr directly so Excel NEVER displays raw numeric string indexes!
                    # Also record in shared_strings for full spec completeness
                    get_string_idx(val)
                    escaped_text = xml_escape(str(val))
                    ws_lines.append(f'      <c r="{cell_ref}" t="inlineStr"{s_attr}><is><t>{escaped_text}</t></is></c>')
            ws_lines.append('    </row>')

        ws_lines.append('  </sheetData>')
        ws_lines.append('</worksheet>')
        sheet_xmls[sheet_idx] = "\n".join(ws_lines)

    # Build sharedStrings.xml
    sst_lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{len(shared_strings)}" uniqueCount="{len(shared_strings)}">'
    ]
    for s in shared_strings:
        sst_lines.append(f'  <si><t>{xml_escape(s)}</t></si>')
    sst_lines.append('</sst>')
    shared_strings_xml = "\n".join(sst_lines)

    # Build styles.xml
    styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font>
      <sz val="11"/>
      <name val="Calibri"/>
      <family val="2"/>
    </font>
    <font>
      <b/>
      <sz val="11"/>
      <color rgb="FFFFFFFF"/>
      <name val="Calibri"/>
      <family val="2"/>
    </font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FF1E3A8A"/>
        <bgColor indexed="64"/>
      </patternFill>
    </fill>
  </fills>
  <borders count="2">
    <border>
      <left/><right/><top/><bottom/><diagonal/>
    </border>
    <border>
      <left style="thin"><color rgb="FFD1D5DB"/></left>
      <right style="thin"><color rgb="FFD1D5DB"/></right>
      <top style="thin"><color rgb="FFD1D5DB"/></top>
      <bottom style="thin"><color rgb="FFD1D5DB"/></bottom>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center"/>
    </xf>
  </cellXfs>
</styleSheet>"""

    # Build workbook.xml
    wb_lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '  <sheets>'
    ]
    for idx, name in enumerate(sheet_names, start=1):
        wb_lines.append(f'    <sheet name="{xml_escape(name)}" sheetId="{idx}" r:id="rId{idx}"/>')
    wb_lines.append('  </sheets>')
    wb_lines.append('</workbook>')
    workbook_xml = "\n".join(wb_lines)

    # Build workbook.xml.rels
    wb_rels_lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    ]
    for idx in range(1, len(sheet_names) + 1):
        wb_rels_lines.append(f'  <Relationship Id="rId{idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{idx}.xml"/>')
    styles_rid = len(sheet_names) + 1
    shared_str_rid = len(sheet_names) + 2
    wb_rels_lines.append(f'  <Relationship Id="rId{styles_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
    wb_rels_lines.append(f'  <Relationship Id="rId{shared_str_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>')
    wb_rels_lines.append('</Relationships>')
    wb_rels_xml = "\n".join(wb_rels_lines)

    # Build [Content_Types].xml (Note: sharedStrings+xml with 's' at the end is REQUIRED by Excel!)
    ct_lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '  <Default Extension="xml" ContentType="application/xml"/>',
        '  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
        '  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
    ]
    for idx in range(1, len(sheet_names) + 1):
        ct_lines.append(f'  <Override PartName="/xl/worksheets/sheet{idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    ct_lines.append('</Types>')
    content_types_xml = "\n".join(ct_lines)

    # Build _rels/.rels
    root_rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

    # Write out ZIP file
    parent_dir = os.path.dirname(filepath)
    if parent_dir:
        os.makedirs(parent_dir, exist_ok=True)

    with zipfile.ZipFile(filepath, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml)
        zf.writestr("_rels/.rels", root_rels_xml)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", wb_rels_xml)
        zf.writestr("xl/styles.xml", styles_xml)
        zf.writestr("xl/sharedStrings.xml", shared_strings_xml)
        for idx, xml_content in sheet_xmls.items():
            zf.writestr(f"xl/worksheets/sheet{idx}.xml", xml_content)

    return filepath


def read_xlsx(filepath, sheet_name=None):
    """
    Read an Excel .xlsx workbook.
    Returns: list of rows (each row is a list of cell values).
    If sheet_name is None, reads the first worksheet.
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"File not found: {filepath}")

    with zipfile.ZipFile(filepath, "r") as zf:
        # Read shared strings if present
        shared_strings = []
        if "xl/sharedStrings.xml" in zf.namelist():
            sst_root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in sst_root.findall("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si"):
                text_parts = []
                for t in si.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"):
                    if t.text:
                        text_parts.append(t.text)
                shared_strings.append("".join(text_parts))

        # Read workbook.xml to locate sheet
        wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
        sheets_elem = wb_root.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheets")
        
        target_sheet_filename = "worksheets/sheet1.xml"
        if sheets_elem is not None:
            first_sheet = True
            for sheet in sheets_elem.findall("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet"):
                s_name = sheet.attrib.get("name")
                r_id = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                
                if (sheet_name and s_name.lower() == sheet_name.lower()) or (sheet_name is None and first_sheet):
                    wb_rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
                    for rel in wb_rels_root.findall("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                        if rel.attrib.get("Id") == r_id:
                            target_sheet_filename = rel.attrib.get("Target")
                            if not target_sheet_filename.startswith("worksheets/"):
                                target_sheet_filename = "worksheets/" + os.path.basename(target_sheet_filename)
                            break
                    break
                first_sheet = False

        full_path = f"xl/{target_sheet_filename}" if not target_sheet_filename.startswith("xl/") else target_sheet_filename
        if full_path not in zf.namelist():
            candidates = [name for name in zf.namelist() if name.startswith("xl/worksheets/sheet") and name.endswith(".xml")]
            if candidates:
                full_path = sorted(candidates)[0]
            else:
                raise ValueError("No worksheets found in XLSX file.")

        ws_root = ET.fromstring(zf.read(full_path))
        sheet_data = ws_root.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheetData")

        rows_data = []
        if sheet_data is not None:
            for row in sheet_data.findall("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row"):
                row_cells = {}
                max_c_idx = -1
                for c in row.findall("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c"):
                    r_ref = c.attrib.get("r", "")
                    match = re.match(r"([A-Z]+)(\d+)", r_ref)
                    if match:
                        col_str, _ = match.groups()
                        c_idx = _col_index(col_str)
                    else:
                        c_idx = max_c_idx + 1

                    if c_idx > max_c_idx:
                        max_c_idx = c_idx

                    c_type = c.attrib.get("t")
                    v_elem = c.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v")
                    val = None

                    if c_type == "s" and v_elem is not None and v_elem.text:
                        s_idx = int(v_elem.text)
                        val = shared_strings[s_idx] if s_idx < len(shared_strings) else ""
                    elif c_type == "inlineStr":
                        is_elem = c.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}is")
                        if is_elem is not None:
                            val = "".join(is_elem.itertext())
                    elif c_type == "b" and v_elem is not None:
                        val = (v_elem.text == "1")
                    elif v_elem is not None and v_elem.text:
                        raw_text = v_elem.text
                        try:
                            if "." in raw_text:
                                val = float(raw_text)
                            else:
                                val = int(raw_text)
                        except ValueError:
                            val = raw_text

                    row_cells[c_idx] = val

                if max_c_idx >= 0:
                    row_list = [row_cells.get(i, None) for i in range(max_c_idx + 1)]
                    rows_data.append(row_list)
                else:
                    rows_data.append([])

        return rows_data
