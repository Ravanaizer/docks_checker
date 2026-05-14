from docx.oxml.ns import qn


def _is_list_paragraph(para) -> bool:
    """Checks if a paragraph is a bulleted or numbered list item."""
    pPr = para._element.find(qn("w:pPr"))
    if pPr is not None:
        numPr = pPr.find(qn("w:numPr"))
        return numPr is not None
    return False


def _get_table_font_size(table) -> float:
    """Returns the font size of the first table cell with explicit formatting."""
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    if run.font.size is not None and run.font.size.pt > 0:
                        return run.font.size.pt
    return 12.0


def _get_table_font_name(doc, table) -> str:
    """Returns the font name of the first non-empty table cell with explicit formatting."""
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    if run.text.strip():
                        font_name = _get_effective_font_name(doc.doc, run)
                        if font_name:
                            return font_name
    return "Calibri"


THEME_FONT_MAP = {
    "minorHAnsi": "Calibri",
    "minorAscii": "Calibri",
    "majorHAnsi": "Cambria",
    "majorAscii": "Cambria",
    "minorEastAsia": "Calibri",
    "minorCs": "Calibri",
    "majorEastAsia": "Cambria",
    "majorCs": "Cambria",
}


def _get_font_from_xml(run) -> str | None:
    """Directly extracts the font name from <w:rFonts> if the API returns None."""
    rPr = run._element.find(qn("w:rPr"))
    if rPr is not None:
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is not None:
            for attr in (qn("w:hAnsi"), qn("w:ascii"), qn("w:cs"), qn("w:eastAsia")):
                val = rFonts.get(attr)
                if val:
                    return val
            for attr in [
                qn("w:hAnsiTheme"),
                qn("w:asciiTheme"),
                qn("w:eastAsiaTheme"),
                qn("w:csTheme"),
            ]:
                theme_name = rFonts.get(attr)
                if theme_name:
                    return THEME_FONT_MAP.get(theme_name, f"Theme:{theme_name}")
    return None


def _get_effective_font_name(document, run) -> str | None:
    # 1. python-docx API (fast check)
    if run.font.name:
        return run.font.name

    # 2. Direct XML parsing of <w:rFonts>
    # Word sometimes leaves run.font.name empty but specifies the font in the XML
    val = _get_font_from_xml(run)
    if val:
        return val

    # 3. Character Style (applied to the specific run/selection)
    if run.style is not None:
        try:
            if run.style.font.name:
                return run.style.font.name
        except Exception:
            pass

    # 4. Paragraph Style
    try:
        if run.paragraph.style.font.name:
            return run.paragraph.style.font.name
    except Exception:
        pass

    # 5. Global "Normal" style
    try:
        if document.styles["Normal"].font.name:
            return document.styles["Normal"].font.name
    except Exception:
        pass

    # 6. Document settings (settings.xml → w:defaultFonts)
    try:
        settings = document.doc.settings.element
        default_fonts = settings.find(qn("w:defaultFonts"))
        if default_fonts is not None:
            return default_fonts.get(qn("w:hAnsi")) or default_fonts.get(qn("w:ascii"))
    except Exception:
        pass

    # 7. If all None → Word falls back to the Document Theme font (usually Calibri or Times New Roman)
    # Return None so the subsequent check correctly flags it as "not Arial"
    return None


def _get_effective_font_size(doc, run) -> float:
    """Reliably retrieves the font size with cascading inheritance."""
    if run.font.size is not None:
        return run.font.size.pt
    try:
        para_size = run._parent.style.font.size
        if para_size is not None:
            return para_size.pt
    except Exception:
        pass
    try:
        normal_size = doc.styles["Normal"].font.size
        if normal_size is not None:
            return normal_size.pt
    except Exception:
        pass
    return 12.0


def _get_empty_paragraph_font_size(doc, para) -> float:
    """
    Returns the font size from the style
    """
    try:
        if para.style and para.style.font.size:
            return para.style.font.size.pt
    except Exception:
        pass

    try:
        normal_style = doc.styles.get("Normal")
        if normal_style and normal_style.font.size:
            return normal_style.font.size.pt
    except Exception:
        pass

    return 12.0


def _get_paragraph_font_size(doc, para) -> float:
    """
    Extracts font size directly from paragraph XML structure.
    Critical for empty paragraphs that lack runs. Checks:
    1. <w:pPr> -> <w:rPr> -> <w:sz> (stored in half-points)
    2. <w:szCs> (complex scripts fallback)
    3. Paragraph style -> Normal style -> 12.0pt default
    """
    pPr = para._element.find(qn("w:pPr"))
    if pPr is not None:
        rPr = pPr.find(qn("w:rPr"))
        if rPr is not None:
            # Check standard size
            sz = rPr.find(qn("w:sz"))
            if sz is not None:
                try:
                    val = sz.get(qn("w:val"))
                    if val is not None:
                        return int(val) / 2.0  # Convert half-points to points
                except (ValueError, TypeError):
                    pass

            # Fallback to complex script size
            sz_cs = rPr.find(qn("w:szCs"))
            if sz_cs is not None:
                try:
                    val = sz_cs.get(qn("w:val"))
                    if val is not None:
                        return int(val) / 2.0
                except (ValueError, TypeError):
                    pass

    # Fallback to style inheritance
    try:
        if para.style and para.style.font.size:
            return para.style.font.size.pt
    except Exception:
        pass

    try:
        normal_style = doc.styles.get("Normal")
        if normal_style and normal_style.font.size:
            return normal_style.font.size.pt
    except Exception:
        pass

    return 12.0
