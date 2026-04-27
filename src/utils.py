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
                        font_name = _get_effective_font_name(doc, run)
                        if font_name:
                            return font_name
    return "Calibri"


def _get_font_from_xml(run) -> str | None:
    """Directly extracts the font name from <w:rFonts> if the API returns None."""
    rPr = run._element.find(qn("w:rPr"))
    if rPr is not None:
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is not None:
            return rFonts.get(qn("w:hAnsi")) or rFonts.get(qn("w:ascii"))
    return None


def _get_effective_font_name(doc, run):
    """Reliably retrieves the font name: Run -> XML -> Paragraph Style -> Normal."""
    # 1. Explicit font in Run (via API or directly from XML)
    font_name = run.font.name or _get_font_from_xml(run)
    if font_name:
        return font_name
    # 2. Paragraph style
    try:
        style_name = run._parent.style.font.name
        if style_name:
            return style_name
    except Exception:
        pass
    # 3. Default document style
    try:
        normal_name = doc.styles["Normal"].font.name
        if normal_name:
            return normal_name
    except Exception:
        pass
    return "Calibri"


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


def _get_paragraph_font_size(para, doc) -> float:
    """Returns the font size from XML."""
    # 1. w:pPr -> w:rPr -> w:sz
    pPr = para._element.find(qn("w:pPr"))
    if pPr is not None:
        rPr = pPr.find(qn("w:rPr"))
        if rPr is not None:
            sz = rPr.find(qn("w:sz"))
            if sz is not None:
                try:
                    val = sz.get(qn("w:val"))
                    if val is not None:
                        return int(val) / 2.0
                except (ValueError, TypeError):
                    pass

            sz_cs = rPr.find(qn("w:szCs"))
            if sz_cs is not None:
                try:
                    val = sz_cs.get(qn("w:val"))
                    if val is not None:
                        return int(val) / 2.0
                except (ValueError, TypeError):
                    pass

    # 2. from style
    try:
        if para.style and para.style.font.size:
            return para.style.font.size.pt
    except Exception:
        pass

    # 3. from style Normal
    try:
        normal_style = doc.styles.get("Normal")
        if normal_style and normal_style.font.size:
            return normal_style.font.size.pt
    except Exception:
        pass

    return 12.0
