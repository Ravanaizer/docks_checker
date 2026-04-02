from docx.oxml.ns import qn


def _is_list_paragraph(para) -> bool:
    """Определяет, является ли абзац элементом маркированного/нумерованного списка."""
    pPr = para._element.find(qn("w:pPr"))
    if pPr is not None:
        numPr = pPr.find(qn("w:numPr"))
        return numPr is not None
    return False


# def _get_effective_font_size(document, run) -> float:
#     """Получает реальный размер шрифта с учётом наследования стилей."""
#     if run.font.size is not None:
#         return run.font.size.pt

#     # Проверяем стиль параграфа
#     paragraph = run._parent
#     if paragraph is not None:
#         try:
#             style_font = paragraph.style.font
#             if style_font and style_font.size:
#                 return style_font.size.pt
#         except Exception:
#             pass

#     # Проверяем стиль документа по умолчанию
#     try:
#         normal_style = document.doc.styles["Normal"].font
#         if normal_style and normal_style.size:
#             return normal_style.size.pt
#     except Exception:
#         pass

#     return 12.0


def _get_table_font_size(table) -> float:
    """Возвращает размер шрифта первой ячейки таблицы с явным форматированием."""
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    if run.font.size is not None and run.font.size.pt > 0:
                        return run.font.size.pt
    return 12.0


def _get_table_font_name(table) -> str:
    """Возвращает имя шрифта первой непустой ячейки таблицы с явным форматированием."""
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                # Проверяем, что в параграфе есть видимый текст
                if not paragraph.text.strip():
                    continue

                # Ищем run с текстом и заданным шрифтом
                for run in paragraph.runs:
                    if run.text.strip() and run.font.name:
                        return run.font.name
    return "Calibri"


# def _get_effective_font_name(doc, run):
#     """Возвращает имя_шрифта с учётом наследования стилей."""
#     font_name = None

#     if run.font.name:
#         font_name = run.font.name

#     # 2. Наследование от стиля параграфа
#     para = run._parent
#     if font_name is None:
#         try:
#             style = para.style
#             if style.font.name and font_name is None:
#                 font_name = style.font.name
#         except Exception:
#             pass

#     # 3. Наследование от стиля документа по умолчанию
#     if font_name is None:
#         try:
#             normal = doc.styles["Normal"]
#             if normal.font.name and font_name is None:
#                 font_name = normal.font.name
#         except Exception:
#             pass

#     # Дефолтные значения, если ничего не найдено
#     return font_name or "Calibri"


def _get_font_from_xml(run) -> str | None:
    """Прямое извлечение имени шрифта из <w:rFonts>, если API вернул None."""
    rPr = run._element.find(qn("w:rPr"))
    if rPr is not None:
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is not None:
            return rFonts.get(qn("w:hAnsi")) or rFonts.get(qn("w:ascii"))
    return None


def _get_effective_font_name(doc, run):
    """Надёжное получение имени шрифта: Run -> XML -> Стиль параграфа -> Normal."""
    # 1. Явный шрифт в Run (через API или напрямую из XML)
    font_name = run.font.name or _get_font_from_xml(run)
    if font_name:
        return font_name

    # 2. Стиль параграфа
    try:
        style_name = run._parent.style.font.name
        if style_name:
            return style_name
    except Exception:
        pass

    # 3. Стиль документа по умолчанию
    try:
        normal_name = doc.styles["Normal"].font.name
        if normal_name:
            return normal_name
    except Exception:
        pass

    return "Calibri"


def _get_effective_font_size(doc, run) -> float:
    """Надёжное получение размера шрифта с каскадным наследованием."""
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
