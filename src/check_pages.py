import re

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

from config import Severity, ValidationError


def _has_page_field_in_xml(xml_element) -> bool:
    """
    Глубокий скан XML элемента (шапки/подвала/параграфа) на наличие поля PAGE.
    Игнорирует NUMPAGES, обрабатывает вложенные структуры Word и пустые значения.
    """
    if xml_element is None:
        return False

    for elem in xml_element.iter():
        # Современный формат: <w:instrText xml:space="preserve"> PAGE </w:instrText>
        if elem.tag == qn("w:instrText"):
            content = (elem.text or "").upper()
            # Строгое совпадение: PAGE есть, но NUMPAGES нет
            if "PAGE" in content and "NUMPAGES" not in content:
                return True

        # Устаревший формат: <w:fldSimple w:instr=" PAGE \* MERGEFORMAT "/>
        elif elem.tag == qn("w:fldSimple"):
            instr = (elem.get(qn("w:instr")) or "").upper()
            if "PAGE" in instr and "NUMPAGES" not in instr:
                return True

    return False


def _get_effective_alignment(para) -> str | None:
    """
    Возвращает выравнивание параграфа с учётом наследования от стиля.
    """
    # 1. Явное выравнивание
    if para.alignment is not None:
        return para.alignment

    # 2. Наследование от стиля
    if para.style is not None:
        try:
            style_align = para.style.paragraph_format.alignment
            if style_align is not None:
                return style_align
        except Exception:
            pass
    return None


def _check_page_numbering(document) -> None:
    """
    Валидация нумерации страниц:
    1. Стр. 1: БЕЗ номера.
    2. Стр. 2+: НОМЕР ЕСТЬ, ТОЛЬКО в шапке, СТРОГО по центру.
    """
    section = document.doc.sections[0]
    sect_pr = section._sectPr

    if sect_pr is None:
        document.errors.append(
            ValidationError(
                "PAGE_NUMBERING",
                "Unable to read section properties.",
                Severity.WARNING,
                "Structure",
            )
        )
        return

    # 1. Проверяем флаг "Особый колонтитул для первой страницы"
    if sect_pr.find(qn("w:titlePg")) is None:
        document.errors.append(
            ValidationError(
                "PAGE_NUMBERING",
                "'Different First Page' setting is missing. Page 1 must be unique.",
                Severity.WARNING,
                "Page Setup",
            )
        )
        return

    # Вспомогательная функция валидации одного колонтитула
    def _validate_hf(hf_obj, expect_number: bool, location: str):
        if hf_obj is None:
            if expect_number:
                document.errors.append(
                    ValidationError(
                        "PAGE_NUMBERING",
                        f"{location}: Header/Footer object is missing.",
                        Severity.WARNING,
                        "Content",
                    )
                )
            return

        # Глубокий поиск поля PAGE в XML шапки/подвала
        has_page = _has_page_field_in_xml(hf_obj._element)

        # Проверка наличия/отсутствия номера
        if expect_number and not has_page:
            document.errors.append(
                ValidationError(
                    "PAGE_NUMBERING",
                    f"{location}: Page number field is missing.",
                    Severity.WARNING,
                    "Content",
                )
            )
        elif not expect_number and has_page:
            document.errors.append(
                ValidationError(
                    "PAGE_NUMBERING",
                    f"{location}: Page number detected, but must be EMPTY.",
                    Severity.WARNING,
                    "Content",
                )
            )
            return  # Если номер запрещён, дальнейшие проверки не нужны

        # Если номер должен быть -> проверяем центрирование
        if expect_number and has_page:
            for para in hf_obj.paragraphs:
                # Находим именно тот параграф, где лежит поле PAGE
                if _has_page_field_in_xml(para._element):
                    align = _get_effective_alignment(para)
                    if align != WD_ALIGN_PARAGRAPH.CENTER:
                        document.errors.append(
                            ValidationError(
                                "PAGE_NUMBERING",
                                f"{location}: Page number is not centered (alignment: {align}).",
                                Severity.WARNING,
                                "Alignment",
                            )
                        )
                    break  # Достаточно проверить один параграф с номером

    # === ЗАПУСК ПРОВЕРОК ===

    # Страница 1: Номер ЗАПРЕЩЁН
    _validate_hf(
        section.first_page_header, expect_number=False, location="Page 1 Header"
    )
    _validate_hf(
        section.first_page_footer, expect_number=False, location="Page 1 Footer"
    )

    # Страницы 2+: Номер ОБЯЗАТЕЛЕН, только в шапке, по центру
    _validate_hf(section.header, expect_number=True, location="Pages 2+ Header")
    _validate_hf(section.footer, expect_number=False, location="Pages 2+ Footer")
