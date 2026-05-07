from docx.oxml.ns import qn

from config import Severity, ValidationError


def _check_bookmarks_presence(
    document: str, bookmark_names: list[str] | None = None
) -> dict[str, bool]:
    """
    Проверяет наличие закладок (Bookmarks) в .docx файле.
    Возвращает dict вида: {'data': True, 'nomer': False}
    """

    doc = document.doc

    if bookmark_names is None:
        bookmark_names = ["data", "nomer"]

    # Используем нижний регистр для нечувствительного к регистру поиска
    target_lower = {name.lower() for name in bookmark_names}
    found = {name: False for name in bookmark_names}

    def scan_xml(xml_element):
        """Рекурсивно ищет w:bookmarkStart в XML-элементе и его потомках."""
        for bmark_start in xml_element.iter(qn("w:bookmarkStart")):
            bmark_name = bmark_start.get(qn("w:name"))
            if bmark_name and bmark_name.lower() in target_lower:
                # Ставим True для оригинального имени, которое передал пользователь
                for orig_name in bookmark_names:
                    if orig_name.lower() == bmark_name.lower():
                        found[orig_name] = True
                        break

    # 1. Основное тело документа
    scan_xml(doc.element.body)

    # 2. Заголовки и подвалы (проверяем только уникальные, не унаследованные от предыдущих секций)
    for section in doc.sections:
        for header_footer in [section.header, section.footer]:
            if not header_footer.is_linked_to_previous:
                scan_xml(header_footer._element)

    for names in bookmark_names:
        if found[names] == False:
            document.errors.append(
                ValidationError(
                    "BOOKMARKS",
                    f"Not found bookmarks '{names}'",
                    Severity.CRITICAL,
                    "All document",
                )
            )
