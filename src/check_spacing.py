import re

from config import Severity, ValidationError
from text_normalization import _normalize_text
from utils import _get_effective_font_size, _get_table_font_size


def _check_structural_spacing(document) -> None:
    """
    Проверяет обязательные пустые строки между разделами.
    Использует позиционное определение блоков, что исключает ошибки
    при совпадении текста в таблице 'Краткое содержание' с паттернами заголовка/даты.
    """
    # 1. Строим линейную последовательность всех элементов в порядке XML
    doc_sequence = []
    table_idx = 0
    para_idx = 0

    for child in document.doc.element.body:
        tag = child.tag.split("}")[-1]  # 'p' или 'tbl'

        if tag == "p":
            if para_idx < len(document.paragraphs):
                para = document.paragraphs[para_idx]
                doc_sequence.append(
                    {
                        "type": "para",
                        "element": para,
                        "text": _normalize_text(para.text).lower(),
                        "is_empty": not para.text.strip(),
                    }
                )
                para_idx += 1

        elif tag == "tbl":
            if table_idx < len(document.tables):
                table = document.tables[table_idx]
                # Собираем текст из таблицы для возможной верификации
                tbl_text = " ".join(
                    _normalize_text(p.text).lower()
                    for row in table.rows
                    for cell in row.cells
                    for p in cell.paragraphs
                )
                doc_sequence.append(
                    {
                        "type": "table",
                        "element": table,
                        "text": tbl_text,
                        "is_empty": not tbl_text.strip(),
                    }
                )
                table_idx += 1

    # 2. ПОЗИЦИОННОЕ определение индексов (строго по вашей структуре)
    table_indices = [i for i, x in enumerate(doc_sequence) if x["type"] == "table"]

    # Дата и номер - первая таблица
    idx_date = table_indices[0] if len(table_indices) > 0 else None
    # Краткое содержание/Заголовок - вторая таблица
    idx_heading = table_indices[1] if len(table_indices) > 1 else None
    # Подписант - последняя таблица в документе
    idx_signatory = table_indices[-2] if len(table_indices) > 0 else None

    # Поиск преамбулы и распорядительной части среди параграфов
    # Ищем только после таблицы заголовка, чтобы не задеть шапку
    search_start = idx_heading if idx_heading is not None else 0
    idx_preamble = None
    idx_dispositive = None

    preamble_pats = [r"в\s+соответствии", r"согласно", r"в\s+целях", r"во\s+исполнение"]
    dispositive_pats = [r"приказываю"]

    for i in range(search_start, len(doc_sequence)):
        item = doc_sequence[i]
        if item["type"] != "para":
            continue
        if idx_preamble is None and any(
            re.search(p, item["text"]) for p in preamble_pats
        ):
            idx_preamble = i
        if idx_dispositive is None and any(
            re.search(p, item["text"]) for p in dispositive_pats
        ):
            idx_dispositive = i
        if idx_preamble and idx_dispositive:
            break

    if idx_dispositive == None:
        idx_dispositive = idx_preamble + 2

    # 3. Вспомогательные функции
    def count_empty_between(start_idx, end_idx):
        """Считает ТОЛЬКО пустые параграфы. Таблицы игнорируются."""
        if start_idx is None or end_idx is None or end_idx <= start_idx + 1:
            return 0
        return sum(
            1
            for i in range(start_idx + 1, end_idx)
            if doc_sequence[i]["type"] == "para" and doc_sequence[i]["is_empty"]
        )

    def check_font(item_idx, expected_size, rule_name):
        """Проверяет размер шрифта в параграфе или таблице."""
        if item_idx is None or item_idx >= len(doc_sequence):
            return
        item = doc_sequence[item_idx]
        size = None

        if item["type"] == "para":
            for run in item["element"].runs:
                if run.text.strip():
                    size = _get_effective_font_size(document.doc, run)
                    break
        elif item["type"] == "table":
            size = _get_table_font_size(item["element"])

        if size is not None and abs(size - expected_size) > 0.5:
            document.errors.append(
                ValidationError(
                    "SPACING_FONT",
                    f"{rule_name}: {size:.1f}pt (expected {expected_size}pt)",
                    Severity.WARNING,
                    rule_name,
                )
            )
        return size

    # === ПРИМЕНЕНИЕ ПРАВИЛ ===

    # Правило 1: Заголовок ↔ Дата (1 пустая строка, 10pt)
    if idx_heading is not None and idx_date is not None:
        gap = count_empty_between(
            min(idx_heading, idx_date), max(idx_heading, idx_date)
        )
        if gap != 1:
            document.errors.append(
                ValidationError(
                    "SPACING_EMPTY",
                    f"Heading-Date: {gap} (expected 1)",
                    Severity.WARNING,
                    "Header",
                )
            )
        for idx in range(min(idx_heading, idx_date) + 1, max(idx_heading, idx_date)):
            check_font(idx, 10.0, "Heading-Date")

    # Правило 2: Шапка → Текст (2 пустые строки, 12pt)
    header_end_idx = (
        max(idx_heading, idx_date)
        if idx_heading is not None and idx_date is not None
        else None
    )
    if header_end_idx is not None and idx_preamble is not None:
        gap = count_empty_between(header_end_idx, idx_preamble)
        if gap != 2:
            document.errors.append(
                ValidationError(
                    "SPACING_EMPTY",
                    f"Header-Body: {gap} (expected 2)",
                    Severity.WARNING,
                    "Header to Text",
                )
            )
        for idx in range(header_end_idx + 1, idx_preamble):
            check_font(idx, 10.0, "Preamble-Text")

    # Правило 3: Преамбула → Приказываю (1 пустая строка, 12pt)
    if idx_preamble is not None and idx_dispositive is not None:
        gap = count_empty_between(idx_preamble, idx_dispositive)
        if gap != 1:
            document.errors.append(
                ValidationError(
                    "SPACING_EMPTY",
                    f"Preamble-Order: {gap} (expected 1)",
                    Severity.WARNING,
                    "Body",
                )
            )
        for idx in range(idx_preamble, idx_dispositive):
            check_font(idx, 12.0, "Dispositive")

    # Правило 4: Текст → Подписант (2 пустые строки, 12pt)
    body_end_idx = idx_dispositive if idx_dispositive is not None else idx_preamble
    if body_end_idx is not None and idx_signatory is not None:
        gap = count_empty_between(body_end_idx, idx_signatory)
        if gap != 2:
            document.errors.append(
                ValidationError(
                    "SPACING_EMPTY",
                    f"Body-Signature: {gap} (expected 2)",
                    Severity.WARNING,
                    "Text to Signature",
                )
            )
        for idx in range(body_end_idx, idx_signatory):
            check_font(idx, 12.0, "Signature")
