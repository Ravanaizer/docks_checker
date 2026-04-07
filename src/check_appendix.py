import re

from config import Severity, ValidationError
from text_normalization import _normalize_text


def _check_appendix_format(document):
    forbidden = "ёзйочьыъ"
    headers = []

    # Поиск в параграфах
    for para in document.appendix_paragraphs[:30]:
        text = _normalize_text(para.text.lower())
        if "приложение" in text:
            headers.append({"text": text, "source": "paragraph"})

    # Поиск в таблицах
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    text = _normalize_text(para.text.lower())
                    if "приложение" in text:
                        headers.append({"text": text, "source": "table"})

    found_letters = []
    found_numbers = []

    for h in headers:
        text = h["text"]
        letter_match = re.search(r"приложение\s*[№№]?\s*([а-яё])", text)
        number_match = re.search(r"приложение\s*[№№]?\s*(\d+)", text)

        if letter_match:
            letter = letter_match.group(1)
            found_letters.append(letter)
            if letter in forbidden:
                document.errors.append(
                    ValidationError(
                        "APPENDIX",
                        f"Forbidden letter in appendix: '{letter.upper()}'",
                        Severity.WARNING,
                        "Appendix header",
                    )
                )

        if number_match:
            found_numbers.append(int(number_match.group(1)))

    if found_letters:
        _check_sequential_numbering(document, found_letters, "letter")
    if found_numbers:
        _check_sequential_numbering(document, found_numbers, "number")


def _check_sequential_numbering(document, items, type="letter"):
    if len(items) < 2:
        return

    if type == "letter":
        alphabet = "абвгдежзийклмнопрстуфхцчшщъыьэюя"
        indices = [alphabet.index(i) if i in alphabet else -1 for i in items]
        for i in range(1, len(indices)):
            if (
                indices[i] != -1
                and indices[i - 1] != -1
                and indices[i] - indices[i - 1] > 1
            ):
                document.errors.append(
                    ValidationError(
                        "APPENDIX",
                        f"Appendix numbering not sequential: missing letter between "
                        f"'{alphabet[indices[i - 1]].upper()}' and '{alphabet[indices[i]].upper()}'",
                        Severity.WARNING,
                        "Appendix section",
                    )
                )
    elif type == "number":
        sorted_items = sorted(set(items))
        for i in range(1, len(sorted_items)):
            if sorted_items[i] - sorted_items[i - 1] > 1:
                document.errors.append(
                    ValidationError(
                        "APPENDIX",
                        f"Appendix numbering not sequential: missing number {sorted_items[i - 1] + 1}",
                        Severity.WARNING,
                        "Appendix section",
                    )
                )
