from config import Severity, ValidationError
from utils import _get_effective_font_name, _get_effective_font_size, _is_list_paragraph


def _check_body_font_size(document):
    if not document.main_paragraphs:
        return
    sample = document.main_paragraphs[:-2]
    for para in sample[::5]:
        for run in para.runs:
            if run.text.strip():
                size = _get_effective_font_size(document, run)
                if size is None or not (size == 12):
                    document.errors.append(
                        ValidationError(
                            "BODY_FONT",
                            f"Body font size is incorrect: {size}pt (expected 12pt)",
                            Severity.WARNING,
                            "Document body",
                        )
                    )
                    return


def _check_body_font_name(document):
    if not document.main_paragraphs:
        return

    for para in document.main_paragraphs[:-2]:
        if not para.text.strip():
            continue
        for run in para.runs:
            name = _get_effective_font_name(document.doc, run)
            # print(name, run.text)
            if name == "Calibri":
                document.errors.append(
                    ValidationError(
                        "BODY_FONT",
                        "To solve this problem, you can select all the text, change the font to Arial (even if it is still valid), and save the changes.",
                        Severity.INFO,
                        "Document body",
                    )
                )
            if name != "Arial":
                document.errors.append(
                    ValidationError(
                        "BODY_FONT",
                        f"Body font name is incorrect: {name} (expected Arial)",
                        Severity.WARNING,
                        "Document body",
                    )
                )
                return


def _check_list_font_name(document) -> None:
    """Checks the font name of all list items in the main text."""
    list_paras = [p for p in document.main_paragraphs if _is_list_paragraph(p)]
    if not list_paras:
        return
    for para in list_paras:
        for run in para.runs:
            if run.text.strip():
                name = _get_effective_font_name(document.doc, run)
                if name == "Calibri":
                    document.errors.append(
                        ValidationError(
                            "LIST_FONT",
                            "To solve this problem, you can select all the text, change the font to Arial (even if it is still valid), and save the changes.",
                            Severity.INFO,
                            "List items",
                        )
                    )
                if name != "Arial":
                    # print(name, run.text)
                    document.errors.append(
                        ValidationError(
                            "LIST_FONT",
                            f"List item font name is incorrect: {name} (expected Arial)",
                            Severity.WARNING,
                            "List items",
                        )
                    )
                    return


def _check_list_font_size(
    document, expected_size: float = 12.0, tolerance: float = 0.5
) -> None:
    """Checks the font size of all list items in the main text."""
    list_paras = [p for p in document.main_paragraphs if _is_list_paragraph(p)]
    if not list_paras:
        return
    for para in list_paras:
        for run in para.runs:
            if run.text.strip():
                size = _get_effective_font_size(document, run)
                if size is None or abs(size - expected_size) > tolerance:
                    document.errors.append(
                        ValidationError(
                            "BODY_FONT",
                            f"List item font size is incorrect: {size}pt (expected {expected_size}pt)",
                            Severity.WARNING,
                            "List items",
                        )
                    )
                    return
