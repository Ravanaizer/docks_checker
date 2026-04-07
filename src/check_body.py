import re

from config import Severity, ValidationError
from tables import _check_table_font_size, _get_table_text
from text_normalization import _normalize_text
from utils import _get_effective_font_name, _get_effective_font_size, _is_list_paragraph


def _check_heading(document):
    patterns = [
        r"об\s+утверждении",
        r"о\s+подготовке\s+номенклатуры\s+дел",
        r"о\s+внесении\s+изменений\s+в\s+приказ",
    ]
    found = False
    for table in document.tables[:3]:
        text = "\n".join(_get_table_text(document, table)).lower()
        if any(re.search(p, text) for p in patterns):
            found = True
            if not _check_table_font_size(document, table, 10):
                document.errors.append(
                    ValidationError(
                        "HEADING",
                        "Font size in heading table is incorrect (expected 10pt)",
                        Severity.CRITICAL,
                        "Table at start",
                    )
                )
            break

    if not found:
        for para in document.paragraphs[:10]:
            text = _normalize_text(para.text).lower()
            if any(re.search(p, text) for p in patterns):
                found = True
                for run in para.runs:
                    if run.text.strip():
                        size = _get_effective_font_size(document, run)
                        if not (9.5 <= size <= 10.5):
                            document.errors.append(
                                ValidationError(
                                    "HEADING",
                                    f"Font size is incorrect: {size}pt (expected 10pt)",
                                    Severity.CRITICAL,
                                    "Start of document",
                                )
                            )
                break

    if not found:
        document.errors.append(
            ValidationError(
                "HEADING",
                "Heading is missing.",
                Severity.CRITICAL,
                "Start of document",
            )
        )


def _check_body_font_name(document):
    if not document.main_paragraphs:
        return
    sample = document.main_paragraphs[:-2]
    for para in sample[::5]:
        for run in para.runs:
            if run.text.strip():
                name = _get_effective_font_name(document, run)
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


def _check_signature_font_size(document):
    if not document.main_paragraphs:
        return
    sample = document.main_paragraphs[-2:]
    for para in sample[::5]:
        for run in para.runs:
            if run.text.strip():
                size = _get_effective_font_size(document, run)
                if size is None or not (size == 10):
                    document.errors.append(
                        ValidationError(
                            "SIGNATURE",
                            f"Signature font size is incorrect: {size}pt (expected 10pt)",
                            Severity.WARNING,
                            "Document body",
                        )
                    )
                    return


def _check_preamble_structure(document):
    patterns = [
        r"в\s+соответствии",
        r"во\s+исполнение",
        r"согласно",
        r"в\s+целях",
        r"в\s+связи\s+с",
    ]
    text = _normalize_text(
        "\n".join([p.text for p in document.paragraphs[:10]])
    ).lower()
    if not any(re.search(p, text) for p in patterns):
        document.errors.append(
            ValidationError(
                "PREAMBLE",
                "Justification (preamble) is missing.",
                Severity.CRITICAL,
                "Start of document",
            )
        )


def _check_command_word(document):
    patterns = [r"п\sр\sи\sк\sа\sз\sы\sв\sа\s*ю", r"приказываю"]
    if not any(re.search(p, document.full_text.lower()) for p in patterns):
        document.errors.append(
            ValidationError(
                "COMMAND",
                "Word 'Приказываю' contains extra spaces or is missing.",
                Severity.CRITICAL,
                "Dispositive section",
            )
        )


def _check_control_clause_position(document):
    patterns = [
        r"контроль\s+(за\s+исполнением|оставляю|возложить)",
        r"контроль\s+оставляю\s+за\s+собой",
        r"возложить\s+контроль",
    ]
    text = _normalize_text(
        "\n".join([p.text for p in document.main_paragraphs[-10:]])
    ).lower()
    if not any(re.search(p, text) for p in patterns):
        document.errors.append(
            ValidationError(
                "CONTROL",
                "Clause about control of order execution is missing.",
                Severity.CRITICAL,
                "End of dispositive section",
            )
        )


def _check_reduction_position(document):
    patterns = [
        r"довести\s+настоящий\s+приказ\s+до\s+сведения",
        r"довести\s+до\s+сведения",
    ]
    text = _normalize_text(
        "\n".join([p.text for p in document.main_paragraphs[-10:]])
    ).lower()
    if not any(re.search(p, text) for p in patterns):
        document.errors.append(
            ValidationError(
                "REDUCTION",
                "No section requiring information distribution.",
                Severity.CRITICAL,
                "End of dispositive section",
            )
        )


def _check_signature_block(document):
    patterns = [
        r"[а-яё]+\s+[а-яё].[а-яё].",
        r"[а-яё].[а-яё].\s+[а-яё]+",
        r"директор.[а-яё]+\s+[а-яё].",
        r"руководитель.[а-яё]+\s+[а-яё].",
    ]
    found = False
    text = _normalize_text(
        "\n".join([p.text for p in document.main_paragraphs[-5:]])
    ).lower()
    if any(re.search(p, text) for p in patterns):
        found = True
    if not found and document.tables:
        for table in document.tables[-3:]:
            table_text = "\n".join(_get_table_text(document, table)).lower()
            if any(re.search(p, table_text) for p in patterns):
                found = True
                break

    if not found:
        document.errors.append(
            ValidationError(
                "SIGNATURE",
                "No head signature block found.",
                Severity.CRITICAL,
                "End of main text",
            )
        )


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


def _check_list_font_name(document) -> None:
    """Checks the font name of all list items in the main text."""
    list_paras = [p for p in document.main_paragraphs if _is_list_paragraph(p)]
    if not list_paras:
        return
    for para in list_paras:
        for run in para.runs:
            if run.text.strip():
                name = _get_effective_font_name(document.doc, run)
                if name != "Arial":
                    document.errors.append(
                        ValidationError(
                            "LIST_FONT",
                            f"List item font name is incorrect: {name} (expected Arial)",
                            Severity.WARNING,
                            "List items",
                        )
                    )
                    return


def _check_indents(doc):
    """
    Validates that first-line indents in main paragraphs match the standard (1.25 cm).
    Accounts for style inheritance, EMU unit conversion, and floating-point tolerance.
    """
    EXPECTED_INDENT_CM = 1.25
    TOLERANCE_CM = (
        0.05  # Allow small rounding differences from Word's internal precision
    )

    for para in doc.main_paragraphs:
        # Skip empty paragraphs to avoid false positives
        if not para.text.strip():
            continue

        # Step 1: Try to get explicit first-line indent from paragraph formatting
        first_line = para.paragraph_format.first_line_indent

        # Step 2: If not explicitly set, check the paragraph style (Word often stores indents there)
        if first_line is None and para.style is not None:
            try:
                first_line = para.style.paragraph_format.first_line_indent
            except Exception:
                pass  # Style may not have indent defined; fall through to default handling

        # Step 3: Evaluate correctness and prepare display value
        is_incorrect = False
        display_value = "Not set (inherited)"

        if first_line is not None:
            # Convert EMU (English Metric Units) to centimeters for human-readable comparison
            indent_cm = first_line.cm
            display_value = f"{indent_cm:.2f} cm"
            # Use tolerance-based comparison to handle Word's internal rounding
            if abs(indent_cm - EXPECTED_INDENT_CM) > TOLERANCE_CM:
                is_incorrect = True
        else:
            # Standard requires explicit 1.25 cm indent; missing value is an issue
            is_incorrect = True

        if is_incorrect:
            # Truncate text for readability in the report
            preview = para.text.strip()
            preview = preview[:60] + "..." if len(preview) > 60 else preview

            doc.errors.append(
                ValidationError(
                    "PARAGRAPH_INDENTS",
                    f"First-line indent incorrect: {display_value} (expected {EXPECTED_INDENT_CM} cm) | Text: '{preview}'",
                    Severity.WARNING,
                    "Paragraph formatting",
                )
            )
