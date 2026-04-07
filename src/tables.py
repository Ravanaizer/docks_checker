from typing import List

from config import Severity, ValidationError
from utils import _get_effective_font_size, _get_table_font_name, _get_table_font_size


def _check_table_fonts_size(document):
    for table in document.tables:
        if _get_table_font_size(table) < 8:
            document.errors.append(
                ValidationError(
                    "TABLES",
                    "Font size is too small in table.",
                    Severity.WARNING,
                    "Tables",
                )
            )


def _check_table_fonts_name(document):
    for table in document.tables:
        name = _get_table_font_name(document, table)
        if name != "Arial":
            document.errors.append(
                ValidationError(
                    "TABLES",
                    f"Table font name is incorrect: {name} (expected Arial)",
                    Severity.WARNING,
                    "Tables",
                )
            )


def _get_table_text(document, table) -> List[str]:
    """Collects text from all cells in the table."""
    lines = []
    for row in table.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                if para.text.strip():
                    lines.append(para.text.strip())
    return lines


def _check_table_font_size(
    document, table, expected_size: float, tolerance: float = 0.5
) -> bool:
    """Checks that all text in the table matches the expected font size."""
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    if run.text.strip():
                        size = _get_effective_font_size(document, run)
                        if size is not None and abs(size - expected_size) > tolerance:
                            return False
    return True
