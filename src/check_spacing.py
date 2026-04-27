import re

from config import Severity, ValidationError
from text_normalization import _normalize_text
from utils import _get_paragraph_font_size


def _check_structural_spacing(document) -> None:
    """
    Validates mandatory empty lines between document sections.
    Uses positional indexing to handle mixed structures (tables + paragraphs).
    Checks both line count and font size of the empty separator paragraphs.
    """
    # 1. Build a linear sequence of all body elements preserving XML order
    doc_sequence = []
    table_idx = 0
    para_idx = 0

    for child in document.doc.element.body:
        tag = child.tag.split("}")[-1]  # 'p' or 'tbl'

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

    # 2. Positional indexing based on document structure:
    # [Header Image] -> Table[0]=Date -> Table[1]=Heading/Content -> Text -> Table[-1]=Signature
    table_indices = [i for i, x in enumerate(doc_sequence) if x["type"] == "table"]

    idx_date = table_indices[0] if len(table_indices) > 0 else None
    idx_heading = table_indices[1] if len(table_indices) > 1 else None
    idx_signatory = table_indices[-2] if len(table_indices) > 0 else None

    # Search for preamble and dispositive markers only after the heading table
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

    if idx_dispositive is None and idx_preamble is not None:
        idx_dispositive = idx_preamble + 2

    # 3. Helper: Count only empty paragraphs between two indices
    def count_empty_between(start_idx, end_idx):
        if start_idx is None or end_idx is None or end_idx <= start_idx + 1:
            return 0
        return sum(
            1
            for i in range(start_idx + 1, end_idx)
            if doc_sequence[i]["type"] == "para" and doc_sequence[i]["is_empty"]
        )

    # 4. Helper: Validate font size of empty separator paragraphs
    def check_empty_lines_font(start_idx, end_idx, expected_size, rule_name):
        if start_idx is None or end_idx is None:
            return
        for i in range(start_idx + 1, end_idx):
            item = doc_sequence[i]
            if item["type"] == "para" and item["is_empty"]:
                size = _get_paragraph_font_size(document.doc, item["element"])
                if size is not None and abs(size - expected_size) > 0.5:
                    document.errors.append(
                        ValidationError(
                            "SPACING_FONT",
                            f"{rule_name}: Empty line {size:.1f}pt (expected {expected_size}pt)",
                            Severity.WARNING,
                            rule_name,
                        )
                    )

    # === APPLY SPACING RULES ===

    # Rule 1: Heading ↔ Date (1 empty line, 10pt)
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
        check_empty_lines_font(
            min(idx_heading, idx_date), max(idx_heading, idx_date), 10.0, "Heading-Date"
        )

    # Rule 2: Header → Body Text (2 empty lines, 12pt)
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
        check_empty_lines_font(header_end_idx, idx_preamble, 12.0, "Header-Body")

    # Rule 3: Preamble → Dispositive (1 empty line, 12pt)
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
        check_empty_lines_font(idx_preamble, idx_dispositive, 12.0, "Preamble-Order")

    # Rule 4: Body Text → Signature (2 empty lines, 12pt)
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
        check_empty_lines_font(body_end_idx, idx_signatory, 12.0, "Body-Signature")
