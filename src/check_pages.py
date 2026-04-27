import re

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

from config import Severity, ValidationError


# ==========================================
# HELPER XML FUNCTIONS
# ==========================================
def _has_page_field_in_xml(xml_elem) -> bool:
    """Reliably searches for the PAGE field in XML, ignoring NUMPAGES and nested structures."""
    if xml_elem is None:
        return False
    for elem in xml_elem.iter():
        if elem.tag == qn("w:instrText"):
            content = (elem.text or "").upper()
            if "PAGE" in content and "NUMPAGES" not in content:
                return True
        elif elem.tag == qn("w:fldSimple"):
            instr = (elem.get(qn("w:instr")) or "").upper()
            if "PAGE" in instr and "NUMPAGES" not in instr:
                return True
    return False


def _is_centered_in_xml(para_xml) -> bool:
    """Checks centering directly via the w:jc XML attribute."""
    pPr = para_xml.find(qn("w:pPr"))
    if pPr is not None:
        jc = pPr.find(qn("w:jc"))
        if jc is not None:
            return jc.get(qn("w:val")) == "center"
    return False


# ==========================================
# PAGE NUMBERING VALIDATION
# ==========================================
def _validate_hf(hf_obj, expect_number: bool, label: str, errors: list):
    """Universal validation for a single header/footer."""
    if hf_obj is None or hf_obj._element is None:
        if expect_number:
            errors.append(
                ValidationError(
                    "PAGE_NUMBERING",
                    f"{label}: Header/Footer object is empty.",
                    Severity.WARNING,
                    "Header/Footer",
                )
            )
        return

    has_field = _has_page_field_in_xml(hf_obj._element)

    if expect_number and not has_field:
        errors.append(
            ValidationError(
                "PAGE_NUMBERING",
                f"{label}: Page number field (PAGE) is missing.",
                Severity.WARNING,
                "Content",
            )
        )
    elif not expect_number and has_field:
        errors.append(
            ValidationError(
                "PAGE_NUMBERING",
                f"{label}: Contains a page number field, but must be empty.",
                Severity.WARNING,
                "Content",
            )
        )
        return

    if expect_number and has_field:
        centered_found = False
        for para_xml in hf_obj._element.iter(qn("w:p")):
            if _has_page_field_in_xml(para_xml):
                if _is_centered_in_xml(para_xml):
                    centered_found = True
                break
        if not centered_found:
            errors.append(
                ValidationError(
                    "PAGE_NUMBERING",
                    f"{label}: Page number is not centered.",
                    Severity.WARNING,
                    "Alignment",
                )
            )


def _check_page_numbering(document) -> None:
    """Page numbering validation: Page 1 has no number, Pages 2+ have a centered number in the header."""
    section = document.doc.sections[0]
    sect_pr = section._sectPr
    if sect_pr is None:
        document.errors.append(
            ValidationError(
                "PAGE_NUMBERING",
                "Failed to read section properties.",
                Severity.WARNING,
                "Structure",
            )
        )
        return

    if sect_pr.find(qn("w:titlePg")) is None:
        document.errors.append(
            ValidationError(
                "PAGE_NUMBERING",
                "'Different First Page' option is not enabled.",
                Severity.WARNING,
                "Page Setup",
            )
        )
        return

    _validate_hf(
        section.first_page_header,
        expect_number=False,
        label="Page 1 Header",
        errors=document.errors,
    )
    _validate_hf(
        section.first_page_footer,
        expect_number=False,
        label="Page 1 Footer",
        errors=document.errors,
    )
    _validate_hf(
        section.header,
        expect_number=True,
        label="Pages 2+ Header",
        errors=document.errors,
    )
    _validate_hf(
        section.footer,
        expect_number=False,
        label="Pages 2+ Footer",
        errors=document.errors,
    )


# ==========================================
# EMPTY LINES IN FOOTER CHECK (XML LEVEL)
# ==========================================
def _check_footer_empty_lines(document) -> None:
    """
    Counts empty paragraphs in the first page footer directly via XML.
    Bypasses python-docx limitations to detect all paragraph marks (¶).
    """
    section = document.doc.sections[0]
    sect_pr = section._sectPr
    if sect_pr is None:
        document.errors.append(
            ValidationError(
                "FOOTER_LAYOUT",
                "Section properties are inaccessible.",
                Severity.WARNING,
                "Structure",
            )
        )
        return

    # Determine which footer to check: first page or main
    has_diff_first = sect_pr.find(qn("w:titlePg")) is not None
    target_hf = section.first_page_footer if has_diff_first else section.footer
    location_label = "First Page Footer" if has_diff_first else "Main Footer"

    if target_hf is None or target_hf._element is None:
        document.errors.append(
            ValidationError(
                "FOOTER_LAYOUT",
                f"{location_label} is inaccessible.",
                Severity.WARNING,
                "Structure",
            )
        )
        return

    # === DIRECT PARAGRAPH COUNT VIA XML ===
    empty_para_count = 0
    for p_xml in target_hf._element.iter(qn("w:p")):
        # Collect all text from <w:t> nodes within the paragraph
        text = "".join(t.text for t in p_xml.iter(qn("w:t")) if t.text)
        if not text.strip():
            empty_para_count += 1

    if empty_para_count != 2:
        document.errors.append(
            ValidationError(
                "FOOTER_LAYOUT",
                f"{location_label}: Found {empty_para_count} empty line(s), exactly 2 required.",
                Severity.WARNING,
                "Footer Layout",
            )
        )
