from typing import List

from docx import Document

from check_appendix import (
    _check_appendix_font_name,
    _check_appendix_font_size,
    _check_appendix_format,
)
from check_body import (
    _check_body_font_name,
    _check_body_font_size,
    _check_command_word,
    _check_control_clause_position,
    _check_heading,
    _check_indents,
    _check_list_font_name,
    _check_list_font_size,
    _check_orientation,
    _check_preamble_structure,
    _check_reduction_position,
    _check_signature_block,
    _check_signature_font_size,
)
from config import Severity, ValidationError
from tables import _check_table_fonts_name, _check_table_fonts_size
from text_normalization import (
    _clean_empty_paragraphs,
    _find_main_document_boundary,
    _normalize_text,
)


class DocumentArchitectureValidator:
    def __init__(self, filepath: str):
        self.filepath = filepath
        self.doc = Document(filepath)
        self.errors: List[ValidationError] = []

        # Clean up empty paragraphs
        _clean_empty_paragraphs(self)

        self.paragraphs = self.doc.paragraphs
        self.tables = self.doc.tables

        # Find the boundary between the main text and appendices
        self.main_doc_end_idx = _find_main_document_boundary(self)

        # Split document parts
        self.main_paragraphs = self.paragraphs[: self.main_doc_end_idx]
        self.appendix_paragraphs = self.paragraphs[self.main_doc_end_idx :]

        # Normalize text
        self.main_text = _normalize_text(
            "\n".join([p.text.strip() for p in self.main_paragraphs if p.text.strip()])
        )
        self.full_text = _normalize_text(
            "\n".join([p.text.strip() for p in self.paragraphs if p.text.strip()])
        )

    def validate(self) -> List[ValidationError]:
        _check_heading(self)
        _check_body_font_size(self)
        _check_preamble_structure(self)
        _check_command_word(self)
        _check_control_clause_position(self)
        _check_signature_block(self)
        _check_table_fonts_size(self)
        _check_table_fonts_name(self)
        _check_reduction_position(self)
        _check_list_font_size(self)
        _check_body_font_name(self)
        _check_list_font_name(self)
        _check_signature_font_size(self)
        _check_indents(self)
        _check_appendix_format(self)
        _check_appendix_font_name(self)
        _check_appendix_font_size(self)
        _check_orientation(self)
        return self.errors

    def print_report(self) -> None:
        if not self.errors:
            print("Document architecture meets requirements.")
            return

        critical = [e for e in self.errors if e.severity == Severity.CRITICAL]
        warnings = [e for e in self.errors if e.severity == Severity.WARNING]

        print("\n" + "=" * 70)
        print("DOCUMENT ARCHITECTURE VALIDATION REPORT")
        print("=" * 70)
        print(f"File: {self.filepath}")
        print(f"Total findings: {len(self.errors)}\n")

        if critical:
            print("CRITICAL ERRORS:")
            for err in critical:
                print(f"  {err.severity.value} | {err.rule}")
                print(f"    {err.message}")
                print()

        if warnings:
            print("WARNINGS:")
            for err in warnings:
                print(f"  {err.severity.value} | {err.rule}")
                print(f"    {err.message}")
                print()

        print("=" * 70)


if __name__ == "__main__":
    validator = DocumentArchitectureValidator("standart.docx")
    validator.validate()
    validator.print_report()
