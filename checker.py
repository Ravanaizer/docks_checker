# document_architect_validator.py
import re
from dataclasses import dataclass
from enum import Enum
from typing import List

from docx import Document


class Severity(Enum):
    CRITICAL = "CRITICAL"
    WARNING = "WARNING"
    INFO = "INFO"


@dataclass
class ValidationError:
    rule: str
    message: str
    severity: Severity
    location: str = ""


class DocumentArchitectureValidator:
    def __init__(self, filepath: str):
        self.filepath = filepath
        self.doc = Document(filepath)
        self.errors: List[ValidationError] = []
        self.check_paragraph()
        self.paragraphs = self.doc.paragraphs
        self.tables = self.doc.tables

        # Find boundary before normalization
        self.main_doc_end_idx = self._find_main_document_boundary()

        # Split into parts
        self.main_paragraphs = self.paragraphs[: self.main_doc_end_idx]
        self.appendix_paragraphs = self.paragraphs[self.main_doc_end_idx :]

        # Normalize text (collapse multiple whitespaces)
        self.main_text = self._normalize_text(
            "\n".join([p.text.strip() for p in self.main_paragraphs if p.text.strip()])
        )
        self.main_text_lower = self.main_text.lower()

        self.full_text = self._normalize_text(
            "\n".join([p.text.strip() for p in self.paragraphs if p.text.strip()])
        )
        self.full_text_lower = self.full_text.lower()

    def check_paragraph(self):
        """Delete empty paragraph (only enter)"""
        for p in reversed(self.doc.paragraphs):
            if not p.text.strip():
                p._element.getparent().remove(p._element)
                # p._p = None

    def _normalize_text(self, text: str) -> str:
        """Replaces all multiple whitespaces/tabs with single space."""
        return re.sub(r"\s+", " ", text).strip()

    def _find_main_document_boundary(self) -> int:
        """Finds where main text ends and appendices begin."""
        appendix_markers = [
            r"^приложение\s*[№№]?\s*[а-яё1-9]",
            r"^прил\.\s*",
        ]

        for i, para in enumerate(self.paragraphs):
            text = self._normalize_text(para.text.strip().lower())
            if any(re.search(pattern, text) for pattern in appendix_markers):
                return i
        return len(self.paragraphs)

    def validate(self) -> List[ValidationError]:
        self._check_preamble_structure()
        self._check_command_word()
        self._check_control_clause_position()
        self._check_signature_block()
        self._check_appendix_format()
        self._check_table_fonts()
        self._check_reduction_position()
        return self.errors

    def _check_preamble_structure(self):
        preamble_patterns = [
            r"в\s+соответствии",
            r"во\s+исполнение",
            r"согласно",
            r"в\s+целях",
            r"в\s+связи\s+с",
        ]
        preamble_section = self._normalize_text(
            "\n".join([p.text for p in self.paragraphs[:10]])
        ).lower()

        if not any(re.search(p, preamble_section) for p in preamble_patterns):
            self.errors.append(
                ValidationError(
                    rule="PREAMBLE",
                    message="Justification (preamble) is missing. Should be: 'В соответствии.../Во исполнение...' ",
                    severity=Severity.CRITICAL,
                    location="Start of document",
                )
            )

    def _check_command_word(self):
        # Check for command word without broken spaces between letters
        good_pattern = [r"п\s*р\s*и\s*к\s*а\s*з\s*ы\s*в\s*а\s*ю", r"приказываю"]

        if not any(re.search(g, self.full_text_lower) for g in good_pattern):
            self.errors.append(
                ValidationError(
                    rule="COMMAND",
                    message="Word 'Приказываю' contains extra spaces between letters or is missing",
                    severity=Severity.CRITICAL,
                    location="Dispositive section",
                )
            )

    def _check_control_clause_position(self):
        """Searches control clause in normalized main text."""
        control_patterns = [
            r"контроль\s+(за\s+исполнением|оставляю|возложить)",
            r"контроль\s+оставляю\s+за\s+собой",
            r"возложить\s+контроль",
        ]

        # Search in last 10 paragraphs of MAIN text
        search_text = self._normalize_text(
            "\n".join([p.text for p in self.main_paragraphs[-10:]])
        ).lower()

        if not any(re.search(p, search_text) for p in control_patterns):
            self.errors.append(
                ValidationError(
                    rule="CONTROL",
                    message="Clause about control of order execution is missing",
                    severity=Severity.CRITICAL,
                    location="End of dispositive section",
                )
            )

    def _check_reduction_position(self):
        """Searches reduction list in normalized main text."""
        control_patterns = [
            r"довести\s+настоящий\s+приказ\s+до\s+сведения",
            r"довести\s+до\s+сведения",
            # r"довести",
        ]

        # Search in last 10 paragraphs of MAIN text
        search_text = self._normalize_text(
            "\n".join([p.text for p in self.main_paragraphs[-10:]])
        ).lower()

        if not any(re.search(p, search_text) for p in control_patterns):
            self.errors.append(
                ValidationError(
                    rule="REDUCTION",
                    message="No section that requires information",
                    severity=Severity.CRITICAL,
                    location="End of dispositive section",
                )
            )

    def _check_signature_block(self):
        """FIXED SIGNATURE BLOCK CHECK"""
        # After normalization "Джежора  Е.А." becomes "Джежора Е.А."
        signature_patterns = [
            r"[а-яё]+\s+[а-яё]\.[а-яё]\.",  # Lastname I.O. (Djegera E.A.)
            r"[а-яё]\.[а-яё]\.\s+[а-яё]+",  # I.O. Lastname
            r"директор.*[а-яё]+\s+[а-яё]\.",  # Director ... Name
            r"руководитель.*[а-яё]+\s+[а-яё]\.",  # Head ... Name
        ]

        signature_section = self._normalize_text(
            "\n".join([p.text for p in self.main_paragraphs[-5:]])
        ).lower()

        if not any(re.search(p, signature_section) for p in signature_patterns):
            self.errors.append(
                ValidationError(
                    rule="SIGNATURE",
                    message="No head signature block found (Position + Initials + Lastname)",
                    severity=Severity.CRITICAL,
                    location="End of main text",
                )
            )

    def _check_appendix_format(self):
        if not self.appendix_paragraphs:
            return

        forbidden_letters = "ёзйочьыъ"
        for para in self.appendix_paragraphs[:20]:  # Check only headers
            text = self._normalize_text(para.text.lower())
            if "приложение" in text:
                match = re.search(r"приложение\s*[№№]?\s*([а-яё])", text)
                if match and match.group(1) in forbidden_letters:
                    self.errors.append(
                        ValidationError(
                            rule="APPENDIX",
                            message=f"Forbidden letter in appendix designation: '{match.group(1).upper()}'",
                            severity=Severity.WARNING,
                            location="Appendix header",
                        )
                    )

    def _check_table_fonts(self):
        if not self.tables:
            return

        for table in self.tables:
            if self._is_table_font_too_small(table):
                self.errors.append(
                    ValidationError(
                        rule="TABLES",
                        message="Font size is too small in table.",
                        severity=Severity.WARNING,
                        location="Tables",
                    )
                )

    def _is_table_font_too_small(self, table):
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if run.font.size and run.font.size.pt < 8:
                            return True
        return False

    def print_report(self):
        if not self.errors:
            print("Document architecture meets requirements")
            return

        critical = [e for e in self.errors if e.severity == Severity.CRITICAL]
        warnings = [e for e in self.errors if e.severity == Severity.WARNING]
        info = [e for e in self.errors if e.severity == Severity.INFO]

        print("\n" + "=" * 70)
        print("DOCUMENT ARCHITECTURE VALIDATION REPORT")
        print("=" * 70)
        print(f"File: {self.filepath}")
        print(f"Main text boundary: paragraph #{self.main_doc_end_idx}")
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

        if info:
            print("INFO:")
            for err in info:
                print(f"  {err.severity.value} | {err.rule}")
                print(f"    {err.message}")
                print()

        print("=" * 70)


if __name__ == "__main__":
    validator = DocumentArchitectureValidator("test.docx")
    errors = validator.validate()
    validator.print_report()
