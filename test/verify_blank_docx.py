#!/usr/bin/env python3
"""Check that a .docx really is an empty, blank A4 document.

Usage:
    python3 verify_blank_docx.py [path.docx]

Exits non-zero and prints the problem when the document contains text,
tables or images, or is not A4 portrait.
Requires python-docx (``pip install python-docx``).
"""

from __future__ import annotations

import sys
from pathlib import Path

from docx import Document

A4_WIDTH_MM = 210
A4_HEIGHT_MM = 297
TOLERANCE_MM = 1


def check(path: Path) -> list[str]:
    document = Document(path)
    problems: list[str] = []

    text = "".join(paragraph.text for paragraph in document.paragraphs).strip()
    if text:
        problems.append(f"body text is not empty: {text!r}")
    if document.tables:
        problems.append(f"document has {len(document.tables)} table(s)")
    if document.inline_shapes:
        problems.append(f"document has {len(document.inline_shapes)} image(s)")
    if len(document.sections) != 1:
        problems.append(f"document has {len(document.sections)} section(s), expected 1")

    for index, section in enumerate(document.sections):
        if abs(section.page_width.mm - A4_WIDTH_MM) > TOLERANCE_MM:
            problems.append(f"section {index} width is {section.page_width.mm:.1f}mm, expected {A4_WIDTH_MM}mm")
        if abs(section.page_height.mm - A4_HEIGHT_MM) > TOLERANCE_MM:
            problems.append(f"section {index} height is {section.page_height.mm:.1f}mm, expected {A4_HEIGHT_MM}mm")
        for name, part in (("header", section.header), ("footer", section.footer)):
            if "".join(paragraph.text for paragraph in part.paragraphs).strip():
                problems.append(f"section {index} {name} is not empty")

    return problems


def main(argv: list[str]) -> int:
    path = Path(argv[1]) if len(argv) > 1 else Path(__file__).with_name("blank.docx")
    problems = check(path)
    if problems:
        print(f"{path}: NOT blank")
        for problem in problems:
            print(f"  - {problem}")
        return 1
    print(f"{path}: blank A4 document, no text, tables or images")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
