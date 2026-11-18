#!/usr/bin/env python3
"""Generate a blank Word (.docx) document.

The result is a single empty A4 page: no text, headers, footers, tables or
images, just like a new Word document.

Usage:
    python3 make_blank_docx.py [output.docx]

Defaults to writing ``blank.docx`` next to this script.
Requires python-docx (``pip install python-docx``).
"""

from __future__ import annotations

import sys
from pathlib import Path

from docx import Document
from docx.shared import Mm


def build(output: Path) -> Path:
    """Write a blank A4 document to *output* and return the path."""
    document = Document()
    section = document.sections[0]
    # A4 portrait, the usual default for Chinese documents.
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    # A new Word document starts with a single empty paragraph.
    if not document.paragraphs:
        document.add_paragraph("")
    props = document.core_properties
    props.title = ""
    props.subject = ""
    props.author = ""
    props.comments = ""
    props.last_modified_by = ""
    document.save(output)
    return output


def main(argv: list[str]) -> int:
    output = Path(argv[1]) if len(argv) > 1 else Path(__file__).with_name("blank.docx")
    build(output)
    print(f"wrote {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
