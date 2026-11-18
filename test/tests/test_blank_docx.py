"""Tests for the blank .docx generator and verifier.

Run with: python3 -m unittest discover -s tests -v
Requires python-docx; the tests are skipped when it is unavailable.
"""

from __future__ import annotations

import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

try:
    import docx  # noqa: F401
except ModuleNotFoundError:  # pragma: no cover - dependency missing
    docx = None

PROJECT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT))

if docx is not None:
    import make_blank_docx
    import verify_blank_docx


@unittest.skipIf(docx is None, "python-docx is not installed")
class BlankDocxTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.out = Path(self.tmp.name) / "blank.docx"

    def test_generated_document_passes_verification(self) -> None:
        make_blank_docx.build(self.out)
        self.assertEqual(verify_blank_docx.check(self.out), [])

    def test_generated_document_is_a_valid_zip_package(self) -> None:
        make_blank_docx.build(self.out)
        with zipfile.ZipFile(self.out) as package:
            self.assertIsNone(package.testzip())
            self.assertIn("word/document.xml", package.namelist())

    def test_verifier_rejects_text(self) -> None:
        make_blank_docx.build(self.out)
        document = docx.Document(self.out)
        document.add_paragraph("hello")
        document.save(self.out)
        problems = verify_blank_docx.check(self.out)
        self.assertTrue(any("body text" in problem for problem in problems))

    def test_verifier_rejects_letter_page(self) -> None:
        make_blank_docx.build(self.out)
        document = docx.Document(self.out)
        from docx.shared import Inches

        document.sections[0].page_width = Inches(8.5)
        document.sections[0].page_height = Inches(11)
        document.save(self.out)
        problems = verify_blank_docx.check(self.out)
        self.assertTrue(any("width" in problem for problem in problems))

    def test_committed_blank_document_matches_generator(self) -> None:
        committed = PROJECT / "blank.docx"
        self.assertTrue(committed.is_file(), "blank.docx is missing from the project")
        self.assertEqual(verify_blank_docx.check(committed), [])


if __name__ == "__main__":
    unittest.main()
