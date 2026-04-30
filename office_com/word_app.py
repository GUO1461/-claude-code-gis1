# -*- coding: utf-8 -*-
"""MS Word COM automation — PDF export, format conversion, macro execution."""

import os
from ._base import _OfficeApp, AppError


class WordApp(_OfficeApp):
    """Word COM application wrapper.

    Usage:
        with WordApp() as word:
            word.export_pdf("input.docx", "output.pdf")
            word.convert("old.doc", "new.docx")
    """

    _progids = ["Word.Application", "Word.Application.16"]
    _app_name = "Word"

    def _setup(self):
        super()._setup()

    # ── document operations ──────────────────────────────────────

    def open(self, path):
        path = self._abs_path(path)
        if not os.path.exists(path):
            raise AppError(f"File not found: {path}")
        try:
            return self._app.Documents.Open(path)
        except Exception as e:
            raise AppError(f"Failed to open {path}: {e}") from e

    def create(self):
        return self._app.Documents.Add()

    # ── PDF export ───────────────────────────────────────────────

    def export_pdf(self, source, target=None):
        """Export a Word document to PDF.

        Args:
            source: Path to .docx/.doc file
            target: Output PDF path (defaults to source stem + .pdf)
        """
        source = self._abs_path(source)
        if target is None:
            target = os.path.splitext(source)[0] + ".pdf"
        target = self._abs_path(target)
        self._ensure_dir(target)

        doc = self.open(source)
        try:
            doc.ExportAsFixedFormat(target, 17)  # wdFormatPDF
        finally:
            try: doc.Close(0)
            except Exception: pass
        return target

    # ── format conversion ────────────────────────────────────────

    _FORMAT_MAP = {
        ".docx": 16, ".doc": 0, ".pdf": 17,
        ".html": 8, ".txt": 5, ".rtf": 6,
    }

    def convert(self, source, target):
        source = self._abs_path(source)
        target = self._abs_path(target)
        self._ensure_dir(target)

        ext = os.path.splitext(target)[1].lower()
        if ext == ".pdf":
            return self.export_pdf(source, target)

        fmt = self._FORMAT_MAP.get(ext)
        if fmt is None:
            raise AppError(f"Unsupported target format: {ext}")

        doc = self.open(source)
        try:
            doc.SaveAs(target, fmt)
        finally:
            try: doc.Close(0)
            except Exception: pass
        return target

    # ── run macro ────────────────────────────────────────────────

    def run_macro(self, macro_name, *args):
        self._app.Run(macro_name, *args)

    # ── classmethod shortcuts ────────────────────────────────────

    @classmethod
    def quick_export_pdf(cls, source, target=None):
        with cls() as app:
            return cls.export_pdf(app, source, target)

    @classmethod
    def quick_convert(cls, source, target):
        with cls() as app:
            return cls.convert(app, source, target)
