# -*- coding: utf-8 -*-
"""
office_com — MS Office COM Automation Library
==============================================
Unified Pythonic interface for Word, Excel, and PowerPoint COM automation.

Capabilities beyond file-level libraries (python-docx/openpyxl/python-pptx):
    - Exact PDF export (rendered by the real Office engine)
    - Format conversion (.doc -> .docx, .xls -> .xlsx, .ppt -> .pptx, etc.)
    - Chart image export (Excel)
    - Slide image export (PowerPoint)
    - VBA macro execution
    - .doc / .xls / .ppt old-format support

Usage:
    from office_com import WordApp, ExcelApp, PptApp

    # Context manager (recommended):
    with WordApp() as w:
        w.export_pdf("report.docx", "report.pdf")
        w.convert("old.doc", "new.docx")

    with ExcelApp() as xl:
        xl.export_pdf("data.xlsx", "data.pdf")
        xl.convert("old.xls", "new.xlsx")
        xl.export_charts("data.xlsx", "charts/")

    with PptApp() as ppt:
        ppt.export_pdf("slides.pptx", "slides.pdf")
        ppt.export_slides("slides.pptx", "images/")

Requirements:
    - Windows only (COM is Windows-only)
    - Microsoft Office installed
    - pywin32 (pip install pywin32)
"""

from ._base import AppError
from .word_app import WordApp
from .excel_app import ExcelApp
from .ppt_app import PptApp

__version__ = "0.1.0"
__all__ = ["WordApp", "ExcelApp", "PptApp", "AppError"]
