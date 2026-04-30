# -*- coding: utf-8 -*-
"""MS Excel COM automation — PDF export, format conversion, chart/image export."""

import os
from ._base import _OfficeApp, AppError


class ExcelApp(_OfficeApp):
    """Excel COM application wrapper.

    Usage:
        with ExcelApp() as excel:
            wb = excel.open("data.xlsx")
            excel.export_pdf("data.xlsx", "data.pdf")
            excel.convert("old.xls", "new.xlsx")
    """

    _progids = ["Excel.Application", "Excel.Application.16"]
    _app_name = "Excel"

    def _setup(self):
        super()._setup()

    # ── workbook operations ──────────────────────────────────────

    def open(self, path, read_only=False):
        path = self._abs_path(path)
        if not os.path.exists(path):
            raise AppError(f"File not found: {path}")
        try:
            return self._app.Workbooks.Open(path, ReadOnly=read_only)
        except Exception as e:
            raise AppError(f"Failed to open {path}: {e}") from e

    def create(self):
        """Create a new blank workbook."""
        return self._app.Workbooks.Add()

    def save_as(self, wb, path, file_format=None):
        """Save workbook to path."""
        path = self._abs_path(path)
        self._ensure_dir(path)
        if file_format is not None:
            wb.SaveAs(path, file_format)
        else:
            wb.SaveAs(path)

    # ── PDF export ───────────────────────────────────────────────

    def export_pdf(self, source, target=None, sheets=None):
        """Export Excel workbook (or specific sheets) to PDF.

        Args:
            source: Path to .xlsx/.xls file
            target: Output PDF path
            sheets: List of sheet names/indices to export (default: all)

        Returns:
            Path to generated PDF.
        """
        source = self._abs_path(source)
        if target is None:
            target = os.path.splitext(source)[0] + ".pdf"
        target = self._abs_path(target)
        self._ensure_dir(target)

        wb = self.open(source)
        try:
            if sheets:
                # Select specific sheets
                sheet_names = []
                for s in sheets:
                    if isinstance(s, int):
                        sheet_names.append(wb.Worksheets(s).Name)
                    else:
                        sheet_names.append(str(s))
                wb.Worksheets(sheet_names).Select()

            wb.ExportAsFixedFormat(0, target)  # 0 = xlTypePDF
        finally:
            try: wb.Close(0)
            except Exception: pass

        if not os.path.exists(target):
            raise AppError(f"PDF export failed: {target}")
        return target

    # ── format conversion ────────────────────────────────────────

    _FORMAT_MAP = {
        ".xlsx": 51,   # xlOpenXMLWorkbook
        ".xls": 56,    # xlExcel8
        ".xlsm": 52,   # xlOpenXMLWorkbookMacroEnabled
        ".csv": 6,     # xlCSV
        ".txt": 20,    # xlTextWindows
        ".pdf": 0,     # handled separately
        ".html": 44,   # xlHtml
    }

    def convert(self, source, target):
        """Convert between formats. Extension determines output format."""
        source = self._abs_path(source)
        target = self._abs_path(target)
        self._ensure_dir(target)

        ext = os.path.splitext(target)[1].lower()
        if ext == ".pdf":
            return self.export_pdf(source, target)

        fmt = self._FORMAT_MAP.get(ext)
        if fmt is None:
            raise AppError(f"Unsupported target format: {ext}")

        wb = self.open(source)
        try:
            wb.SaveAs(target, fmt)
        finally:
            try: wb.Close(0)
            except Exception: pass
        return target

    # ── chart / image export ─────────────────────────────────────

    def export_charts(self, source, output_dir=None, image_format="png"):
        """Export all charts in a workbook as images.

        Args:
            source: Path to Excel workbook
            output_dir: Output dir for chart images (default: source dir)
            image_format: 'png' or 'jpg'

        Returns:
            List of exported image paths.
        """
        source = self._abs_path(source)
        if output_dir is None:
            output_dir = os.path.dirname(source) or "."
        output_dir = os.path.abspath(output_dir)
        os.makedirs(output_dir, exist_ok=True)

        base_name = os.path.splitext(os.path.basename(source))[0]
        wb = self.open(source)
        exported = []

        try:
            for ws in wb.Worksheets:
                for i, co in enumerate(ws.ChartObjects(), 1):
                    filename = f"{base_name}_{ws.Name}_chart{i}.{image_format}"
                    filepath = os.path.join(output_dir, filename)
                    co.Chart.Export(filepath)
                    exported.append(filepath)
        finally:
            try: wb.Close(0)
            except Exception: pass

        return exported

    # ── run macro ────────────────────────────────────────────────

    def run_macro(self, macro_name, *args):
        """Execute a VBA macro."""
        self._app.Run(macro_name, *args)

    # ── classmethod shortcuts ────────────────────────────────────

    @classmethod
    def quick_export_pdf(cls, source, target=None, sheets=None):
        with cls() as app:
            return cls.export_pdf(app, source, target, sheets)

    @classmethod
    def quick_convert(cls, source, target):
        with cls() as app:
            return cls.convert(app, source, target)

    @classmethod
    def quick_export_charts(cls, source, output_dir=None, image_format="png"):
        with cls() as app:
            return cls.export_charts(app, source, output_dir, image_format)
