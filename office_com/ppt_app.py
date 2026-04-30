# -*- coding: utf-8 -*-
"""MS PowerPoint COM automation — PDF export, slide image export, format conversion."""

import os
from ._base import _OfficeApp, AppError


class PptApp(_OfficeApp):
    """PowerPoint COM application wrapper.

    Usage:
        with PptApp() as ppt:
            ppt.export_pdf("slides.pptx", "slides.pdf")
            ppt.export_slides("slides.pptx", "output/images/")
    """

    _progids = ["PowerPoint.Application", "PowerPoint.Application.16"]
    _app_name = "PowerPoint"

    def _setup(self):
        try:
            self._app.DisplayAlerts = 2  # ppAlertsNone
        except Exception:
            pass

    # ── presentation operations ──────────────────────────────────

    def open(self, path):
        path = self._abs_path(path)
        if not os.path.exists(path):
            raise AppError(f"File not found: {path}")
        try:
            return self._app.Presentations.Open(path)
        except Exception as e:
            raise AppError(f"Failed to open {path}: {e}") from e

    def create(self):
        """Create a new blank presentation."""
        return self._app.Presentations.Add()

    def save_as(self, pres, path):
        """Save presentation."""
        path = self._abs_path(path)
        self._ensure_dir(path)
        try:
            self._app.DisplayAlerts = 2
        except Exception:
            pass
        pres.SaveAs(path)

    # ── PDF export ───────────────────────────────────────────────

    def export_pdf(self, source, target=None):
        """Export presentation to PDF.

        Args:
            source: Path to .pptx/.ppt file
            target: Output PDF path

        Returns:
            Path to generated PDF.
        """
        source = self._abs_path(source)
        if target is None:
            target = os.path.splitext(source)[0] + ".pdf"
        target = self._abs_path(target)
        self._ensure_dir(target)

        pres = self.open(source)
        try:
            self._app.DisplayAlerts = 2
            pres.SaveAs(target, 32)
        finally:
            pres.Close()

        if not os.path.exists(target):
            raise AppError(f"PDF export failed: {target}")
        return target

    # ── slide image export ───────────────────────────────────────

    def export_slides(self, source, output_dir=None, image_format="png",
                      width=None, height=None):
        """Export each slide as an image.

        Args:
            source: Path to .pptx/.ppt file
            output_dir: Output dir (default: source dir)
            image_format: 'png', 'jpg', 'bmp', 'gif', 'tif'
            width, height: Output dimensions in pixels (default: slide default)

        Returns:
            List of exported image paths.
        """
        source = self._abs_path(source)
        if output_dir is None:
            output_dir = os.path.dirname(source) or "."
        output_dir = os.path.abspath(output_dir)
        os.makedirs(output_dir, exist_ok=True)

        base_name = os.path.splitext(os.path.basename(source))[0]
        pres = self.open(source)
        exported = []

        try:
            for i, slide in enumerate(pres.Slides, 1):
                filename = f"{base_name}_slide_{i:02d}.{image_format}"
                filepath = os.path.join(output_dir, filename)

                if width and height:
                    slide.Export(filepath, image_format.upper(), width, height)
                else:
                    # Use slide default dimensions
                    sw = pres.PageSetup.SlideWidth
                    sh = pres.PageSetup.SlideHeight
                    slide.Export(filepath, image_format.upper(), sw, sh)

                exported.append(filepath)
        finally:
            pres.Close()

        return exported

    # ── format conversion ────────────────────────────────────────

    def convert(self, source, target):
        """Convert presentation to different format."""
        source = self._abs_path(source)
        target = self._abs_path(target)
        self._ensure_dir(target)

        ext = os.path.splitext(target)[1].lower()
        if ext == ".pdf":
            return self.export_pdf(source, target)

        # PpSaveAsFileType constants
        fmt_map = {
            ".pptx": 11,  # ppSaveAsOpenXMLPresentation
            ".ppt": 1,    # ppSaveAsPresentation
            ".pptm": 29,  # ppSaveAsOpenXMLPresentationMacroEnabled
            ".ppsx": 27,  # ppSaveAsOpenXMLShow
            ".jpg": 17,   # ppSaveAsJPG
            ".png": 18,   # ppSaveAsPNG
            ".bmp": 19,   # ppSaveAsBMP
        }
        fmt = fmt_map.get(ext)
        if fmt is None:
            raise AppError(f"Unsupported target format: {ext}")

        pres = self.open(source)
        try:
            self._app.DisplayAlerts = 2
            pres.SaveAs(target, fmt)
        finally:
            pres.Close()
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
    def quick_export_slides(cls, source, output_dir=None, image_format="png",
                            width=None, height=None):
        with cls() as app:
            return cls.export_slides(app, source, output_dir, image_format, width, height)

    @classmethod
    def quick_convert(cls, source, target):
        with cls() as app:
            return cls.convert(app, source, target)
