# -*- coding: utf-8 -*-
"""Base COM application class — lifecycle, cleanup, common operations."""

import os
import pythoncom
import win32com.client


class AppError(Exception):
    """Raised when COM application operations fail."""
    pass


class _OfficeApp:
    """Base class for MS Office COM applications.

    Manages COM lifecycle: CoInitialize → Dispatch → ... → Quit → CoUninitialize.
    Use as context manager to guarantee cleanup even on exceptions.
    """

    _progids = []
    _app_name = "Office"

    def __init__(self, visible=False):
        self._app = None
        self._visible = visible

    def __enter__(self):
        pythoncom.CoInitialize()
        for progid in self._progids:
            try:
                self._app = win32com.client.Dispatch(progid)
                break
            except Exception:
                continue
        if self._app is None:
            pythoncom.CoUninitialize()
            raise AppError(f"Could not connect to {self._app_name}. Tried: {self._progids}")
        self._setup()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self._app is not None:
            try:
                self._app.Quit()
            except Exception:
                pass
            self._app = None
        pythoncom.CoUninitialize()
        return False

    @property
    def app(self):
        if self._app is None:
            raise AppError(f"{self._app_name} not connected.")
        return self._app

    def _setup(self):
        try:
            self._app.Visible = self._visible
        except Exception:
            pass
        try:
            self._app.DisplayAlerts = False
        except Exception:
            pass

    @staticmethod
    def _abs_path(path):
        return os.path.abspath(path)

    @staticmethod
    def _ensure_dir(path):
        d = os.path.dirname(os.path.abspath(path))
        os.makedirs(d, exist_ok=True)
        return path
