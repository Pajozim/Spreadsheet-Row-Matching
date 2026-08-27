# src/excel_automation/__init__.py

from .py_files.process_catalogues import process_ctlg_files, process_workbooks

# (Optional) Cleanly expose what is available when importing this package
__all__ = [
    "process_ctlg_files",
    "process_workbooks",
    ]