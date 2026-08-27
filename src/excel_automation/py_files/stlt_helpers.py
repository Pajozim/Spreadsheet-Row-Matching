from pathlib import Path
from config import XLSX_DIR
import streamlit as st
from openpyxl import Workbook
from src.excel_automation.py_files.process_catalogues import process_ctlg_files, process_workbooks


def get_xlsx_files(folder: Path = XLSX_DIR) -> str | list[str]:
    """Scans the directory for .xlsx files and returns absolute path strings."""
    if not folder.exists():
        print(f"Folder {folder} does not exist.")
        return "Folder must exist. Accidently changed path? Check config.py on this root folder (not the config file among main.py)."
        #folder.mkdir(parents=True, exist_ok=True)

    # Find all .xlsx / .xlsm files (excluding hidden/temp ~$ Excel files)
    files = [
        str(f.resolve())
        for f in folder.glob("*.xls*")
        if not f.name.startswith("~$")
    ]
    return sorted(files)

@st.cache_data
def get_cache(ctlg_files: list[str] | None = None) -> list | str | None:
    if not ctlg_files:
        return
    elif ctlg_files:
        return process_ctlg_files(st.session_state["detected_files"])


def filter_deletion(f_count: int) -> None:
    st.session_state["selected_parameters"]["filters"].pop(f_count)
    st.session_state["selected_parameters"]["filter_count"] -= 1


def col_to_num(col: str) -> int | None:
    """Converts Excel column letter(s) like 'A' or 'AA' to 1-based index."""
    col = col.strip().upper()
    if not col or not col.isalpha():
        return None
    
    num = 0
    for char in col:
        num = num * 26 + (ord(char) - ord('A') + 1)
    return num


def add_filter_callback() -> None:
    """Callback to append a new empty filter dictionary."""
    st.session_state["selected_parameters"]["filters"].append({
        "targ_col_raw": "",
        "omit_row_when": ""
    })

def delete_filter_callback(index: int) -> None:
    """Callback to remove a filter by index."""
    filters = st.session_state["selected_parameters"]["filters"]
    if 0 <= index < len(filters):
        filters.pop(index)

def parse_target_cols(targ_col_str: str) -> list:
    """Helper to parse raw text input into column numbers."""
    if not targ_col_str.strip():
        return [None]
    return [
        num
        for c in targ_col_str.split(",") 
        if (num := col_to_num(c)) is not None
    ]

