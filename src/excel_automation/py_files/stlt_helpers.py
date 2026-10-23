import streamlit as st
import uuid
import config as cfg
import openpyxl as opxl

from pathlib import Path
from openpyxl.utils import get_column_letter


def get_xlsx_files(folder: Path = cfg.XLSX_DIR) -> list:
    """Scans the directory for .xlsx files and returns absolute path strings."""
    if not folder.exists():
        return []
        #folder.mkdir(parents=True, exist_ok=True)

    # Find all .xlsx / .xlsm files (excluding hidden/temp ~$ Excel files)
    files = [
        str(f.resolve())
        for f in folder.glob("*.xls*")
        if not f.name.startswith("~$")
    ]
    return sorted(files)


def refresh_callback() -> None:
    fresh_files = get_xlsx_files(cfg.XLSX_DIR)
    st.session_state["file_paths_raw"] = "\n".join(fresh_files)


def list_sheets(parameters: dict) -> None:
    for path in parameters:
        wb = opxl.load_workbook(path, read_only=True)
        parameters[path]["selected_sheets"] = wb.sheetnames
        wb.close()


def col_to_num(col: str, idx: int = 1) -> int:
    """Converts Excel column letter(s) like 'A' or 'AA' to 1-based index.
    Unless idx is set to 0, which then returns 0-based index.
    """
    col = col.strip().upper()
    
    num = 0
    for char in col:
        num = num * 26 + (ord(char) - ord('A') + idx)
    return num


def add_filter_callback(crntPrmtrs_with_path: dict):
    """
    Appends a new filter dictionary with a unique ID to the specific file's filter list.
    """
    # 1. Create new filter item with unique ID
    new_filter_id = str(uuid.uuid4())[:8]  # Short unique hash
    new_filter = {
        "id": new_filter_id,
        "targ_col_raw": "",
        "omit_row_when": "",
    }

    # 2. Mutate list in session state
    crntPrmtrs_with_path["filters"].append(new_filter)


def delete_filter_callback(filter_id: str, crntPrmtrs_with_path: dict):
    """
    Deletes a filter by its unique ID and cleans up associated widget state keys.
    """

    if crntPrmtrs_with_path:
        # 1. Filter out the targeted item by ID
        crntPrmtrs_with_path["filters"] = [
            f for f in crntPrmtrs_with_path["filters"] if f["id"] != filter_id
        ]
    else:
        print("Error: cur_prmters_with_path is None")

    # 2. Clean up explicit widget keys from session state to avoid key ghosting
    widget_keys_to_remove = [
        f"targ_col_{filter_id}",
        f"omit_when_{filter_id}",
        f"del_{filter_id}",
    ]
    for key in widget_keys_to_remove:
        if key in st.session_state:
            del st.session_state[key]


def add_category_callback() -> None:
    st.session_state["db_template"].append(
        {"id": str(uuid.uuid4())[:8], "name": ""} 
    )


def delete_category_callback(item_id: str):
    st.session_state["db_template"] = [
        c for c in st.session_state["db_template"] if c["id"] != item_id
    ]

    # Clean up explicit widget keys from session state to avoid key ghosting
    widget_keys_to_remove = [
        f"cat_name_{item_id}",
        f"del_{item_id}",
    ]
    for key in widget_keys_to_remove:
        if key in st.session_state:
            del st.session_state[key]


def get_last_column(path: Path) -> int:
    return opxl.load_workbook(path).worksheets[0].max_column


def get_letters(min_col: int, max_col: int) -> list:
    return [get_column_letter(i) for i in range(min_col, max_col + 1)]



