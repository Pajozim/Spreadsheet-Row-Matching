import streamlit as st
import uuid
import config as cfg
import openpyxl as opxl
from openpyxl.utils import get_column_letter
from pathlib import Path
from openpyxl import Workbook
#from src.excel_automation.py_files.process_catalogues import process_ctlg_files

#@st.cache_resource
#def get_cache()

"""visualization of:
st.session_state["parameters"] = parameters = {
    "PATH1": {
        "selected_sheets": ["sheet1", "sheet2", ...],
        "min_row": 1,
        "max_row": None,
        "min_col": 1,
        "max_col": None,
        "filters": [
            {
                "id": "1",
                "targ_col_raw": "B",
                "omit_row_when": "Sidenote",
                "parsed_tCol": 2
            },
            {...},
            ...
        ],
        "column_mappings": [ 
                {"DB_CAT1": [1, 2, 3], "delimiter": "+"},
                {"DB_CAT2": [4], "delimiter": ""},
                {"DB_CAT3": [5]}
            ],
    },
    "PATH2": {
        ...
    },
    ...
}
"""

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


def col_to_num(col: str) -> int | None:
    """Converts Excel column letter(s) like 'A' or 'AA' to 1-based index."""
    col = col.strip().upper()
    if not col or not col.isalpha():
        return None
    
    num = 0
    for char in col:
        num = num * 26 + (ord(char) - ord('A') + 1)
    return num


def add_filter_callback(file_path: str):
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
    st.session_state["parameters"][file_path]["filters"].append(new_filter)


def delete_filter_callback(filter_id: str, file_path: str):
    """
    Deletes a filter by its unique ID and cleans up associated widget state keys.
    """
    state = st.session_state.get("parameters", {})
    fPath = state.get(file_path, {})
    filters_data = fPath.get("filters", [])

    if state and file_path and filters_data:
        # 1. Filter out the targeted item by ID
        st.session_state["parameters"][file_path]["filters"] = [
            f for f in st.session_state["parameters"][file_path]["filters"] if f["id"] != filter_id
        ]

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