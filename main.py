
from pathlib import Path
import streamlit as st
import config as cfg
import pandas as pd
from itertools import chain

from src.excel_automation.py_files.process_catalogues import process_ctlg_pipeline
from src.excel_automation.py_files.stlt_helpers import get_xlsx_files, refresh_callback, col_to_num

print("Starting Streamlit ... \n")


# ==========================================
# SECTION: PAGE CONFIG (MUST BE FIRST STREAMLIT CALL)
# ==========================================
st.set_page_config(page_title="Catalogue Matcher", layout="wide")


# ==========================================
# SECTION: SESSION_STATE VISUALIZATION
# ==========================================

# st.session_state["pipeline"]
# st.session_state["current_state"]
if "paths_count" not in st.session_state:
    st.session_state["paths_count"] = 0
if "route" not in st.session_state:
    st.session_state["route"] = [False, False, False, False, False]


# SECTION: INSTRUCTION FOR "FILE PATHS" =========================================


st.title("📋 Retrieval System on Spreadsheet Files (Excel)")

st.subheader("Instructions before Usage")
st.markdown("""
    - Go from your "root folder" (where you opened this app) to `src/excel_automation/xlsx_files/` folder 
    - Copy & Paste (or drag) your catalogue excel files and your source file to the `xlsx_files/` folder.
    - Ready? Continue on here.
""")

st.divider() # ----------------------------------------------------------------------------


# SECTION: FILE PATHS ==========================================


st.subheader("Step 1: File Reader & Paths")

col1, col2 = st.columns([3, 1])

# Initialization on first run
if "file_paths_raw" not in st.session_state:
    st.session_state["file_paths_raw"] = "\n".join(get_xlsx_files(cfg.XLSX_DIR))

with col1:
    # Text area rendered first
    file_paths_raw = st.text_area(
        f"Files detected in '{cfg.XLSX_DIR}/':",
        height=140,
        key="file_paths_raw",
    )

with col2:
    #  callback fires BEFORE the page re-renders
    st.space("small")
    st.button("🔄 Refresh Directory", on_click=refresh_callback)    

# Parse and clean paths from text area
file_paths: list[str] = [p.strip() for p in file_paths_raw.splitlines() if p.strip()]

# Update Session State
if "detected_files" not in st.session_state:
    st.session_state["detected_files"] = file_paths

# Validation & Processing
if len(file_paths) != st.session_state["paths_count"]:
    st.session_state["paths_count"] = len(file_paths)
    invalid_files = [
        p for p in file_paths if not p.lower().endswith((".xlsx", ".xlsm"))
    ]
    if invalid_files:
        st.warning(
            f"⚠️ Non-supported extensions found: {' -- '.join(invf.split('/')[-1] for invf in invalid_files)}"
        )
    else:
        with st.spinner("Processing catalogue file/s ..."):
            if "pipeline" in st.session_state:
                del st.session_state["pipeline"]
                del st.session_state["current_state"]
            st.session_state["pipeline"] = process_ctlg_pipeline(file_paths)
            st.session_state["current_state"] = next(st.session_state["pipeline"])
            st.session_state["route"][0] = True
            st.rerun()

if not st.session_state["detected_files"]:
    st.info(f"No files found in '{cfg.XLSX_DIR}/'. Copy and paste files there and click Refresh.")
elif st.session_state["route"][0]:
    st.success(f"✓ Step 1 completed.")
else:
    st.info('Please Ctrl + Enter inside text area when finished selecting file paths.')


st.divider() # ----------------------------------------------------------------------------


# SECTION EXTRACTION PARAMETERS ==========================================

if "current_state" in st.session_state:
    state = st.session_state["current_state"]

    if  st.session_state["route"][0] and not st.session_state["route"][1]:

        st.subheader("Step 2: Sheets Filter")

        upd_sheets_slctn = {}

        with st.form(key="sheets_filter_form"):
            st.text("Select sheet/s to process:")

            for file_path, sheets in state["data"].items():
                file_name = Path(file_path).stem
                widget_key = f"sheet_selector_{file_name}"
                selected_sheets = st.multiselect(
                    label=f"File: **{file_name}**",
                    options=sheets,
                    default=sheets,
                    key=widget_key
                )
                upd_sheets_slctn[file_path] = selected_sheets 

            apply_sht_filter = st.form_submit_button(label="Apply Sheets Filter")

        if apply_sht_filter and state["step"] == "SELECT_SHEETS":
            st.session_state["current_state"] = st.session_state["pipeline"].send(upd_sheets_slctn)
            st.session_state["route"][1] = True
            st.rerun()
        elif state["step"] == "SELECT_SHEETS" and not st.session_state["route"][1]:
            st.info('Please confirm selection of sheets by clicking on the "Apply Sheets Filter" button.')

    elif st.session_state["route"][1]:
        st.success("✓ Step 2 completed.")

        st.divider() # ---------------------------------------------------------------------------- 

    if st.session_state["route"][1] and not st.session_state["route"][2]:

        st.subheader("Step 3: Columns Filter")

        with st.form(key="columns_filter_form"): 
            st.text("Select column/s to process:")

            upd_payload = {}

            for path, dict in state["data"].items():
                upd_payload[path] = {}
                widget_key = f"columns_selector_{Path(path).stem}"
                st.caption(f"File: **{Path(path).stem}**")
                col_from = st.text_input("From cloumn:", placeholder="e.g. B", key=f"{widget_key}-cf")
                col_to = st.text_input("To column (empty for all of them):", "", key=f"{widget_key}-ct")
                if col_from or col_to:
                    upd_payload[path].update({"min_col": col_to_num(col_from), "max_col": col_to_num(col_to) if col_to else None})
                elif not col_from and not col_to:
                    upd_payload[path].update({"min_col": 1, "max_col": None})
                

            apply_col_filter = st.form_submit_button(label="Apply Columns Filter")

            if apply_col_filter and state["step"] == "SELECT_COLUMNS":
                st.session_state["current_state"] = st.session_state["pipeline"].send(upd_payload)
                st.session_state["route"][2] = True
                st.rerun()
            elif state["step"] == "SELECT_COLUMNS" and not st.session_state["route"][2]:
                st.info('Please confirm selection of columns.')

    elif st.session_state["route"][2]:
        st.success("✓ Step 3 completed.")

        st.divider() # ----------------------------------------------------------------------------

    if st.session_state["route"][2] and not st.session_state["route"][3]:

        st.subheader("Step 4: Rows Filter")

        with st.form(key="rows_filter_form"): 
            st.text("Select row/s to process:")

            upd_payload = {}

            for path, dict in state["data"].items():
                upd_payload[path] = {}
                widget_key = f"rows_selector_{Path(path).stem}"
                st.caption(f"File: **{Path(path).stem}**")
                row_from = st.text_input("From row:", placeholder="e.g. 2", key=f"{widget_key}-rf")
                row_to = st.text_input("To row (empty for all of them):", "", key=f"{widget_key}-rt")
                if row_from or row_to:
                    upd_payload[path].update({"min_row": int(row_from), "max_row": int(row_to) if row_to else None})
                elif not row_from and not row_to:
                    upd_payload[path].update({"min_row": 1, "max_row": None})

            apply_row_filter = st.form_submit_button(label="Apply Rows Filter")

            if apply_row_filter and state["step"] == "SELECT_ROWS":
                st.session_state["current_state"] = st.session_state["pipeline"].send(upd_payload)
                st.session_state["route"][3] = True
                st.rerun()    
            elif state["step"] == "SELECT_ROWS" and not st.session_state["route"][3]:
                st.info('Please confirm selection of rows.')

    elif st.session_state["route"][3]:
        st.success("✓ Step 4 completed.")

        st.divider() # ----------------------------------------------------------------------------


    if st.session_state["route"][3] and not st.session_state["route"][4]:

        st.subheader("Step 4: Additional Filters")

        with st.form(key="addtl_filter_form"):
            st.text("Compose additional filters:")

    elif st.session_state["route"][3]:
        st.success("✓ Step 4 completed.")