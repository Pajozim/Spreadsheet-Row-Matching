
from pathlib import Path
import streamlit as st
import config as cfg
import pandas as pd

from src.excel_automation.py_files.process_catalogues import process_workbooks
from src.excel_automation.py_files.stlt_helpers import get_xlsx_files, get_cache, col_to_num, add_filter_callback, delete_filter_callback

print("Starting Streamlit ... \n")


# ==========================================
# SECTION: PAGE CONFIG (MUST BE FIRST STREAMLIT CALL)
# ==========================================
st.set_page_config(page_title="Catalogue Matcher", layout="wide")


# ==========================================
# SECTION: STATE INITIALIZATION (RUNS ONCE ONLY)
# ==========================================
if "initialized" not in st.session_state:
    st.session_state.update({
        "detected_files": [],
        "workbooks": [],
        "selected_sheets": [],
        "selected_parameters": {
            "min_col": 1, 
            "max_col": None, 
            "min_row": 1, 
            "max_row": None,
            "filters": []
        },
        "indexing_running": False,
        "steps": [False, False, False, False, False, False],
        "initialized": True  # Prevents state from resetting on subsequent reruns
    })
if "merged_groups" not in st.session_state:
    st.session_state["merged_groups"] = []  # List of dicts: {"name": str, "cols": list[str], "delimiter": str}
if "column_renames" not in st.session_state:
    st.session_state["column_renames"] = {}  # {orig_col: custom_name}


# # SECTION: INSTRUCTION FOR "FILE PATHS" =========================================


st.title("📋 Retrieval System on Spreadsheet Files (Excel)")

st.subheader("Instructions before Usage")
st.markdown("""
    - Go from your "root folder" (where you opened this app) to `src/excel_automation/xlsx_files/` folder 
    - Copy & Paste (or drag) your catalogue excel files and your source file to the `xlsx_files/` folder.
    - Ready? Continue on here.
""")

st.divider()


# # SECTION: FILE PATHS ==========================================


st.subheader("Step 1: File Reader & Paths")

col1, col2 = st.columns([3, 1])

with col1:
    file_paths_raw = st.text_area(
        f"Edit the lines to add or remove file paths.\nFiles detected in '{cfg.XLSX_DIR}/':",
        value="\n".join(get_xlsx_files(cfg.XLSX_DIR)),
        height=140
    )

with col2:
    st.space("small")
    if st.button("🔄 Refresh Directory"):
        st.rerun()

# Parse and clean paths from text area
file_paths: list[str] = [p.strip() for p in file_paths_raw.splitlines() if p.strip()]

# Update Session State
st.session_state["detected_files"] = file_paths

# Validation & Processing
if file_paths:
    invalid_files = [
        p for p in file_paths if not p.lower().endswith((".xlsx", ".xlsm"))
    ]
    if invalid_files:
        st.warning(
            f"⚠️ Non-supported extensions found: {' -- '.join(invf.split('/')[-1] for invf in invalid_files)}"
        )
        st.session_state["steps"][0] = False
    else:
        # Load workbooks if not already loaded into state
        if not st.session_state["workbooks"] or len(file_paths) != len(st.session_state["workbooks"]):
            with st.spinner("Processing catalogue file/s ..."):
                st.session_state["workbooks"] = get_cache(file_paths)
                
        st.session_state["steps"][0] = True
        st.success("✓ Step 1 completed.")
else:
    st.info(f"No files found in '{cfg.XLSX_DIR}/'. Copy files there and click Refresh.")
    st.session_state["steps"][0] = False

st.divider()


# SECTION EXTRACTION PARAMETERS ==========================================


if st.session_state["steps"][0]:

    st.subheader("Step 2: Extraction Parameters")

    # 1. Extract sheets from pre-loaded session workbooks
    processed_sheets = process_workbooks(st.session_state["workbooks"])
    # output: [{"path1": ["sheet1", "sheet2", ...]}, {"path2": ["sheet1", "sheet2", ...]}, ...]
    
    available_sheet_names = []
    if not processed_sheets:
        st.warning("⚠️ No sheets found in selected workbooks.")
    else:
        for wb_dict in processed_sheets:
            path, sheetslist = list(wb_dict.items())[0]
            # print("wb_dict: ", path, sheetslist[0])
            for sheetslist in wb_dict.values():
                sheet_names = [sheet.title for sheet in sheetslist]
                available_sheet_names.extend(sheet_names)

    p_col1, p_col2, p_col3, p_col4 = st.columns([2, 1, 1, 1])

    with p_col1:
        st.markdown("**Sheets Filter**")
        
        with st.form(key="sheets_filter_form"):
            selected_sheets = st.multiselect(
                label="Select sheets to process:",
                options=available_sheet_names,
                default=available_sheet_names,
                key="selected_sheets_filter"
            )
            
            apply_button = st.form_submit_button(label="Apply Filter")

            if apply_button:
                list_of_dict: list[dict[str, list]] = []
                if processed_sheets and selected_sheets:
                    for wb_dict in processed_sheets:
                        for path, sheetlist in wb_dict.items():
                            sheets_on_path = [sheet for sheet in sheetlist if sheet.title in selected_sheets]
                            if sheets_on_path:
                                wb_dict[path] = sheets_on_path
                        

                st.session_state["selected_sheets"] = processed_sheets
                st.session_state["steps"][1] = True

        if not st.session_state["steps"][1]:
            st.info('Please confirm selection of sheets by clicking on the "Apply Filter" button.')
        else:
            st.success("✓")

    if st.session_state["steps"][1]:

        apply_col_filter = False
        apply_row_filter = False
        apply_cond_filters = False

        with p_col2:
            st.markdown("**Columns Filter**")

            with st.form(key="columns_filter_form"): 
                st.text("Select column/s to process:")
                col_from = st.text_input("From cloumn:", "A")
                col_to = st.text_input("To column (empty for all of them):", "")

                apply_col_filter = st.form_submit_button(label="Apply Columns Filter")

                if apply_col_filter:
                    st.session_state["selected_parameters"]["min_col"] = col_to_num(col_from)
                    st.session_state["selected_parameters"]["max_col"] = col_to_num(col_to) if col_to else None
                    st.session_state["steps"][2] = True

            if not st.session_state["steps"][2]:
                st.info('Please confirm selection of columns.')
            else:
                st.success("✓")
                
        with p_col3:
            st.markdown("**Rows Filter**")

            with st.form(key="rows_filter_form"):
                st.text("Select row/s to process:")
                row_from = st.text_input("From row:", "1")
                row_to = st.text_input("To row (empty for all of them):", "")

                apply_row_filter = st.form_submit_button(label="Apply Rows Filter")

                if apply_row_filter:
                    st.session_state["selected_parameters"]["min_row"] = int(row_from)
                    st.session_state["selected_parameters"]["max_row"] = int(row_to) if row_to else None
                    st.session_state["steps"][3] = True

            if not st.session_state["steps"][3]:
                st.info('Please confirm selection of rows.')
            else:
                st.success("✓")

        with p_col4:
            st.markdown("**Additional Filters**")
            st.info("Any empty rows are automatically omitted (including those with accidental whitespaces).")

            # 1. Render existing filter forms dynamically inside p_col4
            filters = st.session_state["selected_parameters"]["filters"]

            for idx, f in enumerate(filters):
                with st.container(border=True):
                    st.caption(f"Filter #{idx + 1}")
                    
                    # Read and bind inputs using unique state keys
                    f["targ_col_raw"] = st.text_input(
                        label="Omit row when in column (leave empty for all):",
                        value=f.get("targ_col_raw", ""),
                        placeholder="e.g.: B",
                        key=f"targ_col_{f['id']}"
                    )
                    
                    f["omit_row_when"] = st.text_input(
                        label="the cell has (leave empty for empty cell):",
                        value=f.get("omit_row_when", ""),
                        placeholder="e.g.: Sidenote",
                        key=f"omit_when_{f['id']}",
                        help="* for anything not empty; for more words (or anything) separate them by ,/comma or ;/semicolon;"
                    )

                    # Validation message directly under the inputs
                    if not f["targ_col_raw"] and not f["omit_row_when"]:
                        st.warning("Enter at least one input.", icon="⚠️")

                    # Delete button with clean callback passing index
                    st.button(
                        "Delete filter",
                        key=f"del_{f['id']}",
                        use_container_width=True,
                        on_click=delete_filter_callback,
                        args=[f["id"]]
                    )

            # 2. Add Filter Button (Triggers pure state update via callback)
            st.button(
                "➕ Add a filter",
                use_container_width=True,
                on_click=add_filter_callback
            )

            st.divider()

            # 3. Apply Button
            apply_cond_filters = st.button("Apply Conditional Filters")

            if apply_cond_filters:
                # Pre-parse inputs upon final submission
                for f in filters:
                    f["parsed_tCol"] = col_to_num(f["targ_col_raw"])
                    
                st.session_state["steps"][4] = True  # Assuming 0-indexed step 4

            if st.session_state["steps"][4]:
                st.success("✓")
            else:
                st.info("Please confirm selection of conditional filters.")

else:
    st.info("Waiting for completion of step 1 ...")

if st.session_state["steps"][1] and st.session_state["steps"][2] and st.session_state["steps"][3] and st.session_state["steps"][4]:    
    st.success(f"✓ Step 2 completed. Ready for data extraction.")


# SECTION: DATABASE STRUCTURIZING ==========================================


st.divider()

if st.session_state["steps"][4]:

    st.subheader("Step 3: Database Architecture")




else:
    st.info("Waiting for completion of step 3 ...")

if st.session_state["steps"][5]:
    st.success(f"✓ Step 3 completed. Ready to transfer data onto database.")
    