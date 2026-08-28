
from pathlib import Path
import streamlit as st
import config as cfg
import pandas as pd

from src.excel_automation.py_files.process_catalogues import process_workbooks, random_3rows
from src.excel_automation.py_files.stlt_helpers import get_xlsx_files, get_cache, col_to_num, add_filter_callback, delete_filter_callback, add_merge_group, remove_merge_group 

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
        "slctd_sheet_titles": [],
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
        height=140,
    )

with col2:
    st.space("small")
    if st.button("🔄 Refresh Directory", use_container_width=True):
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
        if not st.session_state["workbooks"]:
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
    
    if not processed_sheets:
        st.warning("⚠️ No sheets found in selected workbooks.")
        available_sheet_names = []
    else:
        available_sheet_names = [sheet.title for sheet in processed_sheets]

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
                st.session_state["slctd_sheet_titles"] = selected_sheets
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
            apply_cond_filters = st.button("Apply Conditional Filters", use_container_width=True)

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

    # Mock sample dataframe (representing the output of random_3rows)
    sample_df: pd.DataFrame = random_3rows(st.session_state["slctd_sheet_titles"])

    # --- 1. SAMPLE ROWS DISPLAY ---
    st.subheader("1. Sample 3 Rows Preview")
    st.caption("Representation of 3 random rows picked from your selected worksheet(s):")

    # Display raw sampled data
    st.dataframe(sample_df, use_container_width=True)

    st.divider()

    # --- 2. CLICK-TO-RENAME COLUMNS ---
    st.subheader("2. Rename Column Headers")
    cols = list(sample_df.columns)

    # Render side-by-side rename inputs for each column
    rename_cols = st.columns(len(cols))
    for idx, col in enumerate(cols):
        with rename_cols[idx]:
            new_name = st.text_input(
                label=f"Rename `{col}`",
                value=st.session_state["column_renames"].get(col, col),
                key=f"rename_{col}"
            )
            st.session_state["column_renames"][col] = new_name

    st.divider()

    # --- 3. MERGE / UN-MERGE BLOCK MANAGER ---
    st.subheader("3. Merge Columns")

    c1, c2 = st.columns([1, 1], gap="large")

    with c1:
        st.markdown("#### Create New Merged Group")
        
        # Available columns to pick from
        selected_to_merge = st.multiselect(
            "Select columns you want to merge:",
            options=cols,
            format_func=lambda x: f"{x} ({st.session_state['column_renames'].get(x, x)})"
        )

        delim_choice = st.selectbox(
            "Select or type delimiter:",
            options=[" Space ( ' ' )", "Hyphen ( - )", "Underscore ( _ )", "Comma ( , )",  "No Delimiter", "Custom"]
        )
        
        actual_delimiter = ""
        if delim_choice == " Space ( ' ' )": actual_delimiter = " "
        elif delim_choice == "Hyphen ( - )": actual_delimiter = "-"
        elif delim_choice == "Underscore ( _ )": actual_delimiter = "_"
        elif delim_choice == "Comma ( , )": actual_delimiter = ", "
        elif delim_choice == "No Delimiter": actual_delimiter = ""
        elif delim_choice == "Custom":
            actual_delimiter = st.text_input("Custom Delimiter Symbol:", value="-")

        merged_col_name = st.text_input("Merged Column Target Name (optional):", placeholder="e.g. Full_Address")

        if st.button("🔗 Merge Selected Columns", use_container_width=True):
            if len(selected_to_merge) < 2:
                st.warning("Please select at least 2 columns to merge.")
            else:
                add_merge_group(selected_to_merge, actual_delimiter, merged_col_name)
                st.rerun()

    with c2:
        st.markdown("#### Active Merged Groups (Un-merge)")
        
        if not st.session_state["merged_groups"]:
            st.info("No columns merged yet.")
        else:
            for group in st.session_state["merged_groups"]:
                with st.container(border=True):
                    g_col1, g_col2 = st.columns([3, 1])
                    with g_col1:
                        joined_cols_str = " + ".join([st.session_state["column_renames"].get(c, c) for c in group["cols"]])
                        delim_display = f"'{group['delimiter']}'" if group['delimiter'] != "" else "None"
                        st.markdown(f"**Target:** `{group['name']}`")
                        st.caption(f"**Merging:** {joined_cols_str} | **Delimiter:** {delim_display}")
                    
                    with g_col2:
                        st.button(
                            "🔓 Un-merge",
                            key=f"unmerge_{group['id']}",
                            on_click=remove_merge_group,
                            args=[group["id"]]
                        )

    st.divider()

    # --- 4. FINAL TRANSFORMED PREVIEW ---
    st.subheader("4. Final Transformed Preview")

    transformed_df = pd.DataFrame()

    # 1. First apply renames
    for c in sample_df.columns:
        renamed_title = st.session_state["column_renames"].get(c, c)
        transformed_df[renamed_title] = sample_df[c]

    # 2. Apply Merges
    for group in st.session_state["merged_groups"]:
        target_name = group["name"]
        cols_to_join = group["cols"]
        delim = group["delimiter"]
        
        # Concatenate series
        joined_series = sample_df[cols_to_join].astype(str).agg(delim.join, axis=1)
        transformed_df[target_name] = joined_series
        
        # Optionally drop original columns that were merged
        # transformed_df.drop(columns=[st.session_state["column_renames"].get(c, c) for c in cols_to_join], inplace=True)

    st.dataframe(transformed_df, use_container_width=True)












else:
    st.info("Waiting for completion of step 3 ...")

if st.session_state["steps"][5]:
    st.success(f"✓ Step 3 completed. Ready to transfer data onto database.")
    