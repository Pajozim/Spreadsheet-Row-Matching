import streamlit as st
import config as cfg
import pandas as pd

from io import StringIO
from pathlib import Path
from src.excel_automation.py_files.process_files import process_spreadsheet_files
from src.excel_automation.py_files.semantic_search import prep_batch, embed_catalogues
from src.excel_automation.py_files.db_related import init_db, get_db_columns, show_db
from src.excel_automation.py_files.stlt_helpers import get_xlsx_files, refresh_callback, col_to_num, add_filter_callback, delete_filter_callback, add_category_callback, delete_category_callback, get_last_column, get_letters, list_sheets


print("Starting Streamlit ... \n")


# ==========================================
# SECTION: PAGE CONFIG (MUST BE FIRST STREAMLIT CALL)
# ==========================================
st.set_page_config(page_title="Catalogue Matcher", layout="wide")


# ==========================================
# SECTION: SESSION_STATE VISUALIZATION
# ==========================================


if "paths_count" not in st.session_state:
    st.session_state["paths_count"] = 0
if "route" not in st.session_state:
    st.session_state["route"] = [False, False, False, False, False, False, False, False, False] # current length: 9
    
if "embedding_error" not in st.session_state:
    st.session_state.embedding_error = None
if "is_catalog_ready" not in st.session_state:
    st.session_state.is_catalog_ready = False


# SECTION: ACTION MENU ==================================================================

st.title("📋 Retrieval System from Spreadsheet Files (Excel)")

col1, col2 = st.columns([1, 3])

active_key = "c-parameters" if "c-parameters" in st.session_state else "s-parameters"
if active_key not in st.session_state:
    st.session_state[active_key] = {}
cur_prmtrs = st.session_state[active_key]

with st.container(horizontal=True, vertical_alignment="top"):
    with col1:
        st.write("Add data or intialize the database for the Catalogue Matcher?")
        action = st.menu_button("Which action?", ["Create, update or add more data to the database", "Search on the database"])
        if action == "Create, update or add more data to the database":
            st.session_state["route"][0] = True
            for sst in st.session_state["route"][1:]:
                sst = False
            active_key = "c-parameters"
            if "s-parameters" in st.session_state:
                del st.session_state["s-parameters"]
            if active_key not in st.session_state:
                st.session_state[active_key] = {}
            cur_prmtrs = st.session_state[active_key]
        elif action == "Search on the database":
            
            try:
                # When this runs, if it fails, it throws an exception handled right here
                prep_batch()
                embed_catalogues()
                st.session_state.is_catalog_ready = True
                st.session_state.embedding_error = None
            except Exception as e:
                st.session_state.is_catalog_ready = False
                st.session_state.embedding_error = str(e)
            
            st.session_state["route"][0] = True
            for sst in st.session_state["route"][1:]:
                sst = False
            active_key = "s-parameters"
            if "c-parameters" in st.session_state:
                del st.session_state["c-parameters"]
                if active_key not in st.session_state:
                    st.session_state[active_key] = {}
            cur_prmtrs = st.session_state[active_key]

    with col2:
        st.caption("Database preview:")
        st.write(pd.DataFrame(show_db("random")[:5], columns=get_db_columns("catalogue_items")))


# SECTION: SHARED PARAMETER SELECTION ===================================================

if st.session_state["route"][0]: # files paths

    st.divider() # ----------------------------------------------------------------------------

    st.subheader("Instructions before Usage")
    st.markdown("""
        - Go from your "root folder" (where you opened this app) to `src/excel_automation/xlsx_files/` folder.
        - Copy & Paste (or drag after duplicating) your cataloguing/lookup files AND search-consulting files to the `xlsx_files/` folder.
        - Ready? Continue on here.
    """)

    # SECTION: FILE PATHS ==========================================

    st.subheader("Step 1: File Reader & Paths")
    if active_key == "c-parameters":
         st.caption("""Select the spreadsheet files whose data will serve as the search database (data you are inserting INTO the database).
    \n ⚠️ Do NOT select the Excel/spreadsheet files you are reading FROM, whose read data is to be searched on the database. 
    If still unclear, please consult README.md.""")
    else:
        st.caption("""Select the spreadsheet file whose data you are reading FROM, whose read data is to be searched on the database.
    \n ⚠️ Do NOT select the Excel/spreadsheet files you are inserting INTO the database. 
    If still unclear, please consult README.md.""")

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
                for path in cur_prmtrs.keys():
                    if path not in file_paths:
                        del cur_prmtrs[path]
                for file_path in file_paths:
                    cur_prmtrs[file_path] = {
                        "selected_sheets": [],
                        "min_row": 1,
                        "max_row": None,
                        "min_col": 1,
                        "max_col": None,
                        "filters": [],

                    }
                list_sheets(cur_prmtrs)
                st.session_state["route"][1] = True
                st.rerun()

    if not st.session_state["detected_files"]:
        st.info(f"No files found in '{cfg.XLSX_DIR}/'. Copy and paste files there and click Refresh.")
    elif st.session_state["route"][1]:
        st.success(f"✓")
    else:
        st.info('Please Ctrl + Enter inside text area when finished selecting file paths.')

if st.session_state["route"][1]: # sheets filter 
    st.divider() # ----------------------------------------------------------------------------

    st.subheader("Step 2: Sheets Filter")

    with st.form(key="sheets_filter_form"):
        st.text("Select sheet/s to process:")

        for file_path, pDict in cur_prmtrs.items():
            file_name = Path(file_path).stem
            widget_key = f"sheet_selector_{file_name}"
            sheets = pDict.get("selected_sheets", [])
            selected_sheets = st.multiselect(
                label=f"File: **{file_name}**",
                options=sheets,
                default=sheets,
                key=widget_key
            )
            pDict["selected_sheets"] = selected_sheets 

        apply_sht_filter = st.form_submit_button(label="Apply Sheets Filter")

    if apply_sht_filter:
        st.session_state["route"][2] = True
        st.rerun()
    elif st.session_state["route"][2]:
        st.success("✓")
    else:
        st.info('Please confirm selection of sheets by clicking on the "Apply Sheets Filter" button.')

if st.session_state["route"][2]: # columns filter

    st.divider() # ----------------------------------------------------------------------------

    st.subheader("Step 3: Columns Filter")

    with st.form(key="columns_filter_form"): 
        st.text("Select column/s to process:")

        for file_path, pDict in cur_prmtrs.items():
            widget_key = f"columns_selector_{Path(file_path).stem}"
            st.caption(f"File: **{Path(file_path).stem}**")
            col_from = st.text_input("From column:", placeholder="e.g. B", value=st.session_state.get(f"{widget_key}-cf", ""), key=f"{widget_key}-cf")
            col_to = st.text_input("To column (empty for all of them; or for single column, write letter again):", value=st.session_state.get(f"{widget_key}-ct", ""), key=f"{widget_key}-ct")
            if col_from is None:
                st.warning("Please provide a column.")
            elif col_from or col_to:
                cur_prmtrs[file_path].update({"min_col": col_to_num(col_from, 1), "max_col": col_to_num(col_to, 1) if col_to else None})

        apply_col_filter = st.form_submit_button(label="Apply Columns Filter")

        if apply_col_filter:
            st.session_state["route"][3] = True
            st.rerun()
        elif st.session_state["route"][3]:
            st.success("✓")
        else:
            st.info('Please confirm selection of columns.')

if st.session_state["route"][3]: # rows filter

    st.divider() # ----------------------------------------------------------------------------

    st.subheader("Step 4: Rows Filter")

    with st.form(key="rows_filter_form"): 
        st.text("Select row/s to process:")

        for file_path, pDict in cur_prmtrs.items():
            widget_key = f"rows_selector_{Path(file_path).stem}"
            st.caption(f"File: **{Path(file_path).stem}**")
            row_from = st.text_input("From row:", placeholder="e.g. 2", value=st.session_state.get(f"{widget_key}-rf", ""), key=f"{widget_key}-rf")
            row_to = st.text_input("To row (empty for all of them; when only one row, write number again):", value=st.session_state.get(f"{widget_key}-rt", ""), key=f"{widget_key}-rt")
            if row_from:
                cur_prmtrs[file_path].update({"min_row": int(row_from), "max_row": int(row_to) if row_to else None})
            else:
                st.warning("Please provide - at least - the first row number.")

        apply_row_filter = st.form_submit_button(label="Apply Rows Filter")

        if apply_row_filter:
            st.session_state["route"][4] = True
            st.rerun()
        elif st.session_state["route"][4]:
            st.success("✓")
        else:
            st.info('Please confirm selection of rows.')

if st.session_state["route"][4]: # additional filters

    st.divider() # ----------------------------------------------------------------------------

    st.subheader("Step 4: Additional Filters")
    st.write("Note: Any empty rows are automatically omitted (including those with accidental whitespaces).")

    for path, pDict in cur_prmtrs.items():
        file_stem = Path(path).stem
        st.caption(f"File: **{file_stem}**")

        for idx, f in enumerate(pDict["filters"]):
            with st.container(border=True):
                st.caption(f"Filter #{idx + 1}")

                f["targ_col_raw"] = st.text_input(
                    label="Omit row when in column (leave empty for all columns):",
                    value=f.get("targ_col_raw", ""),
                    placeholder="e.g.: B",
                    key=f"targ_col_{f['id']}",
                    help="Write 'any(column, column, ...)' when one of these columns shall fulfill the condition, 'all(column, column, ...)' if all of these columns shall be targeted, or single letters."
                )

                f["wordpiece"] = st.text_input(
                    label="the cell has (leave empty for empty cell):",
                    value=f.get("wordpiece", ""),
                    placeholder="e.g.: Sidenote",
                    key=f"omit_when_{f['id']}",
                    help="Write '*'/asterisk for anything NOT empty or for partial words (e.g.: num* or *33); separate multiple entries with commas or semicolons."
                )

                if not f["targ_col_raw"] and not f["wordpiece"]:
                    st.warning("Enter at least one input.", icon="⚠️")
                elif len(f["targ_col_raw"]) > 1 and not f["targ_col_raw"].startswith("a"):
                    st.warning("Please use 'any(column, column, ...)' or 'all(column, column, ...)' for multiple columns.", icon="⚠️")

                st.button(
                    "🗑️ Delete filter",
                    on_click=delete_filter_callback,
                    args=[f["id"], cur_prmtrs[path]],
                    key=f"del_{f['id']}"
                )

        # "Add a filter" button for this specific file path
        st.button(
            "➕ Add a filter",
            on_click=add_filter_callback,
            args=[cur_prmtrs[path]],
            key=f"add_{path}"  # Safe key using file path
        )

        st.divider()

    # 2. Outer Form reserved specifically for final submission
    with st.form(key="filters_apply_form"):
        apply_cond_filters = st.form_submit_button("Apply Conditional Filters")

        if apply_cond_filters:
            if active_key == "c-parameters":
                st.session_state["route"][5] = True
            else:
                st.session_state["route"][7] = True
            st.rerun()
        elif active_key == "c-parameters" and st.session_state["route"][5] or active_key == "s-parameters" and st.session_state["route"][7]:
            st.success("✓")
        else:
            st.info('Please confirm selection of filters.')


# SECTION: CATALOGUEING INTO DATABASE ===================================================

if st.session_state["route"][5] and active_key == "c-parameters": # database initialization
    st.divider() # ----------------------------------------------------------------------------

    st.subheader("Step 5: Database Initialization")
    st.write("Structurize the template of the database")

    with st.container(border=True, gap="small"):
        st.text("Existing table columns:")
        st.write(" | ".join(ti for ti in get_db_columns("catalogue_items")))

    skip_db_init = st.checkbox("**Skip Database Initialization? That is, data from this session is to be added to the existing database.**", value=False, key="skip_db_init")
    if not skip_db_init:
        st.caption("💡 **Formatting Rule:** Use single words without spaces for identifiers (e.g., use `file_path`, not `file path`).")

        st.caption("""
        🔍 **About Column Indexing:** 
        Enable indexing on columns/categories you frequently search, filter, or sort by.
        > * **Best candidates:** High-variety values (at least 1 unique value per 10–20 rows), such as *Emails, Zip Codes, Serial Numbers, Cities,* or *IDs*.
        > * **Optimal text length:** Short to medium entries (under 150 characters).
        > * **Avoid indexing:** Low-variety columns (e.g., *Status, Yes/No, Gender*), as they increase file size without improving search speed.
        """)

        st.caption("""
        ⚡ **About Full-Text Search (FTS5):** 
        Enable FTS5 for columns with long, unstructured, or highly variable text (e.g., *Descriptions, Remarks, Instructions, Manuals*). This creates a search engine-style index for fast keyword matching.
        """)

        st.info("Sheets and cell references resp. cell coordinates are automatically included.")

        if "db_template" not in st.session_state:
            st.session_state["db_template"] = []
        if "db_addtnl" not in st.session_state:
            st.session_state["db_addtnl"] = {"fts5": [], "index": []}

        with st.container(border=True, width="content", horizontal=True, vertical_alignment="center"):
            for idx, item in enumerate(st.session_state["db_template"]):
                with st.container(border=True, width="content"):
                    st.caption(f"#{idx + 1}")

                    indexing = st.toggle(
                        label="index this category",
                        key=f"indexing_{item['id']}",
                        value=False
                    )
                    if indexing and item["name"] not in st.session_state["db_addtnl"]["index"]:
                        st.session_state["db_addtnl"]["index"].append(item["name"])
                    elif not indexing and item["name"] in st.session_state["db_addtnl"]["index"]:
                        st.session_state["db_addtnl"]["index"].remove(item["name"])

                    fts5 = st.toggle(
                        label="enable FTS5 for this category",
                        key=f"fts5_{item['id']}",
                        value=False
                    )
                    if fts5 and item["name"] not in st.session_state["db_addtnl"]["fts5"]:
                        st.session_state["db_addtnl"]["fts5"].append(item["name"])
                    elif not fts5 and item["name"] in st.session_state["db_addtnl"]["fts5"]:
                        st.session_state["db_addtnl"]["fts5"].remove(item["name"])

                    cat_name = st.text_input(
                        label="Category Name:",
                        value=item["name"],
                        placeholder="e.g.: Keycode",
                        key=f"cat_name_{item['id']}"
                    )

                    st.button(
                        "🗑️Delete Category",
                        on_click=delete_category_callback,
                        args=[item["id"]],
                        key=f"del_{item['id']}"
                    )
                    item["name"] = cat_name

            st.button(
                "➕ Add a category",
                on_click=add_category_callback,
            )

        with st.form(key="db_init_form"):
            apply_db_init = st.form_submit_button("Apply Database Initialization")

            if apply_db_init and st.session_state["db_template"]:
                st.session_state["route"][6] = True
                with st.spinner("Initializing database..."):
                    init_db(st.session_state["db_template"], st.session_state["db_addtnl"])
                st.rerun()
            elif apply_db_init and not st.session_state["db_template"]:
                st.error("Please add at least one category OR skip database initialization.")
            elif st.session_state["route"][6]:
                st.success("✓")
            else:
                st.info('Please confirm database template.')
    else:
        st.session_state["route"][6] = True
        st.success("✓")    

if st.session_state["route"][6] and active_key == "c-parameters": # transfer Coordination

    st.divider() # ----------------------------------------------------------------------------

    st.subheader("Step 6: Transfer Coordination")
    st.caption(
        "Extracting and transferring spreadsheet data to the database. Please select the spreadsheet columns to transfer its content to which database categories."
        )

    db_catgrs = st.session_state.get("db_template", [])


    for path in cur_prmtrs:

        # path shortened/stemmed
        fPath = Path(path).stem
        fDPath = cur_prmtrs[path]

        # getting last/max column and converting column numbers into letters 
        if cur_prmtrs[path].get("max_col") is None:
            fDPath["max_col"] = get_last_column(path)
        available_columns: list = get_letters(fDPath["min_col"], fDPath["max_col"])

        st.markdown(f"### 📄 File: `{fPath}`")

        # Re-initializing mappings
        column_mappings = []

        # loop through DB Categories and ask user to map file columns
        with st.expander(f"Configure Mappings for {fPath}", expanded=True):
            for db_cat in db_catgrs:
                # visualization of db_cat - e.g.: {'id': '98d98e04', 'name': 'clave'}
                col1, col2 = st.columns([3, 2])

                with col1:
                    selected_cols = st.multiselect(
                        label=f"Spreadsheet column(s) for `{db_cat['name']}`:",
                        options=available_columns,
                        key=f"cols_{path}_{db_cat['id']}"
                    )

                delimiter = ""
                with col2:
                    # Only show delimiter if more than 1 column is mapped to this DB category
                    if len(selected_cols) > 1:
                        delim_choice = st.selectbox(
                            label=f"Delimiter for `{db_cat['name']}`:",
                            options=["Space (' ')", "Plus ('+')", "Hyphen ('-')", "Custom"],
                            key=f"delim_choice_{path}_{db_cat['id']}"
                        )

                        if delim_choice == "Custom":
                            delimiter = st.text_input(
                                label="Enter custom delimiter:",
                                placeholder="e.g.: |",
                                key=f"custom_delim_{path}_{db_cat['id']}"
                            )
                        else:
                            delim_map = {
                                "Space (' ')": " ",
                                "Plus ('+')": "+",
                                "Hyphen ('-')": "-"
                            }
                            delimiter = delim_map[delim_choice]

                    if delimiter and selected_cols:
                        preview = delimiter.join(f"(content of {col})" for col in selected_cols) 
                        st.caption(f"Visualization: {preview}")

                # Save choice to "column_mappings"
                if selected_cols:
                    mapping_entry: dict = {
                        db_cat["name"]: [col_to_num(col, 0) for col in selected_cols]                        
                    }
                    if delimiter:
                        mapping_entry.update({"delimiter": delimiter})

                    column_mappings.append(mapping_entry)
                    
        cur_prmtrs[path].update({"column_mappings": column_mappings})

    with st.form(key="transfer_coord_form"):
        transfer_coord = st.form_submit_button(label="Transfer Data")

    result, success = False, ""
    if transfer_coord:
        with st.spinner("Transferring data...", show_time=True):
            result, success = process_spreadsheet_files("catalogue", cur_prmtrs)
            if result:
                st.success("✓")
                st.write(success)
                st.write("The transfer is complete. The database is ready for search. For search go to menu and select 'Search on the database'.")
            st.session_state["route"][8] = True
    elif not st.session_state["route"][8]:
        st.info('Please confirm transfer of data.')
    else:
        st.write("The transfer is complete. The database is ready for search. For search go to menu and select 'Search on the database'.")
    

# SECTION: SEARCHING ON THE DB ==========================================================

if st.session_state["route"][7] and active_key == "s-parameters": # hint for duplicating file, setting result population column and starting search
    st.divider() # ----------------------------------------------------------------------------

    st.subheader("Select the column on which the results shall be populated.")

    with st.form(key="slctn_of_result_col"):
        for path in cur_prmtrs:
            widget_key=f"result_col_slctn_{Path(path).stem}"
            st.caption(f"File: **{Path(path).stem}**")
            chosen_Rcol = st.text_input("choose column:", placeholder="e.g. N", value=st.session_state.get(f"{widget_key}_rc", ""), key=f"{widget_key}_rc")
            if chosen_Rcol:
                st.success("✓")
                for path in cur_prmtrs:
                    cur_prmtrs[path].update({"results_col": chosen_Rcol})
            elif cur_prmtrs[path].get("result_col") is None:
                st.warning("Please provide a column.")

        checked = st.checkbox(label="Source file is duplicated to avoid overwriting the original file.", value=False, key="duplicate_source")
        apply_Rcol = st.form_submit_button(label="Apply Column and start the search.")

    if apply_Rcol and checked:
        with st.spinner(text="working ...", show_time=True):
            end_of_process, success = process_spreadsheet_files("search", cur_prmtrs)
            if end_of_process: 
                st.success("✓")
                st.write(success)
                st.write("The search is complete. Check on the file.")
                st.session_state["route"][8] = True
            else:
                st.warning("Something went wrong.")
    elif not st.session_state["route"][8]:
        st.info('Please confirm selection of column by clicking on the "Apply Column ..." button and checked on the hint.')
    else:
        st.write("The search is complete. Check on the file. For a new search go to menu and select 'Search on the database' or 'Create, update or add more data to the database'.")


# SECTION: Error panel ==================================================================

if st.session_state.get("embedding_error"):
    with st.expander("⚠️ Catalog Embedding Failed (Click for details)", expanded=True):
        st.error(st.session_state.embedding_error)
        
        if st.button("Retry Embedding"):
            st.session_state.embedding_error = None
            # Trigger your embedding function here
            st.rerun()