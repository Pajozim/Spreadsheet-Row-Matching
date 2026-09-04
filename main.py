import streamlit as st
import config as cfg
from pathlib import Path
from src.excel_automation.py_files.process_catalogues import extract_sheets, process_spdsht_files
from src.excel_automation.py_files.db_related import init_db
from src.excel_automation.py_files.stlt_helpers import get_xlsx_files, refresh_callback, col_to_num, add_filter_callback, delete_filter_callback, add_category_callback, delete_category_callback, get_last_column, get_letters


print("Starting Streamlit ... \n")


# ==========================================
# SECTION: PAGE CONFIG (MUST BE FIRST STREAMLIT CALL)
# ==========================================
st.set_page_config(page_title="Catalogue Matcher", layout="wide")


# ==========================================
# SECTION: SESSION_STATE VISUALIZATION
# ==========================================

if "parameters" not in st.session_state:
    st.session_state["parameters"] = {}
if "paths_count" not in st.session_state:
    st.session_state["paths_count"] = 0
if "route" not in st.session_state:
    st.session_state["route"] = [False, False, False, False, False, False, False]


# SECTION: INSTRUCTION FOR "FILE PATHS" =========================================


st.title("📋 Retrieval System on Spreadsheet Files (Excel)")

st.subheader("Instructions before Usage")
st.markdown("""
    - Go from your "root folder" (where you opened this app) to `src/excel_automation/xlsx_files/` folder 
    - Copy & Paste (or drag) your cataloguing/lookup files and your search-consulting file to the `xlsx_files/` folder.
    - Ready? Continue on here.
""")

st.divider() # ----------------------------------------------------------------------------


# SECTION: FILE PATHS ==========================================


st.subheader("Step 1: File Reader & Paths")
st.caption("""Select the spreadsheet files whose data will serve as the search database (data you are inserting INTO the database).
⚠️ Do not select the Excel/spreadsheet files you are reading FROM, whose read data is to be searched on the database. 
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
            for file_path in file_paths:
                st.session_state["parameters"][file_path] = {
                    "selected_sheets": [],
                    "min_row": 1,
                    "max_row": None,
                    "min_col": 1,
                    "max_col": None,
                    "filters": [],

                }
            extract_sheets()
            st.session_state["route"][0] = True
            st.rerun()

if not st.session_state["detected_files"]:
    st.info(f"No files found in '{cfg.XLSX_DIR}/'. Copy and paste files there and click Refresh.")
elif st.session_state["route"][0]:
    st.success(f"✓")
else:
    st.info('Please Ctrl + Enter inside text area when finished selecting file paths.')


st.divider() # ----------------------------------------------------------------------------


# SECTION EXTRACTION PARAMETERS ==========================================

if st.session_state["route"][0]:

    st.subheader("Step 2: Sheets Filter")

    with st.form(key="sheets_filter_form"):
        st.text("Select sheet/s to process:")

        for file_path, pDict in st.session_state["parameters"].items():
            file_name = Path(file_path).stem
            widget_key = f"sheet_selector_{file_name}"
            selected_sheets = st.multiselect(
                label=f"File: **{file_name}**",
                options=pDict["selected_sheets"],
                default=pDict["selected_sheets"],
                key=widget_key
            )
            pDict["selected_sheets"] = selected_sheets 

        apply_sht_filter = st.form_submit_button(label="Apply Sheets Filter")

    if apply_sht_filter:
        st.session_state["route"][1] = True
        st.rerun()
    elif st.session_state["route"][1]:
        st.success("✓")
    else:
        st.info('Please confirm selection of sheets by clicking on the "Apply Sheets Filter" button.')

if st.session_state["route"][1]:

    st.divider() # ----------------------------------------------------------------------------

    st.subheader("Step 3: Columns Filter")

    with st.form(key="columns_filter_form"): 
        st.text("Select column/s to process:")

        for file_path, pDict in st.session_state["parameters"].items():
            widget_key = f"columns_selector_{Path(file_path).stem}"
            st.caption(f"File: **{Path(file_path).stem}**")
            col_from = st.text_input("From cloumn:", placeholder="e.g. B", key=f"{widget_key}-cf")
            col_to = st.text_input("To column (empty for all of them):", "", key=f"{widget_key}-ct")
            if col_from or col_to:
                st.session_state["parameters"][file_path].update({"min_col": col_to_num(col_from), "max_col": col_to_num(col_to) if col_to else None})

        apply_col_filter = st.form_submit_button(label="Apply Columns Filter")

        if apply_col_filter:
            st.session_state["route"][2] = True
            st.rerun()
        elif st.session_state["route"][2]:
            st.success("✓")
        else:
            st.info('Please confirm selection of columns.')

if st.session_state["route"][2]:

    st.divider() # ----------------------------------------------------------------------------

    st.subheader("Step 4: Rows Filter")

    with st.form(key="rows_filter_form"): 
        st.text("Select row/s to process:")

        for file_path, pDict in st.session_state["parameters"].items():
            widget_key = f"rows_selector_{Path(file_path).stem}"
            st.caption(f"File: **{Path(file_path).stem}**")
            row_from = st.text_input("From row:", placeholder="e.g. 2", key=f"{widget_key}-rf")
            row_to = st.text_input("To row (empty for all of them):", "", key=f"{widget_key}-rt")
            if row_from or row_to:
                st.session_state["parameters"][file_path].update({"min_row": int(row_from), "max_row": int(row_to) if row_to else None})
            else:
                st.warning("Please provide at least the first row number.")

        apply_row_filter = st.form_submit_button(label="Apply Rows Filter")

        if apply_row_filter:
            st.session_state["route"][3] = True
            st.rerun()
        elif st.session_state["route"][3]:
            st.success("✓")
        else:
            st.info('Please confirm selection of rows.')


if st.session_state["route"][3]:

    st.divider() # ----------------------------------------------------------------------------

    st.subheader("Step 4: Additional Filters")
    st.write("Note: Any empty rows are automatically omitted (including those with accidental whitespaces).")

    for path, pDict in st.session_state["parameters"].items():
        file_stem = Path(path).stem
        st.caption(f"File: **{file_stem}**")

        for idx, f in enumerate(pDict["filters"]):
            with st.container(border=True):
                st.caption(f"Filter #{idx + 1}")

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
                    help="Write */asterisk for anything NOT empty or for partial words (e.g.: num* or *33); separate multiple entries with commas or semicolons."
                )

                if not f["targ_col_raw"] and not f["omit_row_when"]:
                    st.warning("Enter at least one input.", icon="⚠️")
                else:
                    f["parsed_tCol"] = col_to_num(f["targ_col_raw"])

                st.button(
                    "🗑️ Delete filter",
                    on_click=delete_filter_callback,
                    args=[f["id"], path],
                    key=f"del_{f['id']}"
                )

        # "Add a filter" button for this specific file path
        st.button(
            "➕ Add a filter",
            on_click=add_filter_callback,
            args=[path],
            key=f"add_{path}"  # Safe key using file path
        )

        st.divider()

    # 2. Outer Form reserved specifically for final submission
    with st.form(key="filters_apply_form"):
        apply_cond_filters = st.form_submit_button("Apply Conditional Filters")

        if apply_cond_filters:
            st.session_state["route"][4] = True
            st.rerun()
        elif st.session_state["route"][4]:
            st.success("✓")
        else:
            st.info('Please confirm selection of filters.')


if st.session_state["route"][4]:

    st.divider() # ----------------------------------------------------------------------------

    st.subheader("Step 5: Database Initialization")
    st.write("Structurize the template of the database")

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

            if apply_db_init:
                st.session_state["route"][5] = True
                with st.spinner("Initializing database..."):
                    init_db(st.session_state["db_template"], st.session_state["db_addtnl"])
                st.rerun()
            elif st.session_state["route"][5]:
                st.success("✓")
            else:
                st.info('Please confirm database template.')
    else:
        st.session_state["route"][5] = True
        st.success("✓")    

if st.session_state["route"][5]:

    st.divider() # ----------------------------------------------------------------------------

    st.subheader("Step 6: Transfer Coordination")
    st.caption(
        "Extracting and transferring spreadsheet data to the database. Please select the spreadsheet columns to transfer its content to which database categories."
        )

    db_catgrs = st.session_state.get("db_template", [])

    for path in st.session_state["parameters"]:

        if "column_mappings" not in st.session_state["parameters"][path]:
            st.session_state["parameters"][path]["column_mappings"] = []

        # path shortened/stemmed
        fPath = Path(path).stem

        # getting last/max column and converting column numbers into letters 
        if st.session_state["parameters"][path]["max_col"] is None:
            max_col: int = get_last_column(path)
        fDPath = st.session_state["parameters"][path]
        available_columns: list = get_letters(int(fDPath["min_col"]), int(fDPath["max_col"]))

        st.markdown(f"### 📄 File: `{fPath}`")

        file_mapping = {}
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

                with col2:
                    delimiter = ""
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

                    if delimiter:
                        st.caption(f"Visualization: {delimiter.join("(content of " + col + ")" for col in selected_cols)}")

                # Save choice to dictionary
                file_mapping = {
                    db_cat["name"]: [col_to_num(col) for col in selected_cols],
                    "delimiter": delimiter
                }

                if file_mapping not in st.session_state["parameters"][path]["column_mappings"]: 
                    st.session_state["parameters"][path]["column_mappings"].append(file_mapping)

    with st.form(key="transfer_coord_form"):

        transfer_coord = st.form_submit_button(label="Transfer Data")

    if transfer_coord:
        with st.spinner("Transferring data..."):
            #process_spdsht_files(st.session_state["parameters"], st.session_state["column_mappings"])
            st.session_state["route"][6] = True
            st.rerun()
    elif st.session_state["route"][6]:
        st.success("✓")
    else:
        st.info('Please confirm transfer of data.')

    
