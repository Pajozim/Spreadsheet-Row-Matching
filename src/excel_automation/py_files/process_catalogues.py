import os
import openpyxl
import random
import re
import config as cfg
import streamlit as st
import pandas as pd
from src.excel_automation.py_files.db_related import insert_catalogue_rows
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.cell_range import CellRange
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
from typing import Any

def clean_unit(unit: str) -> str:
    # standardizing units
    preprocessed_unidad = unit.strip().lower()
    cleaned_unidad = preprocessed_unidad if preprocessed_unidad not in cfg.UNIT_ABBREVIATIONS else cfg.UNIT_ABBREVIATIONS[preprocessed_unidad]
    return cleaned_unidad.translate(cfg.UNIT_TRANS)


def clean_description(description: str) -> str:
    # standardizing descriptions
    preprocessed_description: str = description.strip().lower()
    cleaned_description: list = re.findall(r'\w+', preprocessed_description)
    for idx, item in enumerate(cleaned_description):
        if item in cfg.ABBREVIATIONS:
            cleaned_description[idx] = cfg.ABBREVIATIONS[item]

    return " ".join(cleaned_description)



        #print(f"Loaded: \n {cat_path} \n")


def process_catalogues(list_of_catalogues: list | None, db_path: Path | str | None = cfg.DB_PATH) -> None:

    if not list_of_catalogues or db_path is None:
        print("Error: list_of_catalogues and/or db_path is None")
        return

    print(f"Loading catalogue data from {list_of_catalogues} ... \n")

    # Initialize variables
    catalogues_data: list = []
    
    # Load catalogues workbooks into memory
    for cat_path in list_of_catalogues:
        if not os.path.exists(cat_path):
            print(f"Error: Catalogue file '{cat_path}' not found!")
            return
            
        wb_cat = openpyxl.load_workbook(cat_path, data_only=True)
        #print(f"Loaded: \n {cat_path} \n")

        for sheet in wb_cat.worksheets:
            #print(f"\n Processing sheet: {sheet.title}")

            sheetslist = ["Subcontratos", "Catalogo General"]
            if not sheet.title.isdigit() and sheet.title not in sheetslist: continue 

            # Ignores hidden or leftover sheets
            if sheet.sheet_state != 'visible': continue

            # Check cell A2 directly before iterating over rows
            if sheet["A2"].value is None:
                #print(f"Skipping sheet '{sheet.title}': A2 is empty.")
                continue  # Skip to the next sheet immediately
            
            # Do NOT use values_only=True so we can access cell.coordinate
            for row in sheet.iter_rows(min_row=2, max_col=7):

                # if row[6].value in sheet Subcontratos has "NO USAR - CANCELADO"
                if row[6].value == "NO USAR - CANCELADO": continue

                # Extract values safely
                vals = [c.value for c in row]

                # first conditional: Skip if ANY of the first 5 cells (indices 0 to 4) are empty or whitespace
                # second conditional: Skip if ANY of the first 3 cells (indices 0 to 2) are not digits
                # third conditional: Skip if the 4th cell (index 3) is empty
                if any(not str(v).strip() for v in vals[0:5]) or any(not str(v).strip().isdigit() for v in vals[0:3]) or vals[3] is None:
                    continue

                # random checks
                #if row[0].row == 10: print("row:", row, "\nvals:", vals)

                if len(str(vals[3])) > 3:
                    # column 4/D, should be MATRICES
                    c0 = f"{vals[0]:0>2}"
                    c1 = f"{vals[1]:0>2}"
                    c2 = f"{vals[2]:0>3}"
                    
                    clave = f"{c0}-{c1}-{c2}"
                    text = f"{clean_description(str(vals[3]))}"
                    unidad = f"{vals[4]}"

                    # Point directly to the exact cell in the catalogue (e.g., Column D / idx 3)
                    text_cell_coord = row[3].coordinate if len(row) > 3 else "A1"
                else:
                    # column 5/E, should be INSUMOS
                    c0 = f"{vals[0]:0>1}"
                    c1 = f"{vals[1]:0>2}"
                    c2 = f"{vals[2]:0>1}"
                    c3 = f"{vals[3]:0>3}"
                    
                    clave = f"{c0}-{c1}-{c2}-{c3}"
                    text = f"{clean_description(str(vals[4]))}"
                    unidad = f"{vals[5]}"

                    # Point directly to the exact cell in the catalogue (e.g., Column D / idx 3)
                    text_cell_coord = row[4].coordinate if len(row) > 3 else "A1"

                # packing values into the payload
                catalogues_data.append({
                    'clave': clave,
                    'text': text.strip().lower(),
                    'unidad': clean_unit(unidad),
                    'hoja': sheet.title.strip().lower(),
                    'cell_ref': f"='[{os.path.basename(cat_path)}]{sheet.title}'!{text_cell_coord}",
                    'file': os.path.basename(cat_path)
                })

    # Save data to database
    insert_catalogue_rows(catalogues_data, ["clave", "text", "unidad", "hoja", "cell_ref", "file"])

    # Print paremeter results
    #print(f"Loaded {len(catalogues_data)} catalogue entries. \n\n")
    #print(f"Random Catalogue entry: {catalogues_data[random.randint(0, len(catalogues_data)-1)]}. \n\n")
    #print(f"Control set: {control_set}")

    return


def full_ctlg_process(list_of_catalogues: list | None, cdtl_prms: dict) -> list[dict] | None:

    if not list_of_catalogues:
        print("Error: catalogue/s path/s is/are missing.")
        return

    print(f"Loading catalogue data from {list_of_catalogues} ... \n")

    # Initialize variables
    catalogues_data: list = []
    
    # Load catalogues workbooks into memory
    for cat_path in list_of_catalogues:
        if not os.path.exists(cat_path):
            print(f"Error: Catalogue file '{cat_path}' not found!")
            return
            
        wb_cat = openpyxl.load_workbook(cat_path, data_only=True)
        #print(f"Loaded: \n {cat_path} \n")

        for sheet in wb_cat.worksheets:        

            # Do NOT use values_only=True so we can access cell.coordinate
            for row in sheet.iter_rows(min_row=cdtl_prms["min_row"], max_row=cdtl_prms["max_row"], min_col=cdtl_prms["min_col"], max_col=cdtl_prms["max_col"]):

                # Skip empty rows, including those containing (accidental or not) whitespace/s
                if any(not str(v).strip() for v in row): continue

                # Notes:
                # when user left conditional cell "", it is returned as "", but falsy anway
                for col_List, cellV in cdtl_prms["filters"]:
                    # Skip if cell has substring of cellV
                    if any(cellV in str(row[c].value.strip() or "") for c in col_List): continue
                
                # Extract values safely
                vals = [c.value for c in row]

                # packing values
                catalogues_data.append({
                    'cell_ref': f"='[{os.path.basename(cat_path)}]{sheet.title}'!{row[0].coordinate}",
                    'file': os.path.basename(cat_path)
                })


    return catalogues_data


def process_ctlg_files(list_of_catalogues: list[str] | None) -> list[tuple]:
    if not list_of_catalogues:
        return []

    # Initialize variables
    workbooks: list[tuple] = []
    
    # Load catalogues workbooks into memory
    for cat_path in list_of_catalogues:
        if not os.path.exists(cat_path):
            return []
            
        wbk_and_path: tuple = (openpyxl.load_workbook(cat_path, read_only=True, data_only=True), Path(cat_path))
        workbooks.append(wbk_and_path)

    return workbooks


def inspect_sheet(worksheet: Worksheet) -> Worksheet | None:
    if not worksheet or worksheet.sheet_state != 'visible':
        return None
    return worksheet


def process_workbooks(workbooks: list[tuple] | None) -> list[dict[str, list[Worksheet]]] | None:
    if not workbooks:
        return None
    
    sheets_list: list[dict[str, list[Worksheet]]] = [] # intended format: [{"path1": ["sheet1", "sheet2", ...]}, {"path2": ["sheet1", "sheet2", ...]}, ...]

    for wb, path in workbooks:
        path_strfd = str(path)
        lof_on_path = {path_strfd: []} # list of sheets on path

        for sheet in wb.worksheets:
            result: Worksheet | None = inspect_sheet(sheet)
            if result is not None:
                lof_on_path[path_strfd].append(result)
        sheets_list.append(lof_on_path)


    return sheets_list
        
"""
def get_actual_max_row(sheet: Worksheet) -> int:
    for row_num in range(sheet.max_row, 0, -1):
        if any(cell.value is not None for cell in sheet[row_num]):
            return row_num
    return 1


def random_3rows(slctd_sheets: list[dict[str, list[Worksheet]]]) -> pd.DataFrame:
    # Extract 3 random rows from selected sheets into a clean Pandas DataFrame.
    sp = st.session_state["selected_parameters"]

    temp_sheet_list: list[Worksheet] = []
    for wb_dict in slctd_sheets:
        for sheetlist in wb_dict.values():
            temp_sheet_list.extend(sheetlist)

    rndm_3sheets: list[Worksheet] = random.choices(temp_sheet_list, k=3) if len(temp_sheet_list) >= 3 else temp_sheet_list

    rows_data: list[list] = []
    for sheet in rndm_3sheets:
        rndm_row = random.randint(sp["min_row"], sheet.max_row)
        rows_data.append(
            [row[0].value for row in sheet.iter_rows(min_row=rndm_row, max_row=rndm_row)]
        )

    # Generate default column names (Col A, Col B, etc. or index-based)
    cols_count = sp["max_col"]
    col_names: list[str] = [f"Col{i+1}" for i in range(cols_count)]
    
    return pd.DataFrame(rows_data, columns=col_names)       
    


def process_sheets(sheets: list[Worksheet], cdtl_prms: dict) -> list[dict] | None:

    catalogues_data: list[dict] = []

    for sheet in sheets:        

        # Do NOT use values_only=True so we can access cell.coordinate
        for row in sheet.iter_rows(min_row=cdtl_prms["min_row"], max_row=cdtl_prms["max_row"], min_col=cdtl_prms["min_col"], max_col=cdtl_prms["max_col"]):

            # Skip empty rows, including those containing (accidental or not) whitespace/s
            if any(not str(v).strip() for v in row): continue

            # Notes:
            # when user left conditional cell "", it is returned as "", but falsy anway
            for col_List, cellV in cdtl_prms["filters"]:
                # Skip if cell has substring of cellV
                if any(cellV in str(row[c].value.strip() or "") for c in col_List): continue
            
            # Extract values safely
            vals = [c.value for c in row]

            # packing values

            catalogues_data.append({
                'cell_ref': f"='[{os.path.basename(cat_path)}]{sheet.title}'!{row[0].coordinate}",
                'file': os.path.basename(cat_path)
            })

    return        

"""

import openpyxl
import pandas as pd
from typing import Generator, Any

def process_ctlg_pipeline(file_paths: list[str]) -> Generator[Any, Any, str]:
    """
    Single function handling the entire lifecycle.
    Uses 'yield' to pause execution and wait for user input from Streamlit.
    """
    # Load workbooks once — keep them open across yields
    workbooks = [
        (openpyxl.load_workbook(fp, read_only=True, data_only=True), fp) 
        for fp in file_paths
    ]

    """
    Template:
    cData = {
        "PATH1": {
            "selected_sheets": ["Sheet1", "Sheet2", ...],
            "min_col": 1,
            "max_col": None,
            "min_row": 1,
            "max_row": None,
            "filters": []
        },
        "PATH2": {
            ...
        },
        ...
    }
    """
    cData = {}

    try:
        # --- Discover Sheets ---
        found_sheets = {}    
        for wb, path in workbooks:
            found_sheets[path] = [s.title for s in wb.worksheets]
            cData[path] = {}

        # PAUSE 1: Yield discovered sheets to UI, WAIT for user's selection
        selected_sheets = yield {"step": "SELECT_SHEETS", "data": found_sheets}

        # --- Use of filters ---
        payload = {}
        for path, sheets in selected_sheets.items():
            cData[path]["selected_sheets"] = sheets
            payload[path] = {}
    
        slctd_col_filters = yield {"step": "SELECT_COLUMNS", "data": payload}
        for path in slctd_col_filters:
            cData[path].update(slctd_col_filters[path])
        slctd_row_filters = yield {"step": "SELECT_ROWS", "data": payload}
        for path in slctd_row_filters:
            cData[path].update(slctd_row_filters[path])
        addtl_filters = yield {"step": "SELECT_ADDTL_FILTERS", "data": payload}
        for path in addtl_filters:
            cData[path].update(addtl_filters[path])

        print("cData:\n", cData)

        # PAUSE 2: Yield sample data to UI, WAIT for user confirmation
        user_confirmed = yield {"step": "PREVIEW_DATA", "data": pd.DataFrame()}


        # --- STEP 3: Clone to Database ---
        if user_confirmed:
            # Execute your SQL / DB insertion logic here
            pass

        return "Database migration completed successfully!"

    finally:
        # GUARANTEED CLEANUP: Always closes zip handles when pipeline finishes or fails
        for wb, _ in workbooks:
            wb.close()