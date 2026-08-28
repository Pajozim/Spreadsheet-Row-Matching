import os
import openpyxl
import random
import re
import config as cfg
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.cell_range import CellRange
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor

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


def process_catalogues(list_of_catalogues: list | None, db_path: Path | str | None = cfg.DB_PATH) -> list[dict] | None:

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
    #insert_catalogue_rows(catalogues_data)

    # Print paremeter results
    #print(f"Loaded {len(catalogues_data)} catalogue entries. \n\n")
    #print(f"Random Catalogue entry: {catalogues_data[random.randint(0, len(catalogues_data)-1)]}. \n\n")
    #print(f"Control set: {control_set}")

    return catalogues_data


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


def process_ctlg_files(list_of_catalogues: list[str] | None) -> str | list[Workbook]:
    if not list_of_catalogues:
        return "Error: Catalogue list is empty."

    print(f"Loading catalogue data from {list_of_catalogues} ... \n")

    # Initialize variables
    workbooks: list[Workbook] = []
    
    # Load catalogues workbooks into memory
    for cat_path in list_of_catalogues:
        if not os.path.exists(cat_path):
            print(f"Error: Catalogue file '{cat_path}' not found!")
            return "Error: Catalogue file not found!"
            
        workbooks.append(openpyxl.load_workbook(cat_path, data_only=True))

    return workbooks


def inspect_sheet(worksheet: Worksheet) -> Worksheet | None:
    if not worksheet or worksheet.sheet_state != 'visible':
        return
    return worksheet


def process_workbooks(workbooks: list[Workbook] | None) -> list[Worksheet] | str:

    if not workbooks:
        return "Error: Workbook is empty."

    sheets_list: list[Worksheet] = []

    # 1. Open the thread pool once for the entire workbook processing lifecycle
    with ThreadPoolExecutor(max_workers=4) as executor:
        for wb in workbooks:
            
            # 2. Map the concurrent worker function over the current workbook's worksheets
            results = executor.map(lambda s: inspect_sheet(s), wb.worksheets)
            
            # 3. Filter Nones for this specific workbook and extend the master list
            valid_wb_sheets = [s for s in results if s is not None]
            sheets_list.extend(valid_wb_sheets)

    return sheets_list
        

def get_actual_max_row(sheet: Worksheet) -> int:
    for row_num in range(sheet.max_row, 0, -1):
        if any(cell.value is not None for cell in sheet[row_num]):
            return row_num
    return 1


def random_3rows(slctd_sheets: list[str]) -> pd.DataFrame:
    """Extract 3 random rows from selected sheets into a clean Pandas DataFrame."""
    sp = st.session_state["selected_parameters"]
    rndm_3sheets = random.choices(slctd_sheets, k=3) if len(slctd_sheets) >= 3 else slctd_sheets
    
    rows_data = []
    print("sheets", slctd_sheets)
    reconverted_sheets: list[Worksheet] = 


    I need the workbook resp. the file path



    for sheet in reconverted_sheets:
        rnd_int: int = random.randint(sp["min_row"], get_actual_max_row(sheet))
        # openpyxl slice returns a tuple of tuples: ((cell1, cell2, ...),)
        cell_range: tuple = sheet[f"{sp['min_col']}{rnd_int}":f"{sp['max_col']}{rnd_int}"]
        row_values: list = [cell.value for cell in cell_range[0]]
        rows_data.append(row_values)

    # Generate default column names (Col A, Col B, etc. or index-based)
    cols_count = len(rows_data[0]) if rows_data else 0
    col_names: list[str] = [f"Col{i+1}" for i in range(cols_count)]
    
    return pd.DataFrame(rows_data, columns=col_names)       
    

"""
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