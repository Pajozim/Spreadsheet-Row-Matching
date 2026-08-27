import os
import openpyxl
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
import re
from pathlib import Path
import config as cfg
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