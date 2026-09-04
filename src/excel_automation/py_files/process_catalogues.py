import os
import openpyxl
import re
import config as cfg
import streamlit as st
from src.excel_automation.py_files.db_related import insert_catalogue_rows
from pathlib import Path

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


def extract_sheets() -> None:
    parameters = st.session_state["parameters"]

    for path in parameters:
        wb = openpyxl.load_workbook(path, read_only=True)
        parameters[path]["selected_sheets"] = wb.sheetnames
        wb.close()


def process_spdsht_files(parameters: dict, db_path: Path | str | None = cfg.DB_PATH) -> None:

    if not parameters or db_path is None:
        print("Error: parameters and/or db_path is None")
        return

    print(f"Loading catalogue data from {parameters} ... \n")

    # Initialize variables
    catalogues_data: list = []
    
    # Load catalogues workbooks into memory
    for cat_path in parameters:
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
