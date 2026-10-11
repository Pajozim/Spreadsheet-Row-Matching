import os
import openpyxl
import re
import config as cfg
import streamlit as st
import fnmatch
import unicodedata
import shutil

from typing import Literal
from pathlib import Path
from .stlt_helpers import col_to_num
from .db_related import insert_catalogue_rows, keyword_search

# regarding catalogueing files

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


def preparing_filters(item: dict) -> list:
    # pre-processing filters
    compiled_filters = []
    for f in item["filters"]:

        raw_cols: str = str(f.get("targ_col_raw", ""))
        col_add: str | None = None
        if raw_cols.startswith("any"):
            col_add = "any"
        elif raw_cols.startswith("all"):
            col_add = "all"

        extracted_raw_cols: list = re.findall(r'[A-Z]+', raw_cols)
        cols_parsed = [col_to_num(col, 0) for col in extracted_raw_cols]
        
        omit_row = str(f.get("omit_row_when", ""))
        if omit_row:               
            delimiter = ";" if ";" in omit_row else ","
            targets: list = [
                v.strip().lower() 
                for v in omit_row.split(delimiter)
                if v.strip()
                ]
        else:
            targets = [""]
            
        compiled_filters.append({
            "col_add": col_add,
            "columns": cols_parsed,
            "targets": targets
        })

    return compiled_filters


# regarding search file

def insert_matches(matches: list[dict], wb: openpyxl.Workbook, sheet: str, col: int, row: int) -> None:
    ws = wb[sheet]

    if not matches:
        ws.cell(row=row, column=col).value = "Found nothing" # type: ignore (the cells are very likely non-merged)
        return

    for match in matches:
        ws.cell(row=row, column=col).value = match["clave"]
        ws.cell(row=row, column=col + 1).value = match["description"]
        ws.cell(row=row, column=col + 2).value = match["unit"]
        ws.cell(row=row, column=col + 3).value = match["cell_reference"]
        col += 5

    return


def extract_keywords(cellvalue: str) -> list[str]:
    """
    Extracts text from a cell and splits it into keywords.
    """

    keywords = []

    if re.search(r'[.,]', cellvalue): 
        keywords = re.split(r'[.,]', cellvalue)[0]
    else:
        keywords = cellvalue

    return keywords.lower().split()


def search_for_keywords(payload: list[dict]) -> bool:
    
    success = False
    return success

# general/shared process

def process_spreadsheet_files(action: Literal["catalogue", "search"], parameters: dict, db_path: Path | str | None = cfg.DB_PATH) -> bool:

    if not parameters or db_path is None:
        print("Error: parameters and/or db_path is None")
        return False

    print(f"Loading catalogue data from {parameters} ... \n")

    # Initialize variables
    payload: list = []

    # pre-processing filters
    compiled_filters = preparing_filters(parameters)
    
    # Load catalogues workbooks into memory
    for path, item in parameters.items():
        if not os.path.exists(path):
            print(f"Error: Catalogue file '{path}' not found!")
            return False
            
        wb_cat = openpyxl.load_workbook(path, data_only=True)
        #print(f"Loaded: \n {cat_path} \n")

        for sheet in wb_cat.worksheets:

            # Ignores hidden or leftover sheets
            if sheet.sheet_state != 'visible': continue
            
            #print(f"\n Processing sheet: {sheet.title}")
            if sheet.title not in item["selected_sheets"]: continue            
            
            # Do NOT use values_only=True so we can access cell.coordinate
            for row in sheet.iter_rows(min_row=item["min_row"], max_row=item["max_row"], min_col=item["min_col"], max_col=item["max_col"]):

                vals: list = [
                    str(c.value).strip().lower() 
                    if c.value is not None else "" 
                    for c in row
                    ]

                # failsafe for missing columns
                if len(vals) < item["max_col"]:
                    for _ in range(len(vals), item["max_col"]):
                        vals.append("")

                # use of filter conditions
                skip_row = False
                for f in compiled_filters:
                    col_add: str | None = f["col_add"]
                    cols: list = f["columns"]
                    targets: list = f["targets"]

                    if cols:

                        if col_add == "all":
                            # Fix 3: Standardized fnmatch across all branches
                            if all(any(fnmatch.fnmatch(vals[c], t) for t in targets) for c in cols):
                                skip_row = True
                                break

                        elif col_add == "any":
                            if any(any(fnmatch.fnmatch(vals[c], t) for t in targets) for c in cols):
                                skip_row = True
                                break

                        else:
                            if any(fnmatch.fnmatch(vals[cols[0]], t) for t in targets):
                                skip_row = True
                                break

                    else:
                        # --- CASE B: Check ANY Column in Row ---                        
                        if any(
                            fnmatch.fnmatch(cell_str, t) 
                            for cell_str in vals 
                            for t in targets
                        ):
                            skip_row = True
                            break

                if skip_row:
                    continue  # SKIPS TO THE NEXT ROW IN SHEET.ITER_ROWS()

                if action == "search":

                    pass

                else:

                    """visualization of column mappings e.g.:
                        "column_mappings": [ 
                            {"CLAVE": ["A", "B", "C"], "delimiter": "-"},
                            {"Descripcion": [D]},
                            {"Unit": [E]},
                        ...
                        ], ...
                    """

                    cm: list = item.get("column_mappings", [])
                    row_data: dict = {}
                    for mItem in cm:
                        delimiter = mItem.get("delimiter")
                        
                        for k, v in mItem.items():
                            if k == "delimiter":
                                continue
                                
                            if delimiter:
                                iter_vals = [vals[col_idx] for col_idx in v]
                                row_data[k] = delimiter.join(iter_vals)
                            else:
                                col_idx = v[0]
                                row_data[k] = vals[col_idx]

                    # adding extra informations to the row_data
                    row_data.update({"sheets": sheet.title, "cell_ref": f"=[{path}]{sheet.title}!{row[0].coordinate}"})

                    # adding row_data to the payload
                    payload.append(row_data)

    if action == "catalogue":
        # Save data to database
        result: bool = insert_catalogue_rows(payload, st.session_state["db_template"])
    else:
        result: bool = search_for_keywords(payload)

    return result



"""
def process_source_file(src_parameters: dict) -> None:

    src_file = src_parameters.get("source_file", "")
    sf_split = src_file.split(".")
    cpd_file = sf_split[0] + "_POPULATED" + sf_split[1]
    file_path = Path(cfg.THIS_DIR / "src/excel_automation/xlsx_files" / src_file)
    output_path = Path(cfg.THIS_DIR / "src/excel_automation/xlsx_files" / cpd_file)
    
    if not file_path or not os.path.exists(file_path):
        print("Internal Error: file not found.", file_path)
        return

    print(f"Copying source file {file_path} ... \n")

    output_path = shutil.copy2(file_path, output_path)

    wb = openpyxl.load_workbook(file_path, data_only=True, rich_text=True)
    wb_op = openpyxl.load_workbook(output_path, data_only=True)

    process_sheets(wb, wb_op, src_parameters)

    wb_op.save(output_path)
    wb_op.close()
    wb.close()

    return

"""