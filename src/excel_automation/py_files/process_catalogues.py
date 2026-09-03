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


"""visualization of workbook filter collector
st.session_state["fData"] = fData = {
    "PATH1": {
        "selected_sheets": ["sheet1", "sheet2", ...],
        "min_row": 1,
        "max_row": None,
        "min_col": 1,
        "max_col": None,
        "filters": [
            {
                "id": "1",
                "targ_col_raw": "B",
                "omit_row_when": "Sidenote",
                "parsed_tCol": 2
            },
            {...},
            ...
        ]
    },
    "PATH2": {
        ...
    },
    ...
}
"""


def extract_sheets() -> None:
    fData = st.session_state["fData"]

    for path in fData:
        wb = openpyxl.load_workbook(path, read_only=True)
        fData[path]["selected_sheets"] = wb.sheetnames
        wb.close()

