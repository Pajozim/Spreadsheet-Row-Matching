
import os
import openpyxl
import re
from pyxlsb import open_workbook as open_xlsb
from pathlib import Path
from config import SOURCE_FILE, OUTPUT_FILE, SF_DIR
from excel_automation.py_files.db_related import keyword_search
from excel_automation.py_files.process_catalogues import clean_unit


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


def process_sheets(wb: openpyxl.Workbook, wb_op: openpyxl.Workbook):
    for sheet in wb.worksheets:
        print(f"Processing sheet: {sheet.title}")

        # Ignores hidden or leftover sheets
        if sheet.sheet_state != 'visible': continue 

        # to focus the search under headers
        context = []

        for row in sheet.iter_rows(min_row=SF_DIR["MIN_ROW"], min_col=SF_DIR["MIN_COLUMN"], max_col=SF_DIR["MAX_COLUMN"]):

            # setting indexes
            helper1_idx = SF_DIR["HELPER_COLUMN_1"] - SF_DIR["MIN_COLUMN"]
            description_idx = SF_DIR["DESCRIPTION_COLUMN"] - SF_DIR["MIN_COLUMN"]
            unit_idx = SF_DIR["UNIT_COLUMN"] - SF_DIR["MIN_COLUMN"]

            # setting values
            pre_helper1 = row[helper1_idx].value
            pre_description = row[description_idx].value
            pre_unit = row[unit_idx].value

            # Safe extraction + stripping whitespace
            helper1 = str(pre_helper1).strip() if pre_helper1 is not None else "" # helper for conditions
            #helper2 = str(row[2].value).strip() if row[2].value is not None else "" # helper for conditions
            description = str(pre_description).strip() if pre_description is not None else "" # description
            unit = str(pre_unit).strip() if pre_unit is not None else "" # unit column

            # Set context
            #if helper2 == "No." or not helper2 and row[3].value:
            if helper1.isdigit():
                # specific context (partida/concepto)
                cleaned_concept = re.findall(r'\w+', description)
                trimmed_concept = " OR ".join([word.lower() + "*" for word in cleaned_concept if len(word) >= 3])

                # the sheet title, or more exact: disciplina
                # base_sheet = sheet.title.lower()
                # cleaned_sheet = re.findall(r'\w+', base_sheet)
                # trimmed_sheet = " OR ".join([word.lower() + "*" for word in cleaned_sheet if len(word) >= 3]) 

                # context packed
                context = [trimmed_concept]
                continue

            """
            # Clear skip condition:
            # 1. Column 0 is empty
            # 2. Column 0 equals "Plano"
            # 3. Column 3 is empty
            if not helper1 or helper1 == "Plano" or not description or row[0].row is None:
                # print("Skipping row:", row[0].row, "Val:", repr(row[0].value))
                continue
            """

            if not description: continue

            kws: list[str] = extract_keywords(description)
            #print("kws:", kws)

            matches: list[dict] = keyword_search(kws, clean_unit(unit), context)
            print(f"matches: {matches}")

            row_number: int = row[0].row # type: ignore
            if matches: insert_matches(matches, wb_op, sheet.title, SF_DIR["COLUMN_TO_INSERT"], row_number)


def process_source_file(file_path: Path | str = SOURCE_FILE, output_path: Path | str = OUTPUT_FILE) -> None:

    if not file_path or not os.path.exists(file_path):
        print("Error: No file path provided / file not found.")
        return

    print(f"Loading source file {file_path} from .xlsx_files/ ... \n")


    wb = openpyxl.load_workbook(file_path, data_only=True, rich_text=True)
    wb_op = openpyxl.load_workbook(output_path, data_only=True)

    process_sheets(wb, wb_op)

    wb_op.save(output_path)
    wb_op.close()
    wb.close()

    return