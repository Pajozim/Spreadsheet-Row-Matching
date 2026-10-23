"""visualization of:
st.session_state["c-parameters"] = parameters = {
    "PATH1": {
        "selected_sheets": ["sheet1", "sheet2", ...],
        "min_row": 1,
        "max_row": None,
        "min_col": A,
        "max_col": None,
        "filters": [
            {
                "id": "1",
                "targ_col_raw": "B",
                "wordpiece": "Sidenote",
            },
            {
                "id": "2",
                "targ_col_raw": "G",
                "wordpiece": "no, usa*, cancel*",   
            },
            {
                "id": "3",
                "targ_col_raw": "D",
                "wordpiece": "",           
            },
            {
                "id": "4",
                "targ_col_raw": "",
                "wordpiece": "Sidenote",               
            },
            {
                "id": "5",
                "targ_col_raw": "B, C, D"
                "wordpiece": "",
            },
            ...
        ],
        "column_mappings": [
                {"DB_CAT1": [0, 1, 2], "delimiter": "-"},
                {"DB_CAT2": [3]},
                {"DB_CAT3": [4]},
                ...
            ],
    },
    "PATH2": {
        ...
    },
    ...
}
"""

"""visualization of:
st.session_state["db_template"] = db_template = [
    {
        "id": (uuid.uuid4()[:8]),
        "name": "CLAVE"
    },
    {
        "id": (uuid.uuid4()[:8]),
        "name": "Description"
    },
    ...
]
"""

"""visualization of:
st.session_state["db_addtnl"] = {
    "fts5": ["Description", ...], 
    "index": ["CLAVE", ...]
}
"""


"""
st.session_state["s-parameters"] = {
    "PATH1": {
        "selected_sheets": ["sheet1", "sheet2", ...],
        "min_row": 1,
        "max_row": None,
        "min_col": A,
        "max_col": A,
        "filters": [
            {
                "id": "1",
                "targ_col_raw": "B",
                "wordpiece": "Sidenote",
            },
        "result_col": "N",
        ...,
    },
}
"""

"""
payload = {
    "PATH1": [
        {
            "kws": ["LTFR-01", "Suministro", "fabricación", ...],
            "texts": "(LTFR-01) Suministro, fabricación y elaboración de LAMBRIN de panel de yeso con ... pieza ...",
            "wb":  <openpyxl.workbook.workbook.Workbook object at 0x7e7580ba61b0>, 
            "sheet": Luminarios, 
            "result_col": 14, 
            "row": 24
        },
        ...
    ],
    "PATH2": [
        ...
    ],
    ...
`}
"""