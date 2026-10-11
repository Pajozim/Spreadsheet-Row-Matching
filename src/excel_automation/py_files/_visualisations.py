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
                {"DB_CAT1": ["A", "B", "D"], "delimiter": "+"},
                {"DB_CAT2": ["E"]},
                {"DB_CAT3": ["G"]},
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
        ...,
    },
}
"""