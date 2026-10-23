import torch
import pandas as pd
import streamlit as st
from typing import cast
import torch

from torch import Tensor
from pandas import DataFrame
from sentence_transformers import SentenceTransformer
from .db_related import show_db, get_db_columns

@st.cache_resource
def loading_model() -> SentenceTransformer:
    # 1. Choosing the device
    device = "cuda" if torch.cuda.is_available() else "cpu"
    print(f"Running engine on device: {device.upper()}")

    # 2. Load multilingual-e5-small
    return SentenceTransformer('intfloat/multilingual-e5-small', device=device)

@st.cache_resource
def prep_batch() -> DataFrame:

    db_items: list[tuple] = show_db("all")
    col_names = get_db_columns("catalogue_items")

    if not db_items:
        raise ValueError("Error: No database items found.")
    
    if len(col_names) != len(db_items[0]):
        raise ValueError("Error: Number of columns does not match number of database items.")
        
    return pd.DataFrame(db_items, columns=col_names)

@st.cache_resource
def embed_catalogues() -> Tensor:
    
    model = loading_model()    
    database_items = prep_batch()
    
    # E5 format rule: Database entries must be prefixed with "passage: "
    fts_cols = get_db_columns("catalogue_fts")
    if not fts_cols:
        raise ValueError("Error: No FTS columns found.")

    # Use Option A: Combine row-by-row so each catalog item is a single coherent passage
    db_texts = [
        f"passage: {' '.join(str(row[col]) for col in fts_cols if pd.notna(row[col]))}"
        for _, row in database_items.iterrows()
    ]
    
    embeddings = model.encode(db_texts, convert_to_tensor=True, normalize_embeddings=True)
    
    return cast(torch.Tensor, embeddings)

 
def semantic_search(text: str) -> list[dict]:
    
    # call embedded catalogues
    model = loading_model()
    database_items = prep_batch()
    db_embeddings = embed_catalogues()

    # E5 format rule: Search strings must be prefixed with "query: "
    query_embedding: Tensor = model.encode(f"query: {text}", convert_to_tensor=True, normalize_embeddings=True) # type: ignore

    # 5. Compute cosine similarities to find the best match
    similarities: Tensor = model.similarity(query_embedding, db_embeddings)[0]
    
    best_match_idx = int(similarities.argmax().item())
    
    matched_item: list = database_items.iloc[best_match_idx].tolist()
    matched_item.pop(0)
    
    return matched_item

