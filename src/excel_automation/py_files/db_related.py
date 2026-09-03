import sqlite3
import re
import config as cfg
from pathlib import Path

def get_db_connection() -> sqlite3.Connection:
    """Connect to SQLite database and return connection."""
    conn = sqlite3.connect(cfg.DB_PATH)
    # Return rows as dictionary-like objects for easier column access
    conn.row_factory = sqlite3.Row
    return conn


def clean_db() -> None:
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("DROP TABLE IF EXISTS catalogue_items;")
    cursor.execute("DROP TABLE IF EXISTS catalogue_fts;")
    cursor.execute("DROP INDEX IF EXISTS idx_clave;")
    cursor.execute("VACUUM;")
    conn.commit()
    conn.close()
    return


def create_optimized_indexes(table_name: str, sample_rows: list[dict]):
    """
    Automatically creates standard B-Tree indexes for short/filter columns 
    and setups FTS5 for long text columns.
    """
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # Calculate average character length per column from sample data
    col_lengths = {}
    for row in sample_rows:
        for col, val in row.items():
            if col == "cell_reference":
                continue
            val_str = str(val or "")
            col_lengths[col] = col_lengths.get(col, 0) + len(val_str)
            
    num_samples = len(sample_rows) or 1
    avg_lengths = {col: total / num_samples for col, total in col_lengths.items()}

    # Automatically index short columns (< 20 avg chars) as B-Tree filters
    short_cols = [col for col, avg_len in avg_lengths.items() if avg_len < 20]
    for col in short_cols:
        cursor.execute(f"CREATE INDEX IF NOT EXISTS idx_{col} ON {table_name}({col});")

    conn.commit()


def init_db(ctlg_data: list[dict], drop_tables: bool = False) -> None:
    conn = get_db_connection()
    cursor = conn.cursor()

    if drop_tables:
        clean_db()

    # STEP 1: Analyze sample rows in memory (using up to 100 extracted rows)
    sample_size = min(len(ctlg_data), 100)
    samples = ctlg_data[:sample_size]

    col_avg_lengths = {}
    for col in ctlg_data[0].keys():
        total_len = sum(len(str(row.get(col, ""))) for row in samples)
        col_avg_lengths[col] = total_len / max(sample_size, 1)

    # Classify columns dynamically
    short_cols = [col for col, avg in col_avg_lengths.items() if avg < 30]
    long_cols = [col for col, avg in col_avg_lengths.items() if avg >= 30]

    # STEP 2: Create base SQLite table
    cols_sql = ", ".join([f"{col} TEXT" for idx, col in enumerate(short_cols + long_cols)])
    cursor.execute(f"CREATE TABLE IF NOT EXISTS catalogue_items (id INTEGER PRIMARY KEY AUTOINCREMENT, {cols_sql});")

    # STEP 3: Bulk insert extracted rows
    #col_names = ", ".join(columns)
    #placeholders = ", ".join([f":{col}" for col in columns])
    
    cursor.execute("BEGIN TRANSACTION;")
    #cursor.executemany(f"INSERT INTO catalogue_items ({col_names}) VALUES ({placeholders})", extracted_rows)
    conn.commit()

    # STEP 4: Build standard indexes on short filter columns
    for col in short_cols:
        cursor.execute(f"CREATE INDEX IF NOT EXISTS idx_{col} ON catalogue_items({col});")

    # STEP 5: Build FTS5 index on long text columns (descriptions)
    if long_cols:
        fts_cols = ", ".join(long_cols)
        cursor.execute(f"""
            CREATE VIRTUAL TABLE IF NOT EXISTS catalogue_fts USING fts5(
                {fts_cols},
                content='catalogue_items',
                content_rowid='id',
                tokenize='unicode61 remove_diacritics 1'
            );
        """)
        cursor.execute(f"""
            INSERT INTO catalogue_fts(rowid, {fts_cols})
            SELECT id, {fts_cols} FROM catalogue_items;
        """)

    conn.commit()
    print("Database built, inserted, and indexed successfully at: ", cfg.DB_PATH)


def insert_catalogue_rows(rows: list[dict], columns: list[str]):
    """Bulk insert parsed catalogue rows into SQLite."""
    conn = get_db_connection()
    cursor = conn.cursor()

    # 1. Join column names (e.g., "name, price, stock")
    cols_sql = ", ".join(columns)

    # 2. Generate matching placeholders (e.g., ":name, :price, :stock")
    placeholders_sql = ", ".join([f":{col}" for col in columns])

    # 3. Construct the final query
    query = f"""
        INSERT INTO catalogue_items ({cols_sql}, cell_reference)
        VALUES ({placeholders_sql}, :cell_reference)
    """

    # 4. Execute using the dictionary keys matching the placeholders
    cursor.executemany(query, rows)

    conn.commit()
    conn.close()
    print(f"Successfully inserted {len(rows)} items into database.")


def parse_clave_to_sheet(clave: str) -> str:
    """
    Extracts first 3 digits from clave and formats as sheet name.
    Example: '09-1' -> '091' -> strip leading zeros -> '91'
    """
    if not clave:
        return ""
    
    # Extract all digits
    digits = re.sub(r"\D", "", str(clave))  # "09-1" becomes "091"
    
    first_three_digits = digits[:3]
    
    # Strip leading zero if your sheets are named '91' instead of '091'
    sheet_name = first_three_digits.lstrip("0")
    
    return sheet_name if sheet_name else first_three_digits


def keyword_search(keywords: list[str], unit: str, context: list[str], max_results: int) -> list[dict]:
    # context parameter should be lowercase (should've been done in caller function)

    # Connect to database
    conn = sqlite3.connect(cfg.DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    # Clean keywords to prevent SQL FTS syntax errors
    clean_kw = [kw.strip().replace('"', '') for kw in keywords if kw.strip()]
    if not clean_kw:
        return []
    # Construct a FTS query: e.g., "CONDICIONES"* OR "GENERALES"*
    joint_kws = ' OR '.join(f'\"{kw}\"*' for kw in clean_kw)
    fts_query = f'col1:({joint_kws})'

    results = []
    current_search_matches = None

    # =========================================================================
    # STAGE 1: Top 2 results from INSUMOS (Sheet: 'Subcontratos')
    # =========================================================================

    # 1. Try to find the Context Header in Subcontratos (clave ends in '000')
    f3_digits = None

    for scope in context:
        # context parameters, visuals on scope for instance can be "preliminares*" or "mobiliario* OR equipo* OR cocina*"
        context_fts_query = f'col1:{scope}'

        # Try to find the Context Header in Subcontratos
        cursor.execute("""
            SELECT c.clave
            FROM catalogue_items c
            JOIN catalogue_fts fts ON c.id = fts.rowid
            WHERE c.sheet = 'subcontratos'
            AND c.clave LIKE '%000'
            AND catalogue_fts MATCH ?
            ORDER BY rank
            LIMIT 1
        """, (context_fts_query,))
        
        header_match = cursor.fetchone()
    
        if header_match and header_match["clave"]:
            # Extract clave and take the first 3 digits (e.g., "4-03-1-000" -> "4-03")
            f3_digits = str(header_match["clave"])[:4] # str() is defensive here; usually it shall be always a string
            break

    # 2. Query Subcontratos for top 2 keyword matches (Scoped vs. Free Search)
    if f3_digits:
        # Scoped Search: Restrict to rows matching the discipline/context prefix
        cursor.execute("""
            SELECT c.* 
            FROM catalogue_items c
            JOIN catalogue_fts fts ON c.id = fts.rowid
            WHERE c.sheet = 'subcontratos'
            AND c.unit = ?
            AND c.clave LIKE ?
            AND catalogue_fts MATCH ?
            ORDER BY rank
            LIMIT ?
        """, (unit, f"{f3_digits}%", fts_query, max_results,))

    else:
        # Fallback Free Search: Query across all of Subcontratos
        cursor.execute("""
            SELECT c.* 
            FROM catalogue_items c
            JOIN catalogue_fts fts ON c.id = fts.rowid
            WHERE c.sheet = 'subcontratos'
            AND c.unit = ?
            AND catalogue_fts MATCH ?
            ORDER BY rank
            LIMIT ?
        """, (unit, fts_query, max_results,))
    
    insumos_matches: list[dict] = [dict(row) for row in cursor.fetchall()]
    results.extend(insumos_matches)

    # =========================================================================
    # STAGE 2: Top 2 results from MATRICES (via Catalogo General)
    # =========================================================================
    
    general_match = None
    fallback_search = False

    for scope in context:
        # context parameters, visuals on scope for instance can be "preliminares*" or "mobiliario* OR equipo* OR cocina*"
        context_fts_query = f'col1:{scope}'
        
        cursor.execute("""
            SELECT c.clave, c.col1, c.*
            FROM catalogue_items c
            JOIN catalogue_fts fts ON c.id = fts.rowid
            WHERE c.sheet = 'catalogo general'
            AND catalogue_fts MATCH ?
            ORDER BY rank
            LIMIT 1
        """, (context_fts_query,))
        
        general_match = cursor.fetchone()

        # sometimes the first title/context is too specific to be found in 'Catalogo General'
        # but unfortunately, so can be the second title/context
        if general_match:
            break

    if general_match:
        # 1. Extract clave (e.g., "09-1" -> "091" -> "91")
        raw_clave = general_match["clave"]  # e.g., "09-1"
        target_sheet = parse_clave_to_sheet(raw_clave)  # Output: "91"

        # 2. Query target sheet for the best match by ranking quantity of matches
        cursor.execute("""
            SELECT c.* 
            FROM catalogue_items c
            JOIN catalogue_fts fts ON c.id = fts.rowid
            WHERE c.sheet = ?
              AND c.unit = ?
              AND catalogue_fts MATCH ?
            ORDER BY rank
            LIMIT ?
        """, (target_sheet, unit, fts_query, max_results,))

        current_search_matches = cursor.fetchall()
        if not current_search_matches:
            fallback_search = True

    if fallback_search or not general_match:

        # Fallback to all alround search
        cursor.execute("""
            SELECT c.*
            FROM catalogue_items c
            JOIN catalogue_fts fts ON c.id = fts.rowid
            WHERE c.sheet != 'subcontratos'
                AND c.unit = ?
                AND catalogue_fts MATCH ?
            ORDER BY rank
            LIMIT ?
        """, (unit, fts_query, max_results,))
        
    cursor_matches = current_search_matches or cursor.fetchall()
    target_activity_match: list = [dict(row) for row in cursor_matches]
    if target_activity_match:
        results.extend(target_activity_match)

    conn.close()

    print("results: ", results)

    return results


def show_db(fts: bool = False):
    conn = sqlite3.connect(cfg.DB_PATH)
    cursor = conn.cursor()
    if fts:
        cursor.execute("SELECT * FROM catalogue_fts")
    else:
        cursor.execute("SELECT * FROM catalogue_items WHERE sheet IS NOT 'subcontratos'")
    rows = cursor.fetchall()
    for row in rows:
        print(row)
    conn.close()

    