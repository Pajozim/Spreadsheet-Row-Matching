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

    # 1. Dynamically find all user-created B-Tree indexes
    cursor.execute("""
        SELECT name FROM sqlite_master 
        WHERE type = 'index' 
          AND name NOT LIKE 'sqlite_%';
    """)
    indexes = cursor.fetchall()

    # 2. Drop each index dynamically
    for (idx_name,) in indexes:
        cursor.execute(f'DROP INDEX IF EXISTS "{idx_name}";')

    # 3. Drop base and virtual tables
    cursor.execute("DROP TABLE IF EXISTS catalogue_items;")
    cursor.execute("DROP TABLE IF EXISTS catalogue_fts;")

    # 4. Commit drops BEFORE calling VACUUM
    conn.commit()

    # 5. VACUUM in autocommit mode to recover disk space
    conn.isolation_level = None
    cursor.execute("VACUUM;")
    conn.isolation_level = "DEFERRED"

    conn.close()
    print("Database cleaned successfully.")


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


def init_db(cat_list: list[dict], addtnl_list: dict | None = None, drop_tables: bool = True) -> None:
    if addtnl_list is None:
        addtnl_list = {}

    conn = get_db_connection()
    cursor = conn.cursor()

    if drop_tables:
        clean_db()

    # 1. Base Table
    joined_catgrs = " TEXT, ".join(cat["name"] for cat in cat_list)
    query = f"""
        CREATE TABLE IF NOT EXISTS catalogue_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            {joined_catgrs},
            sheets TEXT,
            cell_ref TEXT
        );"""
    cursor.execute(query)

    # 2. B-Tree Indexes (Dynamic Index Names)
    index_cols = addtnl_list.get("index", [])
    for item in index_cols:
        clean_item = item.strip("[]`\"'")
        # Generate a unique index name per column
        query = f"CREATE INDEX IF NOT EXISTS idx_{clean_item} ON catalogue_items({item});"
        cursor.execute(query)

    # 3. FTS5 Table & Triggers
    fts5_cols = addtnl_list.get("fts5", [])
    if fts5_cols:
        joined_fts5cats = ", ".join(fts5_cols)
        query = f"""
            CREATE VIRTUAL TABLE IF NOT EXISTS catalogue_fts USING fts5(
                {joined_fts5cats},
                content='catalogue_items',
                content_rowid='id',
                tokenize='unicode61 remove_diacritics 1'
            );
        """
        cursor.execute(query)

        # Dynamic Column Formatting for Triggers
        new_cols = ", ".join([f"new.{col}" for col in fts5_cols])
        old_cols = ", ".join([f"old.{col}" for col in fts5_cols])

        # Separate statements executed cleanly via executescript
        triggers_script = f"""
            CREATE TRIGGER IF NOT EXISTS catalogue_ai AFTER INSERT ON catalogue_items BEGIN
                INSERT INTO catalogue_fts(rowid, {joined_fts5cats}) VALUES (new.id, {new_cols});
            END;

            CREATE TRIGGER IF NOT EXISTS catalogue_ad AFTER DELETE ON catalogue_items BEGIN
                INSERT INTO catalogue_fts(catalogue_fts, rowid, {joined_fts5cats}) VALUES('delete', old.id, {old_cols});
            END;

            CREATE TRIGGER IF NOT EXISTS catalogue_au AFTER UPDATE ON catalogue_items BEGIN
                INSERT INTO catalogue_fts(catalogue_fts, rowid, {joined_fts5cats}) VALUES('delete', old.id, {old_cols});
                INSERT INTO catalogue_fts(rowid, {joined_fts5cats}) VALUES (new.id, {new_cols});
            END;
        """
        cursor.executescript(triggers_script)

    # 4. Commit Schema Changes
    conn.commit()

    # 5. VACUUM in Autocommit Mode
    conn.isolation_level = None
    cursor.execute("VACUUM;")
    conn.isolation_level = "DEFERRED"

    conn.close()
    print("Database initiated successfully.")


def get_db_columns(table_name: str = "catalogue_items") -> list[str]:
    """Return column names and data types for a given table."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(f"PRAGMA table_info({table_name})")
    table_info = cursor.fetchall()
    conn.close()
    return [ti[1] for ti in table_info] if table_info else []


def insert_catalogue_rows(rows: list[dict], columns: list[dict]) -> bool:
    """Bulk insert parsed catalogue rows into SQLite."""
    conn = get_db_connection()
    cursor = conn.cursor()

    col_names = [col["name"] for col in columns]

    # 1. Join column names (e.g., "name, price, stock")
    cols_sql = ", ".join(col_names)

    # 2. Generate matching corresponding placeholders (e.g., ":name, :price, :stock")
    values_sql = ", ".join(f":{col_name}" for col_name in col_names)

    # 3. Construct the final query
    query = f"""
        INSERT INTO catalogue_items ({cols_sql}, sheets, cell_ref)
        VALUES ({values_sql}, :sheets, :cell_ref)
    """

    # 4. Execute using the dictionary keys matching the placeholders
    cursor.executemany(query, rows)

    conn.commit()
    conn.close()
    
    print(f"Successfully inserted {len(rows)} items into database.")

    return True


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


def show_db(fts: bool = False) -> list:
    conn = sqlite3.connect(cfg.DB_PATH)
    cursor = conn.cursor()
    if fts:
        cursor.execute("SELECT * FROM catalogue_fts")
    else:
        cursor.execute("SELECT * FROM catalogue_items LIMIT 20")
    rows = cursor.fetchall()
    #for row in rows: print(row)
    conn.close()
    return rows

    