import sqlite3
import re
import config as cfg

from pathlib import Path
from typing import Literal

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


def get_db_columns(table_name: Literal["catalogue_items", "catalogue_fts"]) -> list[str]:
    """Return column names and data types for a given table."""
    conn = get_db_connection()
    cursor = conn.cursor()
    column_names = []
    match table_name:
        case "catalogue_items":
            cursor.execute(f"PRAGMA table_info({table_name})")
            table_info = cursor.fetchall()
            column_names = [ti[1] for ti in table_info] if table_info else []
        case "catalogue_fts":
            cursor.execute(f"SELECT * FROM {table_name} LIMIT 0")
            column_names = [col[0] for col in cursor.description]
    conn.close()
    return column_names


def insert_catalogue_rows(rows: list[dict], columns: list[dict]) -> tuple[bool, str]:
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
        INSERT INTO catalogue_items ({cols_sql}, cell_ref)
        VALUES ({values_sql}, :cell_ref)
    """

    # 4. Execute using the dictionary keys matching the placeholders
    cursor.executemany(query, rows)

    conn.commit()
    conn.close()

    success = f"Successfully inserted {len(rows)} items into database."
    print(success)

    return True, success


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


def keyword_search(payload: list[str], max_results: int = 3) -> list[dict]:
    # Connect to database
    conn = sqlite3.connect(cfg.DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    # Clean keywords to prevent SQL FTS syntax errors
    clean_kw = [kw.strip().replace('"', '') for kw in payload if kw.strip()]
    if not clean_kw:
        return []
    # Construct a FTS query: e.g., "CONDICIONES"* OR "GENERALES"*
    joint_kws = ' OR '.join(f'\"{kw}\"*' for kw in clean_kw)

    results = []
 
    cursor.execute("""
        SELECT c.*
        FROM catalogue_items c
        JOIN catalogue_fts fts ON c.id = fts.rowid
        WHERE catalogue_fts MATCH ?
        ORDER BY rank
        LIMIT ?
    """, (joint_kws, max_results,))
        
    cursor_matches = cursor.fetchall()
    target_activity_match: list = [dict(row) for row in cursor_matches]
    if target_activity_match:
        results.extend(target_activity_match)

    conn.close()

    return results


def show_db(conditional: Literal["all", "fts", "random"]) -> list[tuple]:
    conn = sqlite3.connect(cfg.DB_PATH)
    cursor = conn.cursor()
    match conditional:
        case "fts":
            cursor.execute("SELECT * FROM catalogue_fts")
        case "all":
            cursor.execute("SELECT * FROM catalogue_items")
        case "random":
            cursor.execute("SELECT * FROM catalogue_items ORDER BY RANDOM() LIMIT 20")
    rows = cursor.fetchall()
    #for row in rows: print(row)
    conn.close()
    return rows

    