from __future__ import annotations

from datetime import datetime, date, timedelta
from pathlib import Path
from typing import Dict, Iterable, Mapping, Optional, Union
import re

import duckdb
import pandas as pd
import pyarrow as pa
import pyarrow.parquet as pq

CACHE_ROOT = ".pos_cache"

TABLE_DATE_COLUMN = {
    "pos_sales_dtls": "c_date",
    "pos_sales_payment_dtls": "c_date",
    "pos_cancel_sales_item_dtls": "c_date",
    "item_master": "m_date",
}


_CORRUPTED_PARQUET_PATTERN = re.compile(r"'([^']+\\.parquet)'")


def quarantine_parquet_from_exception(exc: Exception) -> Optional[Path]:
    """Rename a corrupted parquet file referenced in ``exc`` so DuckDB will ignore it."""

    message = str(exc)
    match = _CORRUPTED_PARQUET_PATTERN.search(message)
    if not match:
        return None

    raw_path = match.group(1)
    file_path = Path(raw_path)

    if not file_path.exists():
        # Try interpreting as relative to workspace root.
        alt_path = Path.cwd() / raw_path
        if alt_path.exists():
            file_path = alt_path
        else:
            return None

    try:
        counter = 0
        while True:
            suffix = ".corrupt" if counter == 0 else f".corrupt{counter}"
            target = file_path.with_suffix(file_path.suffix + suffix)
            if not target.exists():
                file_path.replace(target)
                return target
            counter += 1
    except OSError:
        return None

DEFAULT_PARQUET_SCHEMAS = {
    "pos_sales_dtls": {
        "store_name": "string",
        "sales_no": "string",
        "item_name": "string",
        "qty": "float64",
        "disc_amt": "float64",
        "disc_name": "string",
        "pro_disc_amt": "float64",
        "sub_total": "float64",
        "svc_amt": "float64",
        "tax_amt": "float64",
        "item_sub_total": "float64",
        "c_date": "datetime64[ns]",
        "order_datetime": "datetime64[ns]",
        "take_away_item": "string",
    },
    "pos_sales_payment_dtls": {
        "store_name": "string",
        "sales_no": "string",
        "payment_name": "string",
        "tender_amt": "float64",
        "c_date": "datetime64[ns]",
    },
    "pos_cancel_sales_item_dtls": {
        "store_name": "string",
        "sales_no": "string",
        "item_name": "string",
        "qty": "float64",
        "c_date": "datetime64[ns]",
    },
    "item_master": {
        "item_name": "string",
        "category_code": "string",
        "m_date": "datetime64[ns]",
    },
}


def _db_root(db_name: str) -> Path:
    return Path(CACHE_ROOT) / db_name


def _duckdb_path(db_name: str) -> Path:
    return _db_root(db_name) / "pos.duckdb"


def _pq_dir(db_name: str, table: str) -> Path:
    return _db_root(db_name) / "parquet" / table


def duckdb_file_path(db_name: str) -> Path:
    """Return the path to the DuckDB database file, ensuring layout exists."""
    ensure_layout(db_name)
    return _duckdb_path(db_name)


def ensure_layout(db_name: str) -> None:
    """Ensure cache directory layout and an empty DuckDB file exist."""
    root = _db_root(db_name)
    duckdb_path = _duckdb_path(db_name)
    root.mkdir(parents=True, exist_ok=True)
    (root / "parquet").mkdir(exist_ok=True)
    for table in TABLE_DATE_COLUMN:
        _pq_dir(db_name, table).mkdir(parents=True, exist_ok=True)
        _ensure_stub_parquet(db_name, table)
    if not duckdb_path.exists():
        duckdb.connect(str(duckdb_path)).close()


def _ensure_stub_parquet(db_name: str, table: str) -> None:
    parquet_dir = _pq_dir(db_name, table)
    if any(parquet_dir.glob("*.parquet")):
        return

    schema = DEFAULT_PARQUET_SCHEMAS.get(table)
    if not schema:
        return

    data = {
        column: pd.Series(dtype=dtype)
        for column, dtype in schema.items()
    }
    df = pd.DataFrame(data)
    table_arrow = pa.Table.from_pandas(df, preserve_index=False)
    stub_path = parquet_dir / "__empty__.parquet"
    pq.write_table(table_arrow, stub_path)


def _glob_parquet(pattern_path: Path) -> bool:
    return any(pattern_path.parent.glob(pattern_path.name))


def _as_datetime(value: object) -> Optional[datetime]:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.replace(tzinfo=None)
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    try:
        return datetime.fromisoformat(str(value))
    except ValueError:
        return None


def max_local_date(db_name: str, table: str) -> Optional[datetime]:
    if table not in TABLE_DATE_COLUMN:
        raise ValueError(f"Unknown table: {table}")
    date_col = TABLE_DATE_COLUMN[table]
    parquet_dir = _pq_dir(db_name, table)
    pattern = parquet_dir / "*.parquet"
    if not _glob_parquet(pattern):
        return None
    con = duckdb.connect(database=":memory:")
    try:
        sql = (
            f"SELECT MAX({date_col}) AS max_dt "
            f"FROM parquet_scan('{pattern.as_posix()}', union_by_name = TRUE)"
        )
        row = con.execute(sql).fetchone()
        return _as_datetime(row[0]) if row else None
    finally:
        con.close()


def _resolve_sql_driver(server_cfg: Dict[str, str]) -> str:
    import pyodbc  # local import to avoid dependency when offline

    preferred_raw = server_cfg.get("driver")
    additional_raw = server_cfg.get("driver_candidates", [])
    additional = additional_raw if isinstance(additional_raw, list) else [additional_raw]

    try:
        available_drivers = [drv.strip() for drv in pyodbc.drivers()]
    except Exception:
        available_drivers = []

    available_lookup = {drv.lower(): drv for drv in available_drivers}

    if preferred_raw:
        preferred = preferred_raw.strip()
        matched = available_lookup.get(preferred.lower())
        if matched:
            return matched
        # Allow explicitly requested driver even if pyodbc does not report it.
        return preferred

    candidate_order: list[str] = []
    candidate_order.extend(drv for drv in additional if drv)
    candidate_order.extend(
        [
            "ODBC Driver 18 for SQL Server",
            "ODBC Driver 17 for SQL Server",
            "ODBC Driver 13 for SQL Server",
            "ODBC Driver 11 for SQL Server",
            "SQL Server Native Client 11.0",
            "SQL Server Native Client 10.0",
            "SQL Server",
        ]
    )

    for candidate in candidate_order:
        match = available_lookup.get(candidate.lower())
        if match:
            return match

    if available_drivers:
        return available_drivers[-1]

    raise RuntimeError(
        "No suitable SQL Server ODBC driver found. Please install 'ODBC Driver 17 for SQL Server' or a compatible driver."
    )


def _connect_sql_server(server_cfg: Dict[str, str]):
    import pyodbc  # local import to avoid dependency when offline

    driver = _resolve_sql_driver(server_cfg)
    encrypt = server_cfg.get("encrypt", "yes")
    trust = server_cfg.get(
        "trust_server_certificate",
        "yes" if str(encrypt).lower() in {"yes", "true", "1"} else "no",
    )
    timeout = str(server_cfg.get("timeout", 60))
    application_name = server_cfg.get("application_name", "POSAnalysisTool")
    mars = server_cfg.get("mars", "yes")
    extra = server_cfg.get("extra_params", "")
    if extra and not extra.endswith(";"):
        extra = f"{extra};"

    conn_str = (
        f"DRIVER={{{driver}}};"
        f"SERVER={server_cfg['server']};"
        f"DATABASE={server_cfg['database']};"
        f"UID={server_cfg['user']};"
        f"PWD={server_cfg['password']};"
        f"Encrypt={encrypt};"
        f"TrustServerCertificate={trust};"
        f"Connection Timeout={timeout};"
        f"MARS_Connection={mars};"
        f"Application Name={application_name};"
        f"{extra}"
    )
    return pyodbc.connect(conn_str)


def _clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.copy()
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            continue
        if df[col].dtype == "object":
            df[col] = df[col].map(lambda x: x.strip() if isinstance(x, str) else x)
    return df


def _load_existing_partition(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    try:
        return pq.read_table(path).to_pandas()
    except Exception:
        return pd.DataFrame()


def _write_partitioned_parquet(db_name: str, table: str, df: pd.DataFrame) -> int:
    if df.empty:
        return 0
    date_col = TABLE_DATE_COLUMN[table]
    parquet_dir = _pq_dir(db_name, table)
    parquet_dir.mkdir(parents=True, exist_ok=True)

    if date_col not in df.columns:
        raise ValueError(f"Table {table} expected column {date_col}")

    df = df.copy()
    df[date_col] = pd.to_datetime(df[date_col], errors="coerce")

    if table != "item_master":
        df["_partition_date"] = df[date_col].dt.date
    else:
        df["_partition_date"] = df[date_col].dt.date

    total_rows = 0

    for partition_value, chunk in df.groupby("_partition_date", dropna=False):
        if pd.isna(partition_value):
            partition_filename = "unknown.parquet"
        else:
            partition_filename = f"{partition_value.isoformat()}.parquet"
        target_path = parquet_dir / partition_filename

        existing = _load_existing_partition(target_path)
        combined = (
            pd.concat([existing, chunk.drop(columns=["_partition_date"])], ignore_index=True, copy=False)
            if not existing.empty
            else chunk.drop(columns=["_partition_date"])
        )
        combined = combined.drop_duplicates()
        table_arrow = pa.Table.from_pandas(combined, preserve_index=False)
        pq.write_table(table_arrow, target_path)
        total_rows += len(chunk.index)

    return total_rows


def sync_from_sqlserver(
    db_name: str,
    server_cfg: Dict[str, str],
    since: Optional[Union[datetime, Mapping[str, Optional[datetime]]]] = None,
    until: Optional[Union[datetime, Mapping[str, Optional[datetime]]]] = None,
) -> Dict[str, int]:
    ensure_layout(db_name)
    stats = {table: 0 for table in TABLE_DATE_COLUMN}
    sql_conn = _connect_sql_server(server_cfg)
    try:
        if isinstance(since, Mapping):
            since_map: Mapping[str, Optional[datetime]] = since
        else:
            since_map = {table: since for table in TABLE_DATE_COLUMN}

        if isinstance(until, Mapping):
            until_map: Mapping[str, Optional[datetime]] = until
        else:
            until_map = {table: until for table in TABLE_DATE_COLUMN}

        for table, date_col in TABLE_DATE_COLUMN.items():
            table_since = since_map.get(table) if since_map else None
            table_until = until_map.get(table) if until_map else None

            query = f"SELECT * FROM {table}"
            where_clauses: list[str] = []
            params: list[object] = []

            if table_since is not None:
                start_boundary = table_since - timedelta(days=1) if table_until is None else table_since
                where_clauses.append(f"{date_col} >= ?")
                params.append(start_boundary.strftime("%Y-%m-%d %H:%M:%S"))

            if table_until is not None:
                where_clauses.append(f"{date_col} <= ?")
                params.append(table_until.strftime("%Y-%m-%d %H:%M:%S"))

            if where_clauses:
                query += " WHERE " + " AND ".join(where_clauses)

            query += f" ORDER BY {date_col}"

            cursor = sql_conn.cursor()
            try:
                if params:
                    cursor.execute(query, params)
                else:
                    cursor.execute(query)
                columns = [col[0] for col in cursor.description or []]
                while True:
                    rows = cursor.fetchmany(50000)
                    if not rows:
                        break
                    chunk = pd.DataFrame.from_records(rows, columns=columns)
                    if chunk.empty:
                        continue
                    cleaned = _clean_dataframe(chunk)
                    stats[table] += _write_partitioned_parquet(db_name, table, cleaned)
            finally:
                cursor.close()
    finally:
        sql_conn.close()
    return stats


def init_and_sync(db_name: str, server_cfg: Dict[str, str]) -> Dict[str, int]:
    ensure_layout(db_name)
    since_map = {table: max_local_date(db_name, table) for table in TABLE_DATE_COLUMN}
    if all(value is None for value in since_map.values()):
        since_arg: Optional[Union[datetime, Mapping[str, Optional[datetime]]]] = None
    else:
        since_arg = since_map
    return sync_from_sqlserver(db_name, server_cfg, since_arg)


def open_conn(db_name: str) -> duckdb.DuckDBPyConnection:
    ensure_layout(db_name)
    attempts = 0
    while True:
        conn = duckdb.connect(str(_duckdb_path(db_name)))
        try:
            register_parquet_views(conn, db_name)
            create_clean_views(conn, db_name)
            return conn
        except duckdb.Error as err:
            conn.close()
            quarantined = quarantine_parquet_from_exception(err)
            if quarantined is None or attempts >= 2:
                raise
            attempts += 1


def register_parquet_views(con: duckdb.DuckDBPyConnection, db_name: str) -> None:
    for table in TABLE_DATE_COLUMN:
        parquet_path = _pq_dir(db_name, table) / "*.parquet"
        relation = (
            f"parquet_scan('{parquet_path.as_posix().replace('\\', '/')}', union_by_name = TRUE)"
        )
        con.execute(
            f"CREATE OR REPLACE VIEW {table} AS "
            f"SELECT *, '{db_name}' AS source_database FROM {relation}"
        )
    con.execute(
        """
        CREATE OR REPLACE VIEW v_raw_sales AS
        SELECT *, DATE(c_date) AS c_date_date FROM pos_sales_dtls
        """
    )
    con.execute(
        """
        CREATE OR REPLACE VIEW v_raw_payments AS
        SELECT *, DATE(c_date) AS c_date_date, TRIM(UPPER(payment_name)) AS payment_name_norm
        FROM pos_sales_payment_dtls
        """
    )
    con.execute(
        """
        CREATE OR REPLACE VIEW v_raw_cancel AS
        SELECT DISTINCT store_name, sales_no, DATE(c_date) AS c_date_date
        FROM pos_cancel_sales_item_dtls
        """
    )


def create_clean_views(con: duckdb.DuckDBPyConnection, current_database: str) -> None:
    con.execute(
        """
        CREATE OR REPLACE VIEW v_sales_clean AS
        SELECT s.*
        FROM v_raw_sales AS s
        WHERE NOT EXISTS (
            SELECT 1
            FROM v_raw_cancel AS c
            WHERE s.store_name = c.store_name
              AND s.sales_no = c.sales_no
              AND s.c_date_date = c.c_date_date
        )
        """
    )

    con.execute(
        """
        CREATE OR REPLACE VIEW v_receipt_base AS
        SELECT
            store_name,
            sales_no,
            c_date_date AS c_date,
            SUM(COALESCE(sub_total, 0)) AS sub_total,
            SUM(COALESCE(pro_disc_amt, 0)) AS pro_disc_amt,
            SUM(COALESCE(svc_amt, 0)) AS svc_amt,
            SUM(COALESCE(tax_amt, 0)) AS tax_all,
            SUM(CASE WHEN take_away_item = 'Y' THEN COALESCE(tax_amt, 0) ELSE 0 END) AS tax_takeaway,
            SUM(COALESCE(item_sub_total, COALESCE(sub_total, 0))) AS sum_item_sub_total,
            SUM(COALESCE(qty, 0)) AS sum_qty,
            MAX(CASE WHEN take_away_item = 'Y' THEN 1 ELSE 0 END) AS take_away_any_flag
        FROM v_sales_clean
        GROUP BY store_name, sales_no, c_date_date
        """
    )

    tax_case = (
        "base.tax_all"
        if current_database == "sushi_gogo_pos_live"
        else "base.tax_takeaway"
    )
    nett_rule = (
        "'GOGO_all_tax'"
        if current_database == "sushi_gogo_pos_live"
        else "'EXPRESS_takeaway_tax'"
    )

    con.execute(
        f"""
        CREATE OR REPLACE VIEW v_receipt_nett AS
        SELECT
            base.store_name,
            base.sales_no,
            base.c_date,
            base.sub_total,
            base.pro_disc_amt,
            base.svc_amt,
            base.tax_all,
            base.tax_takeaway,
            base.sum_item_sub_total,
            base.sum_qty,
            base.take_away_any_flag,
            ({tax_case}) AS tax_to_allocate,
            base.sub_total - base.pro_disc_amt + base.svc_amt - ({tax_case}) AS nett_sales,
            {nett_rule} AS nett_rule
        FROM v_receipt_base AS base
        """
    )

    con.execute(
        """
        CREATE OR REPLACE VIEW v_payment_prepared AS
        SELECT
            store_name,
            sales_no,
            c_date_date AS c_date,
            COALESCE(NULLIF(payment_name_norm, ''), 'UNKNOWN') AS payment_name,
            SUM(COALESCE(tender_amt, 0)) AS tender_amt
        FROM v_raw_payments
        GROUP BY
            store_name,
            sales_no,
            c_date_date,
            COALESCE(NULLIF(payment_name_norm, ''), 'UNKNOWN')
        """
    )

    con.execute(
        """
        CREATE OR REPLACE VIEW v_payment_non_cash_totals AS
        SELECT
            store_name,
            sales_no,
            c_date,
            SUM(tender_amt) AS non_cash_amt
        FROM v_payment_prepared
        WHERE payment_name <> 'CASH'
        GROUP BY store_name, sales_no, c_date
        """
    )

    con.execute(
        """
        CREATE OR REPLACE VIEW v_payment_alloc AS
        WITH base AS (
            SELECT
                r.store_name,
                r.sales_no,
                r.c_date,
                r.nett_sales,
                COALESCE(nc.non_cash_amt, 0) AS non_cash_amt
            FROM v_receipt_nett AS r
            LEFT JOIN v_payment_non_cash_totals AS nc
                ON r.store_name = nc.store_name
                AND r.sales_no = nc.sales_no
                AND r.c_date = nc.c_date
        )
        SELECT
            p.store_name,
            p.sales_no,
            p.c_date,
            p.payment_name,
            p.tender_amt AS allocated_amt,
            base.nett_sales AS receipt_nett
        FROM base
        JOIN v_payment_prepared AS p
            ON base.store_name = p.store_name
            AND base.sales_no = p.sales_no
            AND base.c_date = p.c_date
        WHERE p.payment_name <> 'CASH'
        UNION ALL
        SELECT
            base.store_name,
            base.sales_no,
            base.c_date,
            'CASH' AS payment_name,
            GREATEST(base.nett_sales - base.non_cash_amt, 0) AS allocated_amt,
            base.nett_sales AS receipt_nett
        FROM base
        """
    )

    con.execute(
        """
        CREATE OR REPLACE VIEW v_item_base AS
        SELECT
            store_name,
            sales_no,
            c_date_date AS c_date,
            item_name,
            SUM(COALESCE(item_sub_total, COALESCE(sub_total, 0))) AS item_sub_total,
            SUM(COALESCE(qty, 0)) AS qty,
            SUM(COALESCE(disc_amt, 0)) AS disc_amt,
            MAX(disc_name) AS disc_name,
            MAX(order_datetime) AS order_datetime,
            MAX(CASE WHEN take_away_item = 'Y' THEN 1 ELSE 0 END) AS take_away_flag
        FROM v_sales_clean
        GROUP BY store_name, sales_no, c_date_date, item_name
        """
    )

    con.execute(
        """
        CREATE OR REPLACE VIEW v_item_takeaway_totals AS
        SELECT
            store_name,
            sales_no,
            c_date,
            SUM(item_sub_total) AS takeaway_sub_total
        FROM v_item_base
        WHERE take_away_flag = 1
        GROUP BY store_name, sales_no, c_date
        """
    )

    con.execute(
        """
        CREATE OR REPLACE VIEW v_item_latest_category AS
        SELECT item_name, category_code
        FROM (
            SELECT
                item_name,
                category_code,
                ROW_NUMBER() OVER (
                    PARTITION BY item_name
                    ORDER BY m_date DESC NULLS LAST
                ) AS rn
            FROM item_master
            WHERE category_code IS NOT NULL
        )
        WHERE rn = 1
        """
    )

    con.execute(
        """
        CREATE OR REPLACE VIEW v_item_latest_category_all AS
        SELECT item_name, category_code FROM v_item_latest_category
        """
    )

    tax_allocation_case = (
        "CASE WHEN {tax_case} = 0 THEN 0 ELSE ({tax_case}) * ratio.share END"
        if current_database == "sushi_gogo_pos_live"
        else (
            "CASE "
            " WHEN receipt.tax_to_allocate = 0 THEN 0"
            " WHEN COALESCE(tt.takeaway_sub_total, 0) = 0 THEN 0"
            " WHEN item.take_away_flag = 1 THEN receipt.tax_to_allocate * (item.item_sub_total / tt.takeaway_sub_total)"
            " ELSE 0"
            " END"
        )
    ).format(tax_case="receipt.tax_to_allocate")

    con.execute(
        f"""
        CREATE OR REPLACE VIEW v_item_allocation AS
        WITH receipt AS (
            SELECT * FROM v_receipt_nett
        )
        SELECT
            item.store_name,
            item.sales_no,
            item.c_date,
            item.item_name,
            item.qty,
            item.disc_amt,
            item.disc_name,
            item.order_datetime,
            receipt.nett_sales,
            receipt.nett_sales AS receipt_nett,
            CASE
                WHEN receipt.sum_item_sub_total = 0 THEN 0
                ELSE ratio.share
            END AS item_sub_total_share,
            CASE
                WHEN receipt.sum_item_sub_total = 0 THEN 0
                ELSE item.item_sub_total
            END AS item_sub_total,
            CASE
                WHEN receipt.sum_item_sub_total = 0 THEN 0
                ELSE ratio.share * (receipt.sub_total - receipt.pro_disc_amt + receipt.svc_amt) - ({tax_allocation_case})
            END AS item_nett
        FROM v_item_base AS item
        JOIN receipt
            ON item.store_name = receipt.store_name
            AND item.sales_no = receipt.sales_no
            AND item.c_date = receipt.c_date
        LEFT JOIN v_item_takeaway_totals AS tt
            ON item.store_name = tt.store_name
            AND item.sales_no = tt.sales_no
            AND item.c_date = tt.c_date
        CROSS JOIN LATERAL (
            SELECT CASE
                WHEN receipt.sum_item_sub_total = 0 THEN 0
                ELSE item.item_sub_total / receipt.sum_item_sub_total
            END AS share
        ) AS ratio(share)
        """
    )


def materialize_daily_db(con: duckdb.DuckDBPyConnection, current_database: str) -> None:
    create_clean_views(con, current_database)

    con.execute(
        """
        CREATE OR REPLACE TABLE fact_daily_receipts AS
        SELECT
            c_date,
            store_name,
            sales_no,
            nett_sales,
            1 AS receipts,
            ? AS db_name,
            nett_rule
        FROM v_receipt_nett
        """,
        [current_database],
    )

    con.execute(
        """
        CREATE OR REPLACE TABLE fact_daily_payments AS
        SELECT
            c_date,
            store_name,
            sales_no,
            payment_name,
            allocated_amt,
            receipt_nett
        FROM v_payment_alloc
        """
    )

    con.execute(
        """
        CREATE OR REPLACE TABLE fact_daily_items AS
        SELECT
            c_date,
            store_name,
            sales_no,
            item_name,
            item_nett,
            receipt_nett,
            item_sub_total_share
        FROM v_item_allocation
        """
    )

    con.execute(
        """
        CREATE OR REPLACE VIEW vw_daily_store AS
        SELECT
            c_date,
            store_name,
            SUM(receipts) AS receipts,
            SUM(nett_sales) AS nett_sales
        FROM fact_daily_receipts
        GROUP BY c_date, store_name
        """
    )

    con.execute(
        """
        CREATE OR REPLACE VIEW vw_daily_payment AS
        SELECT
            c_date,
            payment_name,
            SUM(allocated_amt) AS allocated_amt
        FROM fact_daily_payments
        GROUP BY c_date, payment_name
        """
    )

    con.execute(
        """
        CREATE OR REPLACE VIEW vw_daily_item AS
        SELECT
            c_date,
            item_name,
            SUM(item_nett) AS item_nett
        FROM fact_daily_items
        GROUP BY c_date, item_name
        """
    )
