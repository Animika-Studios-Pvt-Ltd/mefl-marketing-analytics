"""
Mother Earth Farmlands (MEFL) Marketing Data Synchronization Script
Target: MEFL_MARKETING_ANALYTICS.public.campaign_data_ready
Primary Source: Synced Excel at C:\\Users\\india\\OneDrive - ANIMIKA STUDIOS PRIVATE LIMITED\\MEFL_Analysis folder\\MEFL_Analysis.xlsx
Local CSV Cache: C:\\MEFL_data1\\MEFL_Analysis1.csv
"""

import os
import sys
import csv
import shutil
import logging
import argparse
import datetime
from decimal import Decimal, InvalidOperation
import re
import openpyxl
import psycopg2
from dotenv import load_dotenv

# Base paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOGS_DIR = os.path.join(BASE_DIR, "logs")
BACKUPS_DIR = os.path.join(BASE_DIR, "backups")
os.makedirs(LOGS_DIR, exist_ok=True)
os.makedirs(BACKUPS_DIR, exist_ok=True)
LOG_FILE = os.path.join(LOGS_DIR, "sync.log")

# Default source paths
DEFAULT_EXCEL_PATH = r"C:\Users\india\OneDrive - ANIMIKA STUDIOS PRIVATE LIMITED\MEFL_Analysis folder\MEFL_Analysis.xlsx"
DOWNLOADS_EXCEL_PATH = r"C:\Users\india\Downloads\MEFL_Analysis.xlsx"


def get_best_excel_path():
    candidate_paths = [
        r"C:\Users\india\Downloads\MEFL_Analysis.xlsx",
        r"C:\Users\india\Downloads\MEFL_Analysis (1).xlsx",
        r"C:\Users\india\Downloads\MEFL_Analysis (2).xlsx",
        DEFAULT_EXCEL_PATH,
    ]
    best_path = None
    best_mtime = -1
    for p in candidate_paths:
        if os.path.exists(p):
            m = os.path.getmtime(p)
            if m > best_mtime:
                best_mtime = m
                best_path = p
    if best_path and "Downloads" in best_path:
        try:
            shutil.copy2(best_path, DEFAULT_EXCEL_PATH)
            logger.info(f"Copied newer Excel from Downloads to: {DEFAULT_EXCEL_PATH}")
        except Exception as e:
            logger.warning(f"Could not copy from Downloads to OneDrive folder: {e}")
    return best_path
DEFAULT_CSV_PATH = os.path.join(BASE_DIR, "MEFL_Analysis1.csv")
DEFAULT_WEBSITE_LEADS_CSV = os.path.join(BASE_DIR, "website_leads_cleaned.csv")
DEFAULT_MEFL_VOICE_CSV = os.path.join(BASE_DIR, "mefl_voice_clean.csv")
DEFAULT_MEFL_RESULTS_CSV = os.path.join(BASE_DIR, "mefl_results_clean.csv")

# Logger setup
logger = logging.getLogger("mefl_sync")
logger.setLevel(logging.INFO)

if not logger.handlers:
    fh = logging.FileHandler(LOG_FILE, encoding="utf-8")
    fh.setLevel(logging.INFO)
    fh.setFormatter(
        logging.Formatter(
            "%(asctime)s [%(levelname)s] %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
    )
    logger.addHandler(fh)

    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))
    logger.addHandler(ch)

# 16 Target Columns expected in CSV and Database
CSV_HEADERS = [
    "Campaign No.",
    "Campaign_group",
    "Platform",
    "Form",
    "Campaign Name",
    "Start Date",
    "End Date",
    "Amount",
    "Reach",
    "Views",
    "Link Clicks",
    "Landing Page Views",
    "Impressions",
    "Leads",
    "Male",
    "Female",
]

COLUMN_MAPPING = {
    "Campaign No.": "campaign_no",
    "Campaign_group": "campaign_group",
    "Platform": "platform",
    "Form": "form",
    "Campaign Name": "campaign_name",
    "Start Date": "start_date",
    "End Date": "end_date",
    "Amount": "amount",
    "Reach": "reach",
    "Views": "views",
    "Link Clicks": "link_clicks",
    "Landing Page Views": "landing_page_views",
    "Impressions": "impressions",
    "Leads": "leads",
    "Male": "male",
    "Female": "female",
}

KEY_COLUMNS = ("campaign_no", "platform", "campaign_name", "start_date", "end_date")
UPDATE_COLUMNS = (
    "campaign_group",
    "form",
    "amount",
    "reach",
    "views",
    "link_clicks",
    "landing_page_views",
    "impressions",
    "leads",
    "male",
    "female",
)

DECIMAL_COLUMNS = {"amount"}
INTEGER_COLUMNS = {
    "reach",
    "views",
    "link_clicks",
    "landing_page_views",
    "impressions",
    "leads",
    "male",
    "female",
}
DATE_COLUMNS = {"start_date", "end_date"}


def backup_existing_csv(csv_path):
    """Creates a timestamped backup of the current CSV before modifying it."""
    if not os.path.exists(csv_path):
        return None
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_filename = f"MEFL_Analysis1_{timestamp}.csv"
    backup_path = os.path.join(BACKUPS_DIR, backup_filename)
    shutil.copy2(csv_path, backup_path)
    logger.info(f"Created backup of CSV at: {backup_path}")
    return backup_path


def extract_excel_to_csv(excel_path, target_csv_path, dry_run=False):
    """
    Safely extracts 14 target columns from Excel 'Meta ads' sheet.
    Validates rows, dates, columns, and composite keys before replacing the CSV.
    Safeguards Row 44 by strictly preserving Instagram for (VE_C03, 2026-10-19 to 2026-10-25).
    """
    if not os.path.exists(excel_path):
        raise FileNotFoundError(
            f"Synced Excel workbook not found at: {excel_path}"
        )

    logger.info(f"Opening Excel workbook: {excel_path}")
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    if "Meta ads" not in wb.sheetnames:
        raise ValueError(
            f"Expected sheet 'Meta ads' not found in Excel workbook. Found: {wb.sheetnames}"
        )

    ws = wb["Meta ads"]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError("Excel sheet 'Meta ads' is completely empty.")

    header_row = rows[0]
    # Check that all 12 target columns are present in Excel header
    missing_cols = [c for c in CSV_HEADERS if c not in header_row]
    if missing_cols:
        raise ValueError(
            f"Missing required columns in Excel 'Meta ads' sheet: {missing_cols}"
        )

    col_indices = {col: header_row.index(col) for col in CSV_HEADERS}

    extracted_records = []
    seen_keys = set()
    opening_spend_to_add = 0.0

    for row_idx, r in enumerate(rows[1:], start=2):
        # Skip completely blank lines
        if not any(cell is not None and str(cell).strip() != "" for cell in r):
            continue

        rec = {}
        for col_name in CSV_HEADERS:
            val = r[col_indices[col_name]]
            if col_name in ("Start Date", "End Date"):
                c_no_hint = str(r[col_indices["Campaign No."]] or "").strip()
                cg_hint = str(r[col_indices["Campaign_group"]] or "").strip()
                parsed_d = smart_parse_date(
                    val,
                    is_start_date=(col_name == "Start Date"),
                    campaign_group=cg_hint,
                    campaign_no=c_no_hint,
                )
                if parsed_d:
                    rec[col_name] = f"{parsed_d.month}/{parsed_d.day}/{parsed_d.year}"
                elif val is not None and str(val).strip():
                    rec[col_name] = str(val).strip()
                else:
                    rec[col_name] = ""
            elif col_name == "Amount":
                if val is not None:
                    s_amt = str(val).strip()
                    s_amt = re.sub(r"^[^\d.-]+", "", s_amt).replace(",", "").strip()
                    rec[col_name] = s_amt
                else:
                    rec[col_name] = ""
            elif val is not None:
                rec[col_name] = str(val).strip()
            else:
                rec[col_name] = ""

        # Special handling for authorized non-standard rows:
        # 1. Row 6: Opening Spend (Rs 3,513.39) - Merge directly into Campaign 3 Insta (IN_C03) per user request
        c_name = rec.get("Campaign Name", "").strip()
        c_no = rec.get("Campaign No.", "").strip()

        if not c_no and c_name == "Opening Spend":
            opening_spend_to_add = float(rec.get("Amount", 0) or 0)
            logger.info(
                f"Captured Opening Spend of Rs {opening_spend_to_add:.2f} from row {row_idx} to merge into IN_C03."
            )
            continue

        # Merge opening spend into IN_C03 (Campaign 3 Insta)
        if c_no == "IN_C03" and opening_spend_to_add > 0:
            base_amt = float(rec.get("Amount", 0) or 0)
            merged_amt = round(base_amt + opening_spend_to_add, 2)
            logger.info(
                f"Merged Opening Spend Rs {opening_spend_to_add:.2f} into IN_C03 (was Rs {base_amt:.2f}, now Rs {merged_amt:.2f})."
            )
            rec["Amount"] = f"{merged_amt:.2f}"
            opening_spend_to_add = 0.0

        # 2. Row 41: Google Ads intermediate spend (Rs 7,000.00)
        if not c_no and rec.get("Amount") == "7000" and "9/12/2026" in rec.get("End Date", ""):
            rec["Campaign No."] = "GO_ADJ"
            rec["Campaign_group"] = "C08"
            rec["Platform"] = "Google Ads"
            rec["Form"] = "static"
            rec["Campaign Name"] = "Google Ads Spend"
            if not rec.get("Start Date"):
                rec["Start Date"] = rec["End Date"]
            c_no = "GO_ADJ"
            logger.info(f"Including authorized Row {row_idx}: GO_ADJ (Google Ads, Rs 7,000.00)")

        # Skip rows without Campaign No. (e.g. general notes)
        if not c_no:
            logger.info(
                f"Skipping row {row_idx} without Campaign No. ('{rec.get('Campaign Name', '')}')"
            )
            continue

        s_dt = rec.get("Start Date", "")
        e_dt = rec.get("End Date", "")
        plat = rec.get("Platform", "")

        if not plat:
            # Authorized specific exception for VE_C03 row 44 (2026-10-19 to 2026-10-25)
            if (
                c_no == "VE_C03"
                and s_dt in ("10/19/2026", "2026-10-19", "19-10-2026")
                and e_dt in ("10/25/2026", "2026-10-25", "25-10-2026")
            ):
                logger.info(
                    f"Applying authorized safeguard for row {row_idx}: Preserving 'Instagram' platform for VE_C03 (10/19/2026 to 10/25/2026)."
                )
                rec["Platform"] = "Instagram"
                plat = "Instagram"
            else:
                raise ValueError(
                    f"Safety Violation: Row {row_idx} ({c_no}, {s_dt} to {e_dt}) has an unhandled BLANK Platform. Halting sync."
                )

        # Validate natural key uniqueness (including Campaign Name to support multiple ad sets like GO_C09)
        c_name = rec.get("Campaign Name", "").strip()
        key = (c_no, plat, c_name, s_dt, e_dt)
        if key in seen_keys:
            raise ValueError(
                f"Duplicate natural key detected in Excel row {row_idx}: {key}"
            )
        seen_keys.add(key)
        extracted_records.append(rec)

    # Validation: Ensure row count matches expectation (at least 40 active campaign rows)
    if len(extracted_records) < 40:
        raise ValueError(
            f"Extracted only {len(extracted_records)} rows from Excel, expected at least 40."
        )

    logger.info(
        f"Successfully extracted and validated {len(extracted_records)} records from Excel."
    )

    if not dry_run:
        # 1. Back up existing CSV
        backup_existing_csv(target_csv_path)

        # 2. Write updated records directly to target CSV
        with open(target_csv_path, mode="w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=CSV_HEADERS)
            writer.writeheader()
            writer.writerows(extracted_records)

        logger.info(f"Successfully updated CSV at: {target_csv_path}")
    else:
        logger.info(
            f"[DRY-RUN] CSV update skipped (would write {len(extracted_records)} records to {target_csv_path})."
        )

    return extracted_records


def smart_parse_date(val_str, is_start_date=True, campaign_group=None, campaign_no=None):
    if val_str is None or not str(val_str).strip():
        return None
    cg = str(campaign_group or "").strip()
    cno = str(campaign_no or "").strip()

    if isinstance(val_str, (datetime.datetime, datetime.date)):
        d = val_str.date() if isinstance(val_str, datetime.datetime) else val_str
        # Fix known typo for C15/C09 Diwali campaigns
        if cno.startswith(("VE-C09", "VE-C10", "IN-C15", "FB_C15")):
            if is_start_date:
                return datetime.date(2026, 10, 29)
            else:
                return datetime.date(2026, 11, 4)
        # Fix Excel auto-parsing D/11/2026 as Month D, Day 11 instead of Day D, Month 11
        if d.year == 2026 and d.day == 11 and d.month in (4, 6, 10, 12) and (
            cg in ("C06", "C07", "C16")
            or cno.startswith(("VE-C11", "VE-C12", "VE_C13", "IN-C16", "FB-C16"))
        ):
            return datetime.date(2026, 11, d.month)
        return d

    s = str(val_str).strip()

    # Specific known typo fix
    if cno.startswith(("VE-C09", "VE-C10", "IN-C15", "FB_C15")):
        if is_start_date:
            return datetime.date(2026, 10, 29)
        else:
            return datetime.date(2026, 11, 4)

    # For day-first formats with slash or hyphen (DD/MM/YYYY or DD-MM-YYYY)
    m = re.match(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$", s)
    if m:
        p1, p2, yr = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if p1 > 12:  # unambiguously DD-MM-YYYY
            return datetime.date(yr, p2, p1)
        elif p2 == 11 and yr == 2026 and (
            cg in ("C06", "C07", "C12", "C13", "C16", "C17")
            or cno.startswith(("VE-C11", "VE-C12", "VE_C13", "IN-C16", "FB-C16", "IN_C17", "FB_C17"))
        ):
            # November campaign with DD/11/2026
            return datetime.date(yr, p2, p1)
        elif p2 > 12:  # MM-DD-YYYY
            return datetime.date(yr, p1, p2)
        else:
            # Default for older campaigns (7/4/2026, 10/11/2026): MM/DD/YYYY
            return datetime.date(yr, p1, p2)

    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%m/%d/%Y", "%Y/%m/%d"):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    raise ValueError(
        f"Invalid date format '{s}'. Expected DD/MM/YYYY, MM/DD/YYYY, YYYY-MM-DD, or DD-MM-YYYY."
    )


def parse_date(val_str, is_start_date=True, campaign_group=None, campaign_no=None):
    return smart_parse_date(val_str, is_start_date, campaign_group, campaign_no)


def parse_decimal(val_str):
    if val_str is None or str(val_str).strip() == "":
        return None
    s = str(val_str).strip()
    s = re.sub(r"^[^\d.-]+", "", s).replace(",", "")
    if not s:
        return None
    try:
        return Decimal(s)
    except (InvalidOperation, ValueError) as e:
        raise ValueError(f"Invalid decimal value '{val_str}': {e}")


def parse_integer(val_str):
    if val_str is None or str(val_str).strip() == "":
        return None
    s = str(val_str).strip()
    s = re.sub(r"^[^\d.-]+", "", s).replace(",", "")
    if not s:
        return None
    try:
        return int(round(float(s)))
    except ValueError as e:
        raise ValueError(f"Invalid integer value '{val_str}': {e}")


def clean_csv_row(raw_row):
    record = {}
    for csv_header, db_col in COLUMN_MAPPING.items():
        raw_val = raw_row.get(csv_header, "")
        if db_col in DATE_COLUMNS:
            record[db_col] = parse_date(
                raw_val,
                is_start_date=(db_col == "start_date"),
                campaign_group=raw_row.get("Campaign_group", ""),
                campaign_no=raw_row.get("Campaign No.", ""),
            )
        elif db_col in DECIMAL_COLUMNS:
            record[db_col] = parse_decimal(raw_val)
        elif db_col in INTEGER_COLUMNS:
            record[db_col] = parse_integer(raw_val)
        else:
            cleaned_str = raw_val.strip() if raw_val is not None else ""
            if db_col == "form" and cleaned_str.lower() in ("vedio", "video"):
                cleaned_str = "video"
            record[db_col] = cleaned_str if cleaned_str != "" else None

    for k in KEY_COLUMNS:
        if k != "campaign_name" and record[k] is None:
            raise ValueError(
                f"Missing composite key column '{k}' in row: {raw_row}"
            )
    return record


def values_differ(v1, v2):
    if v1 is None and v2 is None:
        return False
    if v1 is None or v2 is None:
        return True
    if isinstance(v1, Decimal) or isinstance(v2, Decimal):
        try:
            return Decimal(str(v1)) != Decimal(str(v2))
        except Exception:
            return v1 != v2
    if isinstance(v1, str) and isinstance(v2, str):
        return v1.strip() != v2.strip()
    return v1 != v2


def get_target_databases():
    """
    Returns list of target databases to sync to based on TARGET_DB env variable.
    Options: 'BOTH' -> ['LOCAL', 'NEON'], 'NEON' -> ['NEON'], 'LOCAL' -> ['LOCAL']
    """
    env_path = os.path.join(BASE_DIR, ".env")
    if os.path.exists(env_path):
        load_dotenv(env_path)
    mode = os.getenv("TARGET_DB", "BOTH").strip().upper()
    if mode == "BOTH":
        targets = []
        if os.getenv("DB_PASSWORD") or os.getenv("DB_HOST"):
            targets.append("LOCAL")
        if os.getenv("NEON_DATABASE_URL") or os.getenv("NEON_DB_HOST"):
            targets.append("NEON")
        return targets if targets else ["LOCAL"]
    elif mode == "NEON":
        return ["NEON"]
    else:
        return ["LOCAL"]


def get_db_connection(target=None):
    """
    Returns a psycopg2 connection to either 'LOCAL' PostgreSQL or 'NEON' Serverless PostgreSQL.
    """
    env_path = os.path.join(BASE_DIR, ".env")
    if os.path.exists(env_path):
        load_dotenv(env_path)

    target_name = (target or os.getenv("TARGET_DB", "LOCAL")).strip().upper()
    if target_name == "NEON":
        neon_url = os.getenv("NEON_DATABASE_URL")
        if neon_url:
            return psycopg2.connect(neon_url, connect_timeout=15)
        return psycopg2.connect(
            host=os.getenv("NEON_DB_HOST"),
            port=os.getenv("NEON_DB_PORT", "5432"),
            dbname=os.getenv("NEON_DB_NAME", "neondb"),
            user=os.getenv("NEON_DB_USER", "neondb_owner"),
            password=os.getenv("NEON_DB_PASSWORD"),
            sslmode="require",
            connect_timeout=15,
        )

    db_host = os.getenv("DB_HOST", "localhost")
    db_port = os.getenv("DB_PORT", "5432")
    db_name = os.getenv("DB_NAME", "MEFL_MARKETING_ANALYTICS")
    db_user = os.getenv("DB_USER", "postgres")
    db_password = os.getenv("DB_PASSWORD")

    if db_name.strip().upper() != "MEFL_MARKETING_ANALYTICS":
        raise ValueError(
            f"Safety block: DB_NAME must be 'MEFL_MARKETING_ANALYTICS', but got '{db_name}'."
        )

    return psycopg2.connect(
        host=db_host,
        port=db_port,
        dbname=db_name,
        user=db_user,
        password=db_password,
        connect_timeout=10,
    )


def fetch_existing_records(cur, target_table="campaign_data_ready"):
    all_cols = list(KEY_COLUMNS) + list(UPDATE_COLUMNS)
    cols_sql = ", ".join(f'"{c}"' for c in all_cols)
    cur.execute(f"SELECT {cols_sql} FROM public.{target_table};")
    rows = cur.fetchall()

    existing_map = {}
    for r in rows:
        rec = dict(zip(all_cols, r))
        key = tuple(rec[k] for k in KEY_COLUMNS)
        existing_map[key] = rec
    return existing_map


def sync(csv_path, dry_run=False, excel_path=None, target_db=None):
    if target_db is None:
        targets = get_target_databases()
        if len(targets) > 1:
            res = {}
            for t in targets:
                logger.info(f"--- Running campaign_data_ready sync for target database: [{t}] ---")
                res[t] = sync(csv_path=csv_path, dry_run=dry_run, excel_path=excel_path, target_db=t)
            return res
        target_db = targets[0]

    start_time = datetime.datetime.now()
    mode_text = "[DRY-RUN MODE]" if dry_run else "[LIVE SYNC MODE]"
    logger.info(f"{'='*60}")
    logger.info(f"Starting MEFL Campaign Data Sync [{target_db}] {mode_text}")

    records = []
    # If excel_path provided and exists, use the freshly extracted records in memory
    if excel_path and os.path.exists(excel_path):
        logger.info(f"Source Excel: {excel_path}")
        extracted_raw_records = extract_excel_to_csv(
            excel_path=excel_path, target_csv_path=csv_path, dry_run=dry_run
        )
        for raw_row in extracted_raw_records:
            records.append(clean_csv_row(raw_row))
    else:
        logger.info(f"Source CSV: {csv_path}")
        with open(csv_path, mode="r", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for raw_row in reader:
                if not any(v.strip() for v in raw_row.values() if v):
                    continue
                records.append(clean_csv_row(raw_row))

    logger.info(f"Evaluating {len(records)} records against PostgreSQL [{target_db}].")
    logger.info(f"Target DB: [{target_db}].public.campaign_data_ready")

    conn = None
    inserted_count = 0
    updated_count = 0
    unchanged_count = 0

    try:
        conn = get_db_connection(target_db)
        conn.autocommit = False
        cur = conn.cursor()

        existing_map = fetch_existing_records(cur)
        logger.info(
            f"Current records in public.campaign_data_ready: {len(existing_map)}"
        )

        inserts_to_execute = []
        updates_to_execute = []

        for rec in records:
            key = tuple(rec[k] for k in KEY_COLUMNS)
            if key not in existing_map:
                inserts_to_execute.append(rec)
            else:
                existing_rec = existing_map[key]
                changes = {}
                for col in UPDATE_COLUMNS:
                    csv_v = rec[col]
                    db_v = existing_rec[col]
                    if values_differ(csv_v, db_v):
                        changes[col] = {"from": db_v, "to": csv_v}

                if changes:
                    updates_to_execute.append((rec, changes))
                else:
                    unchanged_count += 1

        if inserts_to_execute:
            logger.info(
                f"\n--- Proposed INSERTs ({len(inserts_to_execute)}) ---"
            )
            for rec in inserts_to_execute:
                logger.info(
                    f"INSERT -> Key: ({rec['campaign_no']}, {rec['platform']}, {rec['start_date']}, {rec['end_date']}) "
                    f"| Name: '{rec['campaign_name']}'"
                )
                if not dry_run:
                    cols = list(rec.keys())
                    vals = [rec[c] for c in cols]
                    cols_str = ", ".join(f'"{c}"' for c in cols)
                    placeholders = ", ".join(["%s"] * len(cols))
                    insert_sql = f"INSERT INTO public.campaign_data_ready ({cols_str}) VALUES ({placeholders});"
                    cur.execute(insert_sql, vals)
            inserted_count = len(inserts_to_execute)
        else:
            logger.info("No new records to insert.")

        if updates_to_execute:
            logger.info(
                f"\n--- Proposed UPDATEs ({len(updates_to_execute)}) ---"
            )
            for rec, changes in updates_to_execute:
                change_desc = ", ".join(
                    f"{col}: {c['from']} -> {c['to']}"
                    for col, c in changes.items()
                )
                logger.info(
                    f"UPDATE -> Key: ({rec['campaign_no']}, {rec['platform']}, {rec['campaign_name']}, {rec['start_date']}, {rec['end_date']}) "
                    f"| Changes: [{change_desc}]"
                )
                if not dry_run:
                    set_clauses = [f'"{col}" = %s' for col in changes.keys()]
                    vals = [rec[col] for col in changes.keys()]
                    set_str = ", ".join(set_clauses)
                    update_sql = f"""
                        UPDATE public.campaign_data_ready 
                        SET {set_str}
                        WHERE "campaign_no" = %s 
                          AND "platform" = %s 
                          AND ("campaign_name" = %s OR ("campaign_name" IS NULL AND %s IS NULL))
                          AND "start_date" = %s 
                          AND "end_date" = %s;
                    """
                    vals.extend([
                        rec["campaign_no"],
                        rec["platform"],
                        rec["campaign_name"],
                        rec["campaign_name"],
                        rec["start_date"],
                        rec["end_date"],
                    ])
                    cur.execute(update_sql, vals)
            updated_count = len(updates_to_execute)
        else:
            logger.info("No existing records required updates.")

        logger.info(f"Unchanged records (already matching): {unchanged_count}")

        # Prune obsolete/deleted records not in source so DB strictly matches current data
        csv_keys = {tuple(r[k] for k in KEY_COLUMNS) for r in records}
        db_only_keys = set(existing_map.keys()) - csv_keys
        deleted_count = 0
        if db_only_keys:
            logger.info(
                f"\n--- Pruning obsolete records not in source ({len(db_only_keys)}) ---"
            )
            for key_tuple in db_only_keys:
                c_no, plat, c_name, s_dt, e_dt = key_tuple
                logger.info(
                    f"DELETE -> ({c_no}, {plat}, '{c_name}', {s_dt} to {e_dt})"
                )
                if not dry_run:
                    delete_sql = """
                        DELETE FROM public.campaign_data_ready 
                        WHERE "campaign_no" = %s 
                          AND "platform" = %s 
                          AND ("campaign_name" = %s OR ("campaign_name" IS NULL AND %s IS NULL))
                          AND "start_date" = %s 
                          AND "end_date" = %s;
                    """
                    cur.execute(delete_sql, [c_no, plat, c_name, c_name, s_dt, e_dt])
            deleted_count = len(db_only_keys)

        if dry_run:
            conn.rollback()
            logger.info(
                "\n[DRY-RUN COMPLETE] Database transaction rolled back. Zero changes committed."
            )
        else:
            conn.commit()
            logger.info(
                "\n[LIVE SYNC COMPLETE] Database transaction committed successfully."
            )

        cur.close()

    except Exception as e:
        if conn:
            conn.rollback()
            logger.error(
                "[TRANSACTION ERROR] Transaction rolled back due to error."
            )
        logger.exception(f"Synchronization failed: {e}")
        raise
    finally:
        if conn:
            conn.close()

    duration = datetime.datetime.now() - start_time
    logger.info(
        f"Summary: {inserted_count} inserted, {updated_count} updated, {unchanged_count} unchanged."
    )
    logger.info(f"Total Execution Time: {duration.total_seconds():.3f} seconds")
    logger.info(f"{'='*60}\n")
    return {
        "inserted": inserted_count,
        "updated": updated_count,
        "unchanged": unchanged_count,
        "duration_sec": duration.total_seconds(),
    }


def sync_mefl_leads(csv_path, target_table, dry_run=False, target_db=None):
    """
    Synchronizes mefl_voice or mefl_results datasets into PostgreSQL.
    Key column: 'lead' (bigint, unique primary key).
    Updates only modified records, inserts new leads, preserves historical leads.
    """
    if target_db is None:
        targets = get_target_databases()
        if len(targets) > 1:
            res = {}
            for t in targets:
                logger.info(f"--- Running {target_table} sync for target database: [{t}] ---")
                res[t] = sync_mefl_leads(csv_path=csv_path, target_table=target_table, dry_run=dry_run, target_db=t)
            return res
        target_db = targets[0]

    start_time = datetime.datetime.now()
    mode_text = "[DRY-RUN MODE]" if dry_run else "[LIVE SYNC MODE]"
    logger.info(f"{'='*60}")
    logger.info(f"Starting {target_table} Sync [{target_db}] from {os.path.basename(csv_path)} {mode_text}")

    if not os.path.exists(csv_path):
        logger.warning(f"Source CSV not found for {target_table} at {csv_path}. Skipping.")
        return {"inserted": 0, "updated": 0, "unchanged": 0}

    if target_table in ("mefl_results", "mefl_voice"):
        lead_columns = [
            "sl_no",
            "lead",
            "created_date",
            "subject",
            "customer_name",
            "user",
            "status",
            "channel",
            "sub_channel",
            "last_note",
            "class",
            "next_step",
        ]
    else:
        lead_columns = [
            "sl_no",
            "lead",
            "created_date",
            "subject",
            "customer_name",
            "user",
            "status",
            "channel",
            "sub_channel",
            "last_note",
            "class",
            "next_step",
        ]
    non_key_cols = [c for c in lead_columns if c != "lead"]

    records = []
    seen_leads = set()
    with open(csv_path, mode="r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row_idx, r in enumerate(reader, start=2):
            if not any(v.strip() for v in r.values() if v):
                continue
            raw_lead = r.get("lead", "").strip()
            if not raw_lead:
                raise ValueError(f"Missing 'lead' key in {csv_path} row {row_idx}: {r}")
            lead_val = int(raw_lead)
            if lead_val in seen_leads:
                raise ValueError(f"Duplicate 'lead' key {lead_val} detected in {csv_path} row {row_idx}")
            seen_leads.add(lead_val)

            rec = {"lead": lead_val}
            # sl_no
            sl_val = r.get("sl_no", "").strip()
            rec["sl_no"] = int(sl_val) if sl_val else None

            # created_date
            dt_val = r.get("created_date", "").strip()
            if dt_val:
                dt_obj = None
                for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d", "%d-%m-%Y %H:%M:%S", "%d/%m/%Y %H:%M:%S"):
                    try:
                        dt_obj = datetime.datetime.strptime(dt_val, fmt)
                        break
                    except ValueError:
                        pass
                rec["created_date"] = dt_obj if dt_obj else dt_val
            else:
                rec["created_date"] = None

            # text fields
            text_candidates = (
                "subject",
                "customer_name",
                "user",
                "status",
                "channel",
                "sub_channel",
                "last_note",
                "class",
                "status2",
                "next_step",
            )
            for col in text_candidates:
                if col in lead_columns:
                    val = r.get(col, "")
                    rec[col] = val.strip() if val is not None and val.strip() != "" else None

            records.append(rec)

    logger.info(f"Loaded {len(records)} records from {csv_path}. Target: [{target_db}].public.{target_table}")

    conn = None
    inserted_count = 0
    updated_count = 0
    unchanged_count = 0

    try:
        conn = get_db_connection(target_db)
        conn.autocommit = False
        cur = conn.cursor()

        cols_sql = ", ".join(f'"{c}"' for c in lead_columns)
        cur.execute(f"SELECT {cols_sql} FROM public.{target_table};")
        db_rows = cur.fetchall()

        existing_map = {}
        for row in db_rows:
            db_rec = dict(zip(lead_columns, row))
            existing_map[db_rec["lead"]] = db_rec

        logger.info(f"Current records in public.{target_table}: {len(existing_map)}")

        inserts_to_execute = []
        updates_to_execute = []

        for rec in records:
            lead_val = rec["lead"]
            if lead_val not in existing_map:
                inserts_to_execute.append(rec)
            else:
                db_rec = existing_map[lead_val]
                changes = {}
                for col in non_key_cols:
                    csv_v = rec[col]
                    db_v = db_rec[col]
                    # Safeguard: Do not overwrite existing non-null optional values with blank
                    if col in ("subject", "customer_name", "next_step", "class", "last_note", "sub_channel") and csv_v is None and db_v is not None:
                        continue
                    if values_differ(csv_v, db_v):
                        changes[col] = {"from": db_v, "to": csv_v}

                if changes:
                    updates_to_execute.append((rec, changes))
                else:
                    unchanged_count += 1

        if inserts_to_execute:
            logger.info(f"\n--- Proposed INSERTs for {target_table} ({len(inserts_to_execute)}) ---")
            for rec in inserts_to_execute:
                logger.info(f"INSERT -> lead: {rec['lead']} | name: {rec['customer_name']} | status: {rec['status']}")
                if not dry_run:
                    cols = list(rec.keys())
                    vals = [rec[c] for c in cols]
                    cols_str = ", ".join(f'"{c}"' for c in cols)
                    placeholders = ", ".join(["%s"] * len(cols))
                    insert_sql = f"INSERT INTO public.{target_table} ({cols_str}) VALUES ({placeholders});"
                    cur.execute(insert_sql, vals)
            inserted_count = len(inserts_to_execute)
        else:
            logger.info(f"No new records to insert into {target_table}.")

        if updates_to_execute:
            logger.info(f"\n--- Proposed UPDATEs for {target_table} ({len(updates_to_execute)}) ---")
            for rec, changes in updates_to_execute:
                change_desc = ", ".join(f"{col}: {c['from']} -> {c['to']}" for col, c in changes.items())
                logger.info(f"UPDATE -> lead: {rec['lead']} | Changes: [{change_desc}]")
                if not dry_run:
                    set_clauses = [f'"{col}" = %s' for col in changes.keys()]
                    vals = [rec[col] for col in changes.keys()]
                    set_str = ", ".join(set_clauses)
                    update_sql = f'UPDATE public.{target_table} SET {set_str} WHERE "lead" = %s;'
                    vals.append(rec["lead"])
                    cur.execute(update_sql, vals)
            updated_count = len(updates_to_execute)
        else:
            logger.info(f"No existing records required updates in {target_table}.")

        logger.info(f"Unchanged records in {target_table}: {unchanged_count}")

        # Safety: historical rows in DB not in CSV are preserved
        csv_leads = {r["lead"] for r in records}
        db_only_leads = set(existing_map.keys()) - csv_leads
        if db_only_leads:
            logger.info(f"Existing DB records not in source {target_table} (safely preserved, NOT deleted): {len(db_only_leads)}")

        if dry_run:
            conn.rollback()
            logger.info(f"[DRY-RUN COMPLETE] Database transaction rolled back for {target_table}.")
        else:
            conn.commit()
            logger.info(f"[LIVE SYNC COMPLETE] Database transaction committed successfully for {target_table}.")

        cur.close()

    except Exception as e:
        if conn:
            conn.rollback()
            logger.error(f"[TRANSACTION ERROR] Transaction rolled back for {target_table} due to error.")
        logger.exception(f"Synchronization failed for {target_table}: {e}")
        raise
    finally:
        if conn:
            conn.close()

    duration = datetime.datetime.now() - start_time
    logger.info(f"Summary for {target_table}: {inserted_count} inserted, {updated_count} updated, {unchanged_count} unchanged.")
    logger.info(f"Execution Time for {target_table}: {duration.total_seconds():.3f}s")
    logger.info(f"{'='*60}\n")
    return {
        "inserted": inserted_count,
        "updated": updated_count,
        "unchanged": unchanged_count,
        "duration_sec": duration.total_seconds(),
    }


def sync_website_leads(csv_path, dry_run=False, target_db=None):
    """
    Synchronizes website_leads_cleaned.csv into public.website_leads.
    Preserves all existing records (no deletions).
    Performs safe duplicate-aware multiset comparison:
    Only rows that appear more times in the CSV than in the database are inserted.
    """
    if target_db is None:
        targets = get_target_databases()
        if len(targets) > 1:
            res = {}
            for t in targets:
                logger.info(f"--- Running website_leads sync for target database: [{t}] ---")
                res[t] = sync_website_leads(csv_path=csv_path, dry_run=dry_run, target_db=t)
            return res
        target_db = targets[0]

    from collections import Counter
    start_time = datetime.datetime.now()
    mode_text = "[DRY-RUN MODE]" if dry_run else "[LIVE SYNC MODE]"
    logger.info(f"{'='*60}")
    logger.info(f"Starting website_leads Sync [{target_db}] from {os.path.basename(csv_path)} {mode_text}")

    if not os.path.exists(csv_path):
        logger.warning(f"Source CSV not found for website_leads at {csv_path}. Skipping.")
        return {"inserted": 0, "unchanged": 0}

    # Read and parse CSV records
    csv_records = []
    with open(csv_path, mode="r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for r in reader:
            if not any(v.strip() for v in r.values() if v):
                continue
            dt = parse_date(r.get("created_date", ""))
            src = r.get("source", "").strip() or None
            proj = r.get("project_interest", "").strip() or None
            st = r.get("status", "").strip() or None
            csv_records.append((dt, src, proj, st))

    logger.info(f"Loaded {len(csv_records)} rows from {csv_path}.")

    conn = None
    inserted_count = 0
    unchanged_count = 0

    try:
        conn = get_db_connection(target_db)
        conn.autocommit = False
        cur = conn.cursor()

        cur.execute('SELECT "created_date", "source", "project_interest", "status" FROM public.website_leads;')
        db_rows = cur.fetchall()
        logger.info(f"Current records in public.website_leads: {len(db_rows)}")

        csv_counter = Counter(csv_records)
        db_counter = Counter(db_rows)

        # Difference: instances in CSV exceeding instances in DB
        to_insert_counter = csv_counter - db_counter
        total_new_to_insert = sum(to_insert_counter.values())

        # Unchanged instances (matched between CSV and DB)
        matched_counter = csv_counter & db_counter
        unchanged_count = sum(matched_counter.values())

        if total_new_to_insert > 0:
            logger.info(f"\n--- Proposed INSERTs for website_leads ({total_new_to_insert}) ---")
            for row_tuple, count in to_insert_counter.items():
                logger.info(f"INSERT ({count}x) -> date: {row_tuple[0]}, source: {row_tuple[1]}, project: {row_tuple[2]}, status: {row_tuple[3]}")
                if not dry_run:
                    for _ in range(count):
                        cur.execute(
                            'INSERT INTO public.website_leads ("created_date", "source", "project_interest", "status") VALUES (%s, %s, %s, %s);',
                            row_tuple,
                        )
            inserted_count = total_new_to_insert
        else:
            logger.info("No new records to insert into website_leads (all source rows already present).")

        logger.info(f"Unchanged records in website_leads: {unchanged_count}")

        # Safety: check if DB has records not in CSV (safely preserved, never deleted)
        db_surplus = db_counter - csv_counter
        surplus_count = sum(db_surplus.values())
        if surplus_count > 0:
            logger.info(f"Existing DB records exceeding source in website_leads (safely preserved, NOT deleted): {surplus_count}")

        if dry_run:
            conn.rollback()
            logger.info("[DRY-RUN COMPLETE] Database transaction rolled back for website_leads.")
        else:
            conn.commit()
            logger.info("[LIVE SYNC COMPLETE] Database transaction committed successfully for website_leads.")

        cur.close()

    except Exception as e:
        if conn:
            conn.rollback()
            logger.error("[TRANSACTION ERROR] Transaction rolled back for website_leads due to error.")
        logger.exception(f"Synchronization failed for website_leads: {e}")
        raise
    finally:
        if conn:
            conn.close()

    duration = datetime.datetime.now() - start_time
    logger.info(f"Summary for website_leads: {inserted_count} inserted, {unchanged_count} unchanged.")
    logger.info(f"Execution Time for website_leads: {duration.total_seconds():.3f}s")
    logger.info(f"{'='*60}\n")
    return {
        "inserted": inserted_count,
        "unchanged": unchanged_count,
        "duration_sec": duration.total_seconds(),
    }


def sync_all(
    excel_path=None,
    campaign_csv_path=DEFAULT_CSV_PATH,
    website_leads_csv=DEFAULT_WEBSITE_LEADS_CSV,
    mefl_voice_csv=DEFAULT_MEFL_VOICE_CSV,
    mefl_results_csv=DEFAULT_MEFL_RESULTS_CSV,
    dry_run=False,
):
    """
    Executes full synchronization for all 4 marketing analytics datasets.
    """
    if excel_path is None:
        excel_path = get_best_excel_path()
    results = {}
    logger.info("Starting synchronization of all 4 datasets...")

    # 1. Existing campaign data sync (100% untouched logic)
    results["campaign_data_ready"] = sync(
        csv_path=campaign_csv_path, dry_run=dry_run, excel_path=excel_path
    )

    # 2. MEFL Voice Leads
    results["mefl_voice"] = sync_mefl_leads(
        csv_path=mefl_voice_csv, target_table="mefl_voice", dry_run=dry_run
    )

    # 3. MEFL Results Leads
    results["mefl_results"] = sync_mefl_leads(
        csv_path=mefl_results_csv, target_table="mefl_results", dry_run=dry_run
    )

    # 4. Website Leads
    results["website_leads"] = sync_website_leads(
        csv_path=website_leads_csv, dry_run=dry_run
    )

    # 5. Automated Cloud Backup to OneDrive
    if not dry_run:
        try:
            from backup_database import create_database_backup
            create_database_backup()
        except Exception as e:
            logger.warning(f"Automated database backup skipped or failed: {e}")

    return results


def main():
    parser = argparse.ArgumentParser(
        description="Synchronize Mother Earth Farmlands campaign and leads data."
    )
    parser.add_argument(
        "--excel",
        type=str,
        default=DEFAULT_EXCEL_PATH,
        help="Path to the synced Excel workbook to extract before syncing.",
    )
    parser.add_argument(
        "--csv",
        type=str,
        default=DEFAULT_CSV_PATH,
        help="Path to the target CSV file (default: MEFL_Analysis1.csv)",
    )
    parser.add_argument(
        "--website-leads-csv",
        type=str,
        default=DEFAULT_WEBSITE_LEADS_CSV,
        help="Path to website leads CSV (default: website_leads_cleaned.csv)",
    )
    parser.add_argument(
        "--voice-csv",
        type=str,
        default=DEFAULT_MEFL_VOICE_CSV,
        help="Path to MEFL voice CSV (default: mefl_voice_clean.csv)",
    )
    parser.add_argument(
        "--results-csv",
        type=str,
        default=DEFAULT_MEFL_RESULTS_CSV,
        help="Path to MEFL results CSV (default: mefl_results_clean.csv)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Simulate run without writing to CSV or PostgreSQL.",
    )
    parser.add_argument(
        "--only-campaigns",
        action="store_true",
        help="Run only the campaign_data_ready sync.",
    )
    parser.add_argument(
        "--only-leads",
        action="store_true",
        help="Run only the leads syncs (voice, results, website leads).",
    )
    args = parser.parse_args()

    excel_target = args.excel if (args.excel and os.path.exists(args.excel) and args.excel != DEFAULT_EXCEL_PATH) else get_best_excel_path()
    try:
        if args.only_campaigns:
            sync(csv_path=args.csv, dry_run=args.dry_run, excel_path=excel_target)
        elif args.only_leads:
            sync_mefl_leads(csv_path=args.voice_csv, target_table="mefl_voice", dry_run=args.dry_run)
            sync_mefl_leads(csv_path=args.results_csv, target_table="mefl_results", dry_run=args.dry_run)
            sync_website_leads(csv_path=args.website_leads_csv, dry_run=args.dry_run)
        else:
            sync_all(
                excel_path=excel_target,
                campaign_csv_path=args.csv,
                website_leads_csv=args.website_leads_csv,
                mefl_voice_csv=args.voice_csv,
                mefl_results_csv=args.results_csv,
                dry_run=args.dry_run,
            )
    except Exception:
        sys.exit(1)


if __name__ == "__main__":
    main()
