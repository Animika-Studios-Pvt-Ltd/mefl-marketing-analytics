"""
Read-only itemized comparison of all 43 Excel records against PostgreSQL.
Does not write or commit any changes.
"""

import os
import datetime
import openpyxl
import psycopg2
from dotenv import load_dotenv
from sync_campaign_data import (
    clean_csv_row,
    fetch_existing_records,
    KEY_COLUMNS,
    UPDATE_COLUMNS,
    values_differ,
)

BASE_DIR = r"C:\MEFL_data1"
EXCEL_PATH = r"C:\Users\india\OneDrive - ANIMIKA STUDIOS PRIVATE LIMITED\MEFL_Analysis folder\MEFL_Analysis.xlsx"

# 1. Load Excel in memory
wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
ws = wb["Meta ads"]
rows = list(ws.iter_rows(values_only=True))
header = rows[0]
target_cols = [
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
indices = {c: header.index(c) for c in target_cols}

excel_clean = []
for r in rows[1:]:
    if not any(cell is not None and str(cell).strip() != "" for cell in r):
        continue
    d = {}
    for c in target_cols:
        v = r[indices[c]]
        if c in ("Start Date", "End Date"):
            if isinstance(v, (datetime.datetime, datetime.date)):
                d[c] = f"{v.month}/{v.day}/{v.year}"
            elif v is not None and str(v).strip():
                raw_s = str(v).strip()
                parsed_d = None
                for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%Y/%m/%d"):
                    try:
                        parsed_d = datetime.datetime.strptime(raw_s, fmt).date()
                        break
                    except ValueError:
                        pass
                d[c] = f"{parsed_d.month}/{parsed_d.day}/{parsed_d.year}" if parsed_d else raw_s
            else:
                d[c] = ""
        else:
            d[c] = str(v).strip() if v is not None else ""
    if (
        d["Campaign No."] == "VE_C03"
        and d["Start Date"] in ("10/19/2026", "2026-10-19", "19-10-2026")
        and not d["Platform"]
    ):
        d["Platform"] = "Instagram"
    excel_clean.append(clean_csv_row(d))

# 2. Connect DB read-only
load_dotenv(os.path.join(BASE_DIR, ".env"))
conn = psycopg2.connect(
    host=os.getenv("DB_HOST"),
    port=os.getenv("DB_PORT"),
    dbname=os.getenv("DB_NAME"),
    user=os.getenv("DB_USER"),
    password=os.getenv("DB_PASSWORD"),
)
conn.set_session(readonly=True)
cur = conn.cursor()
existing_map = fetch_existing_records(cur)

print(
    f"=== DETAILED READ-ONLY RECORD BREAKDOWN ({len(excel_clean)} EXCEL ROWS VS {len(existing_map)} DB ROWS) ===\n"
)

inserts = []
updates = []
unchanged = []

for r in excel_clean:
    key = tuple(r[k] for k in KEY_COLUMNS)
    if key not in existing_map:
        inserts.append(r)
    else:
        db_rec = existing_map[key]
        changes = {}
        for c in UPDATE_COLUMNS:
            if values_differ(r[c], db_rec[c]):
                changes[c] = {"from_db": db_rec[c], "to_excel": r[c]}
        if changes:
            updates.append((r, changes))
        else:
            unchanged.append(r)

print(f"PROPOSED INSERTS ({len(inserts)} rows):")
for r in inserts:
    print(
        f"  + INSERT: Key=({r['campaign_no']}, {r['platform']}, {r['start_date']}, {r['end_date']}) | Name='{r['campaign_name']}' | Leads={r['leads']}"
    )

print(f"\nPROPOSED UPDATES ({len(updates)} rows):")
for r, ch in updates:
    print(
        f"  ~ UPDATE: Key=({r['campaign_no']}, {r['platform']}, {r['start_date']}, {r['end_date']}) | Changes={ch}"
    )

print(f"\nUNCHANGED MATCHES ({len(unchanged)} rows):")
for r in unchanged:
    print(
        f"  = MATCH: Key=({r['campaign_no']}, {r['platform']}, {r['start_date']}, {r['end_date']}) | Name='{r['campaign_name']}'"
    )

print(
    f"\nHISTORICAL DB ROWS PRESERVED (Present in DB with old keys, not in Excel):"
)
excel_keys = {tuple(r[k] for k in KEY_COLUMNS) for r in excel_clean}
for k, rec in existing_map.items():
    if k not in excel_keys:
        print(f"  * PRESERVED (Not deleted): Key={k} | Name='{rec['campaign_name']}'")

cur.close()
conn.close()
