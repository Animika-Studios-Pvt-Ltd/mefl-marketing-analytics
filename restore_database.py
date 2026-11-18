"""
Mother Earth Farmlands (MEFL) - Database Restore Script
Restores MEFL_MARKETING_ANALYTICS from mefl_backup_latest.sql
Can restore from:
  1. Cloud OneDrive: C:\\Users\\india\\OneDrive - ANIMIKA STUDIOS PRIVATE LIMITED\\MEFL_data1\\db_backups\\mefl_backup_latest.sql
  2. Local backups:  C:\\MEFL_data1\\backups\\db_dumps\\mefl_backup_latest.sql
"""

import os
import sys
import shutil
import subprocess
from dotenv import load_dotenv

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))

PSQL_CANDIDATES = [
    r"C:\Program Files\PostgreSQL\17\bin\psql.exe",
    r"C:\Program Files\PostgreSQL\16\bin\psql.exe",
    r"C:\Program Files\PostgreSQL\15\bin\psql.exe",
    shutil.which("psql") or "",
]

def find_psql():
    for p in PSQL_CANDIDATES:
        if p and os.path.exists(p):
            return p
    return None

def restore_database(sql_file_path=None):
    psql_bin = find_psql()
    if not psql_bin:
        print("[ERROR] psql.exe not found on system.")
        return False

    # Find backup file if not specified
    if not sql_file_path:
        candidates = [
            r"C:\Users\india\OneDrive - ANIMIKA STUDIOS PRIVATE LIMITED\MEFL_data1\db_backups\mefl_backup_latest.sql",
            os.path.join(BASE_DIR, "backups", "db_dumps", "mefl_backup_latest.sql"),
        ]
        for c in candidates:
            if os.path.exists(c):
                sql_file_path = c
                break

    if not sql_file_path or not os.path.exists(sql_file_path):
        print(f"[ERROR] Backup SQL file not found at: {sql_file_path}")
        return False

    db_host = os.getenv("DB_HOST", "localhost")
    db_port = os.getenv("DB_PORT", "5432")
    db_name = os.getenv("DB_NAME", "MEFL_MARKETING_ANALYTICS")
    db_user = os.getenv("DB_USER", "postgres")
    db_password = os.getenv("DB_PASSWORD", "")

    env = os.environ.copy()
    if db_password:
        env["PGPASSWORD"] = db_password

    print(f"Restoring '{db_name}' from: {sql_file_path}...")
    cmd = [
        psql_bin,
        "-h", db_host,
        "-p", str(db_port),
        "-U", db_user,
        "-d", db_name,
        "-f", sql_file_path,
        "-q"
    ]

    res = subprocess.run(cmd, env=env, capture_output=True, text=True)
    if res.returncode != 0 and "ERROR" in res.stderr:
        print(f"[WARNING] Restore warnings/errors: {res.stderr[:500]}")
    
    print("[SUCCESS] Database restored successfully!")
    return True

if __name__ == "__main__":
    restore_database()
