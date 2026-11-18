"""
Mother Earth Farmlands (MEFL) - Automated Database Backup Script
Target DB: MEFL_MARKETING_ANALYTICS (PostgreSQL 17)
Backs up full schema and data to:
  1. Local directory: C:\\MEFL_data1\\backups\\db_dumps\\
  2. Cloud OneDrive:  C:\\Users\\india\\OneDrive - ANIMIKA STUDIOS PRIVATE LIMITED\\MEFL_data1\\db_backups\\
"""

import os
import sys
import shutil
import datetime
import subprocess
import logging
from dotenv import load_dotenv

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))

logger = logging.getLogger("mefl_backup")
if not logger.handlers:
    logger.setLevel(logging.INFO)
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))
    logger.addHandler(ch)

# Possible paths for pg_dump.exe
PG_DUMP_CANDIDATES = [
    r"C:\Program Files\PostgreSQL\17\bin\pg_dump.exe",
    r"C:\Program Files\PostgreSQL\16\bin\pg_dump.exe",
    r"C:\Program Files\PostgreSQL\15\bin\pg_dump.exe",
    shutil.which("pg_dump") or "",
]

def find_pg_dump():
    for p in PG_DUMP_CANDIDATES:
        if p and os.path.exists(p):
            return p
    return None

def create_database_backup():
    pg_dump_bin = find_pg_dump()
    if not pg_dump_bin:
        logger.error("pg_dump.exe not found on system. Database backup skipped.")
        return None

    db_host = os.getenv("DB_HOST", "localhost")
    db_port = os.getenv("DB_PORT", "5432")
    db_name = os.getenv("DB_NAME", "MEFL_MARKETING_ANALYTICS")
    db_user = os.getenv("DB_USER", "postgres")
    db_password = os.getenv("DB_PASSWORD", "")

    local_backup_dir = os.path.join(BASE_DIR, "backups", "db_dumps")
    onedrive_backup_dir = r"C:\Users\india\OneDrive - ANIMIKA STUDIOS PRIVATE LIMITED\MEFL_data1\db_backups"

    os.makedirs(local_backup_dir, exist_ok=True)
    os.makedirs(onedrive_backup_dir, exist_ok=True)

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    ts_filename = f"mefl_backup_{timestamp}.sql"
    latest_filename = "mefl_backup_latest.sql"

    local_ts_file = os.path.join(local_backup_dir, ts_filename)
    local_latest_file = os.path.join(local_backup_dir, latest_filename)

    env = os.environ.copy()
    if db_password:
        env["PGPASSWORD"] = db_password

    cmd = [
        pg_dump_bin,
        "-h", db_host,
        "-p", str(db_port),
        "-U", db_user,
        "-d", db_name,
        "--clean",
        "--if-exists",
        "-F", "p",
        "-f", local_ts_file,
    ]

    logger.info(f"Creating PostgreSQL backup for '{db_name}'...")
    res = subprocess.run(cmd, env=env, capture_output=True, text=True)
    if res.returncode != 0:
        logger.error(f"pg_dump failed: {res.stderr}")
        return None

    size_bytes = os.path.getsize(local_ts_file)
    logger.info(f"Database backup created: {local_ts_file} ({size_bytes:,} bytes)")

    # Copy to latest
    shutil.copy2(local_ts_file, local_latest_file)

    # Mirror to OneDrive cloud
    try:
        od_ts_file = os.path.join(onedrive_backup_dir, ts_filename)
        od_latest_file = os.path.join(onedrive_backup_dir, latest_filename)
        shutil.copy2(local_ts_file, od_ts_file)
        shutil.copy2(local_ts_file, od_latest_file)
        logger.info(f"Successfully mirrored database backup to cloud OneDrive:")
        logger.info(f"  -> {od_latest_file}")
    except Exception as e:
        logger.warning(f"Could not mirror backup to OneDrive: {e}")

    # Retain only last 10 timestamped backups locally to conserve disk space
    try:
        all_backups = sorted(
            [os.path.join(local_backup_dir, f) for f in os.listdir(local_backup_dir) if f.startswith("mefl_backup_20")],
            key=os.path.getmtime
        )
        if len(all_backups) > 10:
            for old_f in all_backups[:-10]:
                os.remove(old_f)
    except Exception:
        pass

    return local_latest_file

if __name__ == "__main__":
    create_database_backup()
