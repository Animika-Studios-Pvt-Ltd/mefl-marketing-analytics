import os
import sys
import psycopg2
from dotenv import load_dotenv

dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
if not os.path.exists(dotenv_path):
    print("ERROR: .env file not found at C:\\MEFL_data1\\.env")
    sys.exit(1)

load_dotenv(dotenv_path)

db_host = os.getenv('DB_HOST', 'localhost')
db_port = os.getenv('DB_PORT', '5432')
db_name = os.getenv('DB_NAME', 'MEFL_MARKETING_ANALYTICS')
db_user = os.getenv('DB_USER', 'postgres')
db_password = os.getenv('DB_PASSWORD')

print(f"Connecting to database '{db_name}' on {db_host}:{db_port} as user '{db_user}'...")

try:
    conn = psycopg2.connect(
        host=db_host,
        port=db_port,
        dbname=db_name,
        user=db_user,
        password=db_password,
        connect_timeout=10
    )
    conn.set_session(readonly=True, autocommit=True)
    cur = conn.cursor()
    print("SUCCESS: Connection established successfully!\n")

    # 1. Check table existence
    cur.execute("""
        SELECT table_name 
        FROM information_schema.tables 
        WHERE table_schema = 'public' 
          AND table_name IN ('campaign_data', 'campaign_data_ready');
    """)
    found_tables = [r[0] for r in cur.fetchall()]
    print(f"Target/Candidate tables found in public schema: {found_tables}")

    target_table = 'campaign_data' if 'campaign_data' in found_tables else ('campaign_data_ready' if 'campaign_data_ready' in found_tables else None)
    if not target_table:
        print("WARNING: Neither 'campaign_data' nor 'campaign_data_ready' exists in database!")
        sys.exit(0)

    print(f"\n--- Table Columns and Data Types (public.{target_table}) ---")
    cur.execute(f"""
        SELECT 
            column_name, 
            data_type, 
            character_maximum_length,
            is_nullable
        FROM information_schema.columns
        WHERE table_schema = 'public' 
          AND table_name = '{target_table}'
        ORDER BY ordinal_position;
    """)
    columns = cur.fetchall()
    col_names = [col[0] for col in columns]
    for col in columns:
        col_type = col[1]
        if col[2]:
            col_type += f"({col[2]})"
        print(f"  - {col[0]:<25} : {col_type:<20} (Nullable: {col[3]})")

    # 3. Get current row count
    cur.execute(f"SELECT COUNT(*) FROM public.{target_table};")
    row_count = cur.fetchone()[0]
    print(f"\n--- Current Row Count ---")
    print(f"  Total records in public.{target_table}: {row_count}")

    # 4. Check uniqueness of (campaign_no, platform, start_date, end_date)
    required_key_cols = ['campaign_no', 'platform', 'start_date', 'end_date']
    missing_cols = [c for c in required_key_cols if c not in col_names]
    
    print(f"\n--- Composite Key Uniqueness Analysis ---")
    if missing_cols:
        print(f"  Notice: Cannot evaluate candidate key directly. Missing columns in table: {missing_cols}")
    else:
        cur.execute(f"""
            SELECT campaign_no, platform, start_date, end_date, COUNT(*) 
            FROM public.{target_table}
            GROUP BY campaign_no, platform, start_date, end_date
            HAVING COUNT(*) > 1;
        """)
        dupes = cur.fetchall()
        if dupes:
            print(f"  WARNING: Found {len(dupes)} duplicate groups for candidate key (campaign_no, platform, start_date, end_date):")
            for d in dupes:
                print(f"    Key: {d[0]}, {d[1]}, {d[2]}, {d[3]} -> Count: {d[4]}")
        else:
            print("  SUCCESS: The composite key (campaign_no, platform, start_date, end_date) is 100% UNIQUE across all existing records!")

    # 5. Check primary key or unique constraints if any
    cur.execute(f"""
        SELECT conname, pg_get_constraintdef(c.oid)
        FROM pg_constraint c
        JOIN pg_namespace n ON n.oid = c.connamespace
        WHERE n.nspname = 'public'
          AND conrelid = 'public.{target_table}'::regclass;
    """)
    constraints = cur.fetchall()
    print(f"\n--- Existing Constraints ---")
    if constraints:
        for con in constraints:
            print(f"  - {con[0]}: {con[1]}")
    else:
        print(f"  - No formal constraints (PRIMARY KEY / UNIQUE) currently defined on public.{target_table}.")

    cur.close()
    conn.close()

except psycopg2.Error as e:
    print(f"PostgreSQL Error: {e}")
    sys.exit(1)
except Exception as e:
    print(f"Unexpected Error: {e}")
    sys.exit(1)
