"""
Safe Reconciliation Script for MEFL Vedio Ads records.
Validates each exact target row before updating platform in public.campaign_data_ready.
Supports dry-run and live modes with strict rollback.
"""

import os
import sys
import psycopg2
from dotenv import load_dotenv

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))

RECONCILIATION_TARGETS = [
    {
        "campaign_no": "VE_C01",
        "current_platform": "Vedio Ads",
        "new_platform": "Instagram",
        "start_date": "2026-10-05",
        "end_date": "2026-10-11",
    },
    {
        "campaign_no": "VE_C02",
        "current_platform": "Vedio Ads",
        "new_platform": "Facebook",
        "start_date": "2026-10-05",
        "end_date": "2026-10-11",
    },
    {
        "campaign_no": "VE_C03",
        "current_platform": "Vedio Ads",
        "new_platform": "Instagram",
        "start_date": "2026-10-12",
        "end_date": "2026-10-18",
    },
    {
        "campaign_no": "VE_C04",
        "current_platform": "Vedio Ads",
        "new_platform": "Facebook",
        "start_date": "2026-10-12",
        "end_date": "2026-10-18",
    },
]


def run_reconciliation(dry_run=True):
    mode = "[DRY-RUN]" if dry_run else "[LIVE]"
    print(f"\n{'='*60}")
    print(f"Starting Legacy Platform Reconciliation {mode}")
    print(f"Target: MEFL_MARKETING_ANALYTICS.public.campaign_data_ready")

    conn = psycopg2.connect(
        host=os.getenv("DB_HOST", "localhost"),
        port=os.getenv("DB_PORT", "5432"),
        dbname=os.getenv("DB_NAME", "MEFL_MARKETING_ANALYTICS"),
        user=os.getenv("DB_USER", "postgres"),
        password=os.getenv("DB_PASSWORD"),
        connect_timeout=10,
    )
    conn.autocommit = False
    cur = conn.cursor()

    try:
        # Pre-validation 1: Verify current total row count is exactly 43
        cur.execute("SELECT COUNT(*) FROM public.campaign_data_ready;")
        initial_count = cur.fetchone()[0]
        if initial_count != 43:
            raise ValueError(
                f"Safety check failed: Expected exactly 43 rows, found {initial_count}."
            )
        print(f"Initial total rows verified: {initial_count}")

        # Pre-validation 2: For each target, verify exactly 1 row matches current_platform,
        # AND 0 rows match new_platform (ensures no key collision)
        for t in RECONCILIATION_TARGETS:
            c_no = t["campaign_no"]
            c_plat = t["current_platform"]
            n_plat = t["new_platform"]
            s_dt = t["start_date"]
            e_dt = t["end_date"]

            # Check existing target row
            cur.execute(
                """
                SELECT COUNT(*) FROM public.campaign_data_ready
                WHERE campaign_no = %s AND platform = %s AND start_date = %s AND end_date = %s;
            """,
                (c_no, c_plat, s_dt, e_dt),
            )
            count_current = cur.fetchone()[0]
            if count_current != 1:
                raise ValueError(
                    f"Validation failed for ({c_no}, {c_plat}, {s_dt}): Expected 1 row, found {count_current}."
                )

            # Check potential collision with new_platform
            cur.execute(
                """
                SELECT COUNT(*) FROM public.campaign_data_ready
                WHERE campaign_no = %s AND platform = %s AND start_date = %s AND end_date = %s;
            """,
                (n_plat, c_no, s_dt, e_dt),
            )
            count_new = cur.fetchone()[0]
            if count_new != 0:
                raise ValueError(
                    f"Collision detected! Key ({c_no}, {n_plat}, {s_dt}) already exists with {count_new} rows."
                )

            print(
                f"  Verified target: ({c_no}, {s_dt} to {e_dt}) -> Found 1 '{c_plat}' row, 0 '{n_plat}' rows."
            )

        # Execute Updates
        print("\nExecuting in-place key updates:")
        for t in RECONCILIATION_TARGETS:
            cur.execute(
                """
                UPDATE public.campaign_data_ready
                SET platform = %s
                WHERE campaign_no = %s AND platform = %s AND start_date = %s AND end_date = %s;
            """,
                (
                    t["new_platform"],
                    t["campaign_no"],
                    t["current_platform"],
                    t["start_date"],
                    t["end_date"],
                ),
            )
            if cur.rowcount != 1:
                raise RuntimeError(
                    f"Expected rowcount 1 on update for {t['campaign_no']}, got {cur.rowcount}."
                )
            print(
                f"  Updated ({t['campaign_no']}, {t['start_date']}) platform: '{t['current_platform']}' -> '{t['new_platform']}'"
            )

        # Post-validation 3: Verify row count is STILL exactly 43
        cur.execute("SELECT COUNT(*) FROM public.campaign_data_ready;")
        final_count = cur.fetchone()[0]
        if final_count != 43:
            raise ValueError(
                f"Safety check failed: Final row count is {final_count}, expected 43."
            )

        # Post-validation 4: Verify zero duplicate keys exist
        cur.execute(
            """
            SELECT campaign_no, platform, start_date, end_date, COUNT(*)
            FROM public.campaign_data_ready
            GROUP BY campaign_no, platform, start_date, end_date
            HAVING COUNT(*) > 1;
        """
        )
        dupes = cur.fetchall()
        if dupes:
            raise ValueError(
                f"Safety check failed: Found {len(dupes)} duplicate keys after update!"
            )

        print("\nAll post-update safety assertions PASSED (43 rows, 0 duplicate keys).")

        if dry_run:
            conn.rollback()
            print(
                f"[DRY-RUN COMPLETE] Transaction rolled back. Zero changes committed to DB."
            )
        else:
            conn.commit()
            print(
                f"[LIVE COMPLETE] Transaction committed successfully. 4 records updated."
            )

    except Exception as e:
        conn.rollback()
        print(f"[ERROR] Transaction rolled back due to error: {e}")
        raise
    finally:
        cur.close()
        conn.close()
    print(f"{'='*60}\n")


if __name__ == "__main__":
    is_live = len(sys.argv) > 1 and sys.argv[1] == "--live"
    run_reconciliation(dry_run=not is_live)
