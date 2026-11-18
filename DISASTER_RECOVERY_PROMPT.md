# Antigravity Recovery Prompt (MEFL Marketing Analytics)

If your laptop fails, is formatted, or you move to a new PC, follow this simple process:

## Step 1: Minimum Prerequisites on the New PC
1. Log into your **OneDrive** (`OneDrive - ANIMIKA STUDIOS PRIVATE LIMITED`).
2. Install **PostgreSQL** (version 16 or 17).
3. Open **Antigravity** and open the workspace folder (or create `C:\MEFL_data1`).

---

## Step 2: Copy-Paste this Exact Prompt into Antigravity Chat

```text
I am setting up this machine after a laptop replacement/reset. Please restore and verify the MEFL Marketing Analytics environment completely.

Here are the project details and recovery instructions:
1. Target Database: PostgreSQL localhost:5432, Database Name: MEFL_MARKETING_ANALYTICS.
2. Cloud Backup Location: Check for the latest database dump at:
   "C:\Users\india\OneDrive - ANIMIKA STUDIOS PRIVATE LIMITED\MEFL_data1\db_backups\mefl_backup_latest.sql"
   (or search for mefl_backup_latest.sql in OneDrive).
3. Please perform the following automated steps:
   a. Check if PostgreSQL is running and test the connection. If the database MEFL_MARKETING_ANALYTICS does not exist, create it.
   b. Restore all 4 tables from mefl_backup_latest.sql using psql.
   c. Verify that the 4 tables are populated:
      - public.campaign_data_ready (should have 61 campaigns, ~Rs 5,01,388.21 spend, 1,731 leads)
      - public.website_leads (should have ~618+ leads)
      - public.mefl_results (should have ~1,412+ contacted leads)
      - public.mefl_voice (should have ~140 inbound calls)
   d. Check Python environment (verify openpyxl, psycopg2, python-dotenv are installed).
   e. Create/verify .env file with DB credentials.
   f. Verify sync_campaign_data.py is functional.
   g. Confirm the status and give me the green light to open Power BI Desktop (MEFl_2-11.pbix) and click Refresh.
```

---

## Step 3: What Antigravity Will Do Automatically
When you send the prompt above, Antigravity will:
1. Detect PostgreSQL service.
2. Create `MEFL_MARKETING_ANALYTICS` database if it's a fresh PostgreSQL installation.
3. Execute the SQL restore directly from your cloud OneDrive backup.
4. Run validation queries across all 4 tables to verify that the rows and spend totals match down to the rupee.
5. Setup the Python `.env` configuration.
6. Tell you that everything is restored and ready for Power BI!
