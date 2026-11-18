# Mother Earth Farmlands (MEFL) Marketing Data Synchronization Guide

This document explains how to manually run, test, and manage the Python synchronization script for the MEFL marketing analytics pipeline.

---

## 1. System Overview

- **Source Datasets:**
  - **Campaigns:** Synced Excel `MEFL_Analysis.xlsx` (Meta ads) -> `C:\MEFL_data1\MEFL_Analysis1.csv`
  - **Voice Leads:** `C:\MEFL_data1\mefl_voice_clean.csv`
  - **Results Leads:** `C:\MEFL_data1\mefl_results_clean.csv`
  - **Website Leads:** `C:\MEFL_data1\website_leads_cleaned.csv`
- **Destination:** PostgreSQL 17
  - **Database:** `MEFL_MARKETING_ANALYTICS`
  - **Schema:** `public`
  - **Target Tables:**
    - `campaign_data_ready` (42 records)
    - `mefl_voice` (69 records)
    - `mefl_results` (737 records)
    - `website_leads` (381 records)
- **Script Path:** `C:\MEFL_data1\sync_campaign_data.py`
- **Environment & Credentials:** Loaded securely from `C:\MEFL_data1\.env`
- **Log Output:** `C:\MEFL_data1\logs\sync.log`
- **Scheduled Task:** `MEFL_Marketing_Sync` (runs every 15 minutes automatically via Windows Task Scheduler)

---

## 2. Synchronization Rules & Architecture

1. **Unique Natural Keys & Deduplication:**
   - **`campaign_data_ready`:** Composite key $(\texttt{campaign\_no},\ \texttt{platform},\ \texttt{start\_date},\ \texttt{end\_date})$.
   - **`mefl_voice` & `mefl_results`:** Primary key `lead` (`BIGINT`).
   - **`website_leads`:** Multiset duplicate-aware synchronization. Detects genuine row instances without assuming duplicates, inserting only excess instances from source and preserving all historical rows.
2. **Upsert Logic:**
   - **New Records:** Inserted if the key does not exist.
   - **Changed Records:** Updated only if non-key columns have changed.
   - **Unchanged Records:** Skipped with zero overhead.
   - **Omitted Records:** **Never deleted**. Historical rows in PostgreSQL that are not in the CSVs remain intact.
3. **Data Type Handling:**
   - Dates formatted as `MM/DD/YYYY` or `YYYY-MM-DD` are cleanly parsed.
   - Timestamps with hours/minutes/seconds are parsed to `TIMESTAMP`.
   - Blank cells or whitespace-only cells are converted to `NULL`.
   - Numeric and integer fields are cleanly cast.
4. **Transactions & Rollback:**
   - Each dataset executes in an isolated atomic transaction (`BEGIN ... COMMIT`).
   - If an error occurs on any dataset, an automatic `ROLLBACK` is issued for that table without corrupting state.

---

## 3. How to Run the Script Manually

Open **PowerShell** or **Command Prompt** and navigate to `C:\MEFL_data1`:

### A. Dry-Run Mode (Simulation - Safe)
Inspect proposed changes across all 4 datasets without committing to the database:
```powershell
python C:\MEFL_data1\sync_campaign_data.py --dry-run
```

### B. Live Synchronization Mode (Apply Changes to All 4 Datasets)
Commits changes to PostgreSQL for campaigns and all lead datasets:
```powershell
python C:\MEFL_data1\sync_campaign_data.py
```

### C. Selective Execution Options
```powershell
# Only sync campaigns
python C:\MEFL_data1\sync_campaign_data.py --only-campaigns

# Only sync leads (voice, results, website)
python C:\MEFL_data1\sync_campaign_data.py --only-leads
```

---

## 4. Viewing Synchronization Logs

Every run records timestamps, row counts, proposed/executed changes, errors, and execution duration in:
```
C:\MEFL_data1\logs\sync.log
```

To view the latest log entries in PowerShell:
```powershell
Get-Content C:\MEFL_data1\logs\sync.log -Tail 30
```

---

## 5. Verification SQL Queries

Run these in `psql` or pgAdmin on database `MEFL_MARKETING_ANALYTICS` to verify results:

```sql
-- 1. Check total rows
SELECT COUNT(*) FROM public.campaign_data_ready;

-- 2. Check latest updated/inserted campaigns
SELECT campaign_no, platform, campaign_name, start_date, end_date, amount, leads 
FROM public.campaign_data_ready 
ORDER BY start_date DESC;

-- 3. Verify zero key duplicates
SELECT campaign_no, platform, start_date, end_date, COUNT(*) 
FROM public.campaign_data_ready
GROUP BY campaign_no, platform, start_date, end_date
HAVING COUNT(*) > 1;
```
