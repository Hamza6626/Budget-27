# Budget Portal (Free, No VBA)

This portal replaces VBA-based sheet hiding with password-based web access.

## What it does

- Department login by password
- Only that department can enter/edit its budget
- Master login sees detailed consolidated summary (monthly + FY)
- Department and master can download Excel exports
- Data is stored in local SQLite (`budget_data.db`)

## Files

- `budget_portal.py`
- `requirements.txt`
- `.streamlit/secrets.toml.example`
- `DepartmentPasswords_CONFIDENTIAL.csv` (optional local fallback only)

## Run locally (free)

```powershell
pip install -r requirements.txt
streamlit run budget_portal.py
```

## Access model

- First priority: Streamlit Secrets (`MASTER_PASSWORD` and `[DEPARTMENT_PASSWORDS]`)
- Fallback (local only): `DepartmentPasswords_CONFIDENTIAL.csv`

Optional:

- Marketing sheets can be opened via OneDrive Excel Online links using `[MKT_SHEETS_LINKS]`.
- Production is a combined login (domain: "Alaoudin / Usman BK") that uses a single `PRODUCTION` password in `[DEPARTMENT_PASSWORDS]` and can show a constant `PRODUCTION_SHEET_LINK`.

## Notes

- This is macro-free and works even when VBA is disabled.
- For external sharing by link, deploy to Streamlit Community Cloud (free) or host on an internal PC/server.
- If deployed publicly, use strong passwords and rotate them periodically.

## Persistent storage on Streamlit Cloud (recommended)

Local SQLite is fine on your own machine, but Streamlit Cloud redeploys/restarts can lose local file changes.
Use Supabase free tier for durable storage.

1. Create a free Supabase project.
2. In Supabase SQL editor, run:

```sql
create table if not exists public.budget_entries (
	department text primary key,
	payload_json jsonb not null,
	updated_at text not null
);
```

3. In Streamlit app settings -> Secrets, add values from `.streamlit/secrets.toml.example`:

- `MASTER_PASSWORD`
- `[DEPARTMENT_PASSWORDS]` map for all departments
- `[MKT_SHEETS_LINKS]` map (optional)
- `PRODUCTION_SHEET_LINK` (optional)
- `SUPABASE_URL`
- `SUPABASE_KEY`

4. Redeploy/restart app.

When `SUPABASE_URL` and `SUPABASE_KEY` are present, the app uses Supabase and data stays available across restarts and Git pushes.

Note: If you use a publishable/anon key, make sure the `budget_entries` table permissions are set correctly (RLS disabled for that table, or RLS policies added to allow read/write). Otherwise the app will show a Supabase read/write error.
