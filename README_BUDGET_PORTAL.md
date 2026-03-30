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
- `requirements_budget_portal.txt`
- `DepartmentPasswords_CONFIDENTIAL.csv`

## Run locally (free)

```powershell
pip install -r requirements_budget_portal.txt
streamlit run budget_portal.py
```

## Access model

- Department passwords are loaded from `DepartmentPasswords_CONFIDENTIAL.csv`
- Master password is the `[MASTER]` row in that same CSV

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

3. In Streamlit app settings (Secrets/Environment Variables), add:

- `SUPABASE_URL` = your Supabase project URL (for example `https://xxxx.supabase.co`)
- `SUPABASE_KEY` = your Supabase anon key (or service role key)

4. Redeploy/restart app.

When both variables are present, the app automatically uses Supabase and data stays available across restarts and Git pushes.
