import json
import sqlite3
import uuid
import base64
import os
from datetime import datetime
from io import BytesIO
from pathlib import Path
from urllib.parse import quote

import pandas as pd
import requests
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

APP_TITLE = "MGA Budget Portal FY2026-27"
DB_PATH = Path(__file__).with_name("budget_data.db")
PASSWORD_CSV = Path(__file__).with_name("DepartmentPasswords_CONFIDENTIAL.csv")
SUPABASE_URL = os.getenv("SUPABASE_URL", "").strip().rstrip("/")
SUPABASE_KEY = os.getenv("SUPABASE_KEY", "").strip()
USE_SUPABASE = bool(SUPABASE_URL and SUPABASE_KEY)
MONTHS = [
    "Jul-26", "Aug-26", "Sep-26", "Oct-26", "Nov-26", "Dec-26",
    "Jan-27", "Feb-27", "Mar-27", "Apr-27", "May-27", "Jun-27",
]
SECTION_TEMPLATES = [
    ("1. TRAINING & DEVELOPMENT", [("Employee 1", "units_rate"), ("Employee 2", "units_rate"), ("Employee 3", "units_rate")]),
    ("2. SOFTWARE & LICENSES", [("Platform A", "units_rate"), ("Platform B", "units_rate"), ("Platform C", "units_rate")]),
    ("3. IT EQUIPMENT (CAPEX)", [("Laptops", "units_rate"), ("Others", "units_rate")]),
    ("4. STORE CONSUMPTION", [("Pens", "units_rate"), ("Paper (Ream)", "units_rate"), ("Toner/Ink", "units_rate"), ("Other Items", "units_rate")]),
    ("5. ENTERTAINMENT", [("Client Entertainment", "units_rate"), ("Staff Events", "units_rate")]),
    ("6. FOREIGN TRAVEL (USD)", [("Airfare", "units_rate"), ("Food & Lodging", "units_rate"), ("Other Costs", "amount")]),
    ("7. LOCAL TRAVEL (PKR)", [("Multan", "units_rate"), ("Lahore", "units_rate"), ("Islamabad", "units_rate"), ("Karachi", "units_rate"), ("Other", "units_rate")]),
]
SECTION_NAMES = [s for s, _ in SECTION_TEMPLATES]
MAX_ATTACHMENT_BYTES = 5 * 1024 * 1024
ALLOWED_ATTACHMENT_TYPES = {
    "application/pdf",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "image/png",
    "image/jpeg",
}


def normalize_name(name: str) -> str:
    return name.strip().upper()


def month_map(default: float = 0.0) -> dict:
    return {m: float(default) for m in MONTHS}


def create_item(name: str, kind: str = "units_rate", item_id: str | None = None) -> dict:
    return {
        "id": item_id or uuid.uuid4().hex,
        "name": name,
        "kind": kind,
        "units": month_map(0.0),
        "rate": month_map(0.0),
        "amount": month_map(0.0),
        "benefit": month_map(0.0),
        "comment": "",
        "attachments": [],
    }


def attachment_label(att: dict) -> str:
    size_kb = int((_to_float(att.get("size", 0)) + 1023) // 1024)
    return f"{att.get('name', 'document')} ({size_kb} KB)"


def _is_duplicate_attachment(item: dict, uploaded_name: str, uploaded_size: int) -> bool:
    for att in item.get("attachments", []):
        if att.get("name") == uploaded_name and int(_to_float(att.get("size", 0))) == int(uploaded_size):
            return True
    return False


def add_attachment(item: dict, uploaded_file) -> str | None:
    if uploaded_file is None:
        return None

    file_bytes = uploaded_file.getvalue()
    if len(file_bytes) > MAX_ATTACHMENT_BYTES:
        return "File too large. Max allowed size is 5 MB."
    if uploaded_file.type not in ALLOWED_ATTACHMENT_TYPES:
        return "Unsupported file type. Use PDF, XLSX, DOCX, PNG, or JPG."
    if _is_duplicate_attachment(item, uploaded_file.name, len(file_bytes)):
        return None

    item.setdefault("attachments", []).append(
        {
            "id": uuid.uuid4().hex,
            "name": uploaded_file.name,
            "mime": uploaded_file.type,
            "size": len(file_bytes),
            "content_b64": base64.b64encode(file_bytes).decode("ascii"),
            "uploaded_at": datetime.now().isoformat(timespec="seconds"),
        }
    )
    return None


def default_payload() -> dict:
    payload = {"sections": {}, "updated_at": None}
    for section, items in SECTION_TEMPLATES:
        payload["sections"][section] = [create_item(item_name, kind) for item_name, kind in items]
    return payload


def ensure_db() -> None:
    if USE_SUPABASE:
        return

    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS budget_entries (
            department TEXT PRIMARY KEY,
            payload_json TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """
    )
    conn.commit()
    conn.close()


def _supabase_headers() -> dict:
    return {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }


def _supabase_get_payload(department: str):
    dept_encoded = quote(department, safe="")
    url = (
        f"{SUPABASE_URL}/rest/v1/budget_entries"
        f"?select=payload_json&department=eq.{dept_encoded}&limit=1"
    )
    resp = requests.get(url, headers=_supabase_headers(), timeout=20)
    if resp.status_code >= 400:
        raise RuntimeError(f"Supabase read failed: {resp.status_code} {resp.text[:200]}")
    data = resp.json()
    if not data:
        return None
    return data[0].get("payload_json")


def _supabase_upsert_payload(department: str, payload: dict, updated_at: str) -> None:
    url = f"{SUPABASE_URL}/rest/v1/budget_entries"
    body = [{"department": department, "payload_json": payload, "updated_at": updated_at}]
    headers = _supabase_headers()
    headers["Prefer"] = "resolution=merge-duplicates,return=minimal"
    resp = requests.post(url, headers=headers, data=json.dumps(body), timeout=20)
    if resp.status_code >= 400:
        raise RuntimeError(f"Supabase write failed: {resp.status_code} {resp.text[:200]}")


def _load_auth_from_secrets() -> tuple[dict, str] | None:
    try:
        secrets = st.secrets
    except Exception:
        return None

    master_password = str(secrets.get("MASTER_PASSWORD", "")).strip()
    dept_block = secrets.get("DEPARTMENT_PASSWORDS", {})
    try:
        dept_items = dict(dept_block).items()
    except Exception:
        dept_items = []

    dept_passwords = {}
    for dept, pw in dept_items:
        d = normalize_name(str(dept))
        p = str(pw).strip()
        if d and p:
            dept_passwords[d] = p

    if master_password and dept_passwords:
        return dept_passwords, master_password
    return None


def load_auth_map() -> tuple[dict, str]:
    from_secrets = _load_auth_from_secrets()
    if from_secrets:
        return from_secrets

    if not PASSWORD_CSV.exists():
        st.error("No auth source found. Configure Streamlit Secrets or provide DepartmentPasswords_CONFIDENTIAL.csv")
        st.stop()

    df = pd.read_csv(PASSWORD_CSV)
    dept_passwords = {}
    master_password = ""

    for _, row in df.iterrows():
        dept = str(row["Department"]).strip()
        pw = str(row["Password"]).strip()
        if dept == "[MASTER]":
            master_password = pw
        else:
            dept_key = normalize_name(dept)
            dept_passwords[dept_key] = pw

    if not master_password:
        st.error("Master password entry [MASTER] missing in DepartmentPasswords_CONFIDENTIAL.csv")
        st.stop()

    return dept_passwords, master_password


def _to_float(value) -> float:
    try:
        if value is None or value == "":
            return 0.0
        return float(value)
    except Exception:
        return 0.0


def migrate_payload(payload: dict) -> dict:
    if payload.get("sections"):
        for section in SECTION_NAMES:
            payload["sections"].setdefault(section, [])
        for section, items in payload["sections"].items():
            for item in items:
                item.setdefault("id", uuid.uuid4().hex)
                item.setdefault("name", "New Subhead")
                item.setdefault("kind", "units_rate")
                item.setdefault("units", month_map(0.0))
                item.setdefault("rate", month_map(0.0))
                item.setdefault("amount", month_map(0.0))
                item.setdefault("benefit", month_map(0.0))
                item.setdefault("comment", "")
                item.setdefault("attachments", [])
                for m in MONTHS:
                    item["units"][m] = _to_float(item["units"].get(m, 0.0))
                    item["rate"][m] = _to_float(item["rate"].get(m, 0.0))
                    item["amount"][m] = _to_float(item["amount"].get(m, 0.0))
                    item["benefit"][m] = _to_float(item["benefit"].get(m, 0.0))
        payload.setdefault("updated_at", None)
        return payload

    old_rows = payload.get("rows", {})
    old_comments = payload.get("comments", {})
    new_payload = default_payload()

    for section, _ in SECTION_TEMPLATES:
        section_items = new_payload["sections"][section]
        for item in section_items:
            name = item["name"]
            units_key = f"{section}||{name}||units"
            rate_key = f"{section}||{name}||rate"
            amount_key = f"{section}||{name}||other"
            comment_key = f"{section}||{name}||comment"

            if amount_key in old_rows:
                item["kind"] = "amount"
                for m in MONTHS:
                    item["amount"][m] = _to_float(old_rows.get(amount_key, {}).get(m, 0.0))
            else:
                item["kind"] = "units_rate"
                for m in MONTHS:
                    item["units"][m] = _to_float(old_rows.get(units_key, {}).get(m, 0.0))
                    item["rate"][m] = _to_float(old_rows.get(rate_key, {}).get(m, 0.0))

            item["comment"] = str(old_comments.get(comment_key, ""))

    new_payload["updated_at"] = payload.get("updated_at")
    return new_payload


def save_payload(department: str, payload: dict) -> None:
    now = datetime.now().isoformat(timespec="seconds")
    payload["updated_at"] = now

    if USE_SUPABASE:
        _supabase_upsert_payload(department, payload, now)
        return

    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        """
        INSERT INTO budget_entries (department, payload_json, updated_at)
        VALUES (?, ?, ?)
        ON CONFLICT(department)
        DO UPDATE SET payload_json=excluded.payload_json, updated_at=excluded.updated_at
        """,
        (department, json.dumps(payload), now),
    )
    conn.commit()
    conn.close()


def load_payload(department: str) -> dict:
    if USE_SUPABASE:
        raw = _supabase_get_payload(department)
        if not raw:
            return default_payload()
        if isinstance(raw, str):
            raw = json.loads(raw)
        return migrate_payload(raw)

    conn = sqlite3.connect(DB_PATH)
    row = conn.execute(
        "SELECT payload_json FROM budget_entries WHERE department = ?", (department,)
    ).fetchone()
    conn.close()

    if not row:
        return default_payload()

    raw = json.loads(row[0])
    return migrate_payload(raw)


def load_all_payloads(departments: list[str]) -> dict:
    return {dept: load_payload(dept) for dept in departments}


def is_travel_section(section: str) -> bool:
    return "TRAVEL" in section.upper()


def item_cost_by_month(item: dict) -> list[float]:
    if item.get("kind") == "amount":
        return [_to_float(item["amount"].get(m, 0.0)) for m in MONTHS]
    units = [_to_float(item["units"].get(m, 0.0)) for m in MONTHS]
    rate = [_to_float(item["rate"].get(m, 0.0)) for m in MONTHS]
    return [units[i] * rate[i] for i in range(len(MONTHS))]


def item_benefit_by_month(item: dict) -> list[float]:
    return [_to_float(item.get("benefit", {}).get(m, 0.0)) for m in MONTHS]


def item_fy_cost(item: dict) -> float:
    return float(sum(item_cost_by_month(item)))


def item_fy_benefit(item: dict) -> float:
    return float(sum(item_benefit_by_month(item)))


def item_fy_roi(item: dict) -> float | None:
    cost = item_fy_cost(item)
    if cost <= 0:
        return None
    return ((item_fy_benefit(item) - cost) / cost) * 100.0


def section_totals(payload: dict) -> dict:
    totals = {}
    for section in SECTION_NAMES:
        month_totals = {m: 0.0 for m in MONTHS}
        for item in payload.get("sections", {}).get(section, []):
            vals = item_cost_by_month(item)
            for i, month in enumerate(MONTHS):
                month_totals[month] += vals[i]
        totals[section] = {"months": month_totals, "fy": sum(month_totals.values())}
    return totals


def build_summary_dataframe(all_payloads: dict, departments: list[str]) -> pd.DataFrame:
    rows = []
    for section_idx, section in enumerate(SECTION_NAMES, start=1):
        sec_name = section.split(". ", 1)[1] if ". " in section else section
        rows.append({"#": str(section_idx), "Section / Period": sec_name})
        for period in MONTHS + ["FY Total"]:
            row = {"#": "", "Section / Period": f"  {period}"}
            dept_total = 0.0
            for dept in departments:
                sec = section_totals(all_payloads[dept])[section]
                val = sec["fy"] if period == "FY Total" else sec["months"][period]
                row[dept] = round(val, 2)
                dept_total += val
            row["GRAND TOTAL"] = round(dept_total, 2)
            rows.append(row)
        rows.append({"#": "", "Section / Period": ""})

    grand = {"#": "", "Section / Period": "TOTAL - ALL DEPARTMENTS"}
    for dept in departments:
        dept_grand = 0.0
        for section in SECTION_NAMES:
            dept_grand += section_totals(all_payloads[dept])[section]["fy"]
        grand[dept] = round(dept_grand, 2)
    grand["GRAND TOTAL"] = round(sum(grand[d] for d in departments), 2)
    rows.append(grand)

    return pd.DataFrame(rows)


def workbook_bytes(all_payloads: dict, departments: list[str], include_summary: bool) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    if include_summary:
        ws = wb.create_sheet("CONSOLIDATED SUMMARY")
        df_sum = build_summary_dataframe(all_payloads, departments)
        for col_idx, col_name in enumerate(df_sum.columns, start=1):
            c = ws.cell(row=1, column=col_idx, value=col_name)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F3864")
            c.alignment = Alignment(horizontal="center")
        for r_idx, (_, row) in enumerate(df_sum.iterrows(), start=2):
            for c_idx, col_name in enumerate(df_sum.columns, start=1):
                val = row[col_name]
                c = ws.cell(row=r_idx, column=c_idx, value=val)
                if isinstance(val, (int, float)):
                    c.number_format = '#,##0.00;(#,##0.00);"-"'
                if col_name == "Section / Period":
                    c.alignment = Alignment(horizontal="left")

    for dept in departments:
        payload = all_payloads[dept]
        ws = wb.create_sheet(dept[:31])
        headers = ["Section", "Subhead", "Metric"] + MONTHS + ["FY Total", "FY Benefit", "ROI %", "Support Docs", "Comment"]
        for col_idx, h in enumerate(headers, start=1):
            c = ws.cell(row=1, column=col_idx, value=h)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="2E5FAC")
            c.alignment = Alignment(horizontal="center")

        row_idx = 2
        for section in SECTION_NAMES:
            sec_name = section.split(". ", 1)[1] if ". " in section else section
            for item in payload.get("sections", {}).get(section, []):
                comment = item.get("comment", "")
                is_travel = is_travel_section(section)
                fy_benefit = item_fy_benefit(item) if is_travel else 0.0
                fy_roi = item_fy_roi(item) if is_travel else None

                if item.get("kind") == "amount":
                    vals = item_cost_by_month(item)
                    ws.cell(row=row_idx, column=1, value=sec_name)
                    ws.cell(row=row_idx, column=2, value=item.get("name", "Subhead"))
                    ws.cell(row=row_idx, column=3, value="Amount")
                    for i, v in enumerate(vals, start=4):
                        ws.cell(row=row_idx, column=i, value=round(v, 2)).number_format = '#,##0.00;(#,##0.00);"-"'
                    ws.cell(row=row_idx, column=16, value=round(sum(vals), 2)).number_format = '#,##0.00;(#,##0.00);"-"'
                    ws.cell(row=row_idx, column=17, value=round(fy_benefit, 2) if is_travel else "")
                    ws.cell(row=row_idx, column=18, value=round(fy_roi, 2) if fy_roi is not None else "")
                    ws.cell(row=row_idx, column=19, value="; ".join([a.get("name", "") for a in item.get("attachments", [])]))
                    ws.cell(row=row_idx, column=20, value=comment)
                    row_idx += 1
                else:
                    units = [_to_float(item["units"].get(m, 0.0)) for m in MONTHS]
                    rate = [_to_float(item["rate"].get(m, 0.0)) for m in MONTHS]
                    value = [units[i] * rate[i] for i in range(len(MONTHS))]
                    for metric_name, vals in [("Units", units), ("Rate", rate), ("Value", value)]:
                        ws.cell(row=row_idx, column=1, value=sec_name)
                        ws.cell(row=row_idx, column=2, value=item.get("name", "Subhead"))
                        ws.cell(row=row_idx, column=3, value=metric_name)
                        for i, v in enumerate(vals, start=4):
                            ws.cell(row=row_idx, column=i, value=round(v, 2)).number_format = '#,##0.00;(#,##0.00);"-"'
                        ws.cell(row=row_idx, column=16, value=round(sum(vals), 2)).number_format = '#,##0.00;(#,##0.00);"-"'
                        if metric_name == "Value" and is_travel:
                            ws.cell(row=row_idx, column=17, value=round(fy_benefit, 2))
                            ws.cell(row=row_idx, column=18, value=round(fy_roi, 2) if fy_roi is not None else "")
                            ws.cell(row=row_idx, column=19, value="; ".join([a.get("name", "") for a in item.get("attachments", [])]))
                            ws.cell(row=row_idx, column=20, value=comment)
                        row_idx += 1

        for col in [1, 2, 3, 19, 20]:
            ws.column_dimensions[chr(64 + col)].width = 24 if col in [1, 2] else 16
        for col_offset in range(4, 19):
            ws.column_dimensions[chr(64 + col_offset)].width = 12

    output = BytesIO()
    wb.save(output)
    return output.getvalue()


def _editor_to_month_map(df: pd.DataFrame, row_name: str) -> dict:
    return {m: _to_float(df.loc[row_name, m]) for m in MONTHS}


def render_department_form(department: str, payload: dict) -> dict:
    st.subheader(f"Department: {department}")
    st.caption("Users can add/edit subheads and choose calculation type. Travel subheads include ROI.")

    to_delete = []

    for section in SECTION_NAMES:
        with st.expander(section, expanded=False):
            section_items = payload["sections"].setdefault(section, [])

            add_col, _ = st.columns([1, 3])
            with add_col:
                if st.button("Add Subhead", key=f"add_{department}_{section}"):
                    section_items.append(create_item("New Subhead", "units_rate"))
                    st.rerun()

            if not section_items:
                st.info("No subheads in this section. Click Add Subhead.")

            for idx, item in enumerate(section_items):
                item_id = item["id"]

                st.markdown(f"Subhead #{idx + 1}")
                c1, c2, c3 = st.columns([2, 1.2, 0.8])
                item["name"] = c1.text_input(
                    "Subhead Name",
                    value=item.get("name", "New Subhead"),
                    key=f"name_{department}_{item_id}",
                )
                item["kind"] = c2.selectbox(
                    "Field Type",
                    options=["units_rate", "amount"],
                    format_func=lambda x: "Units x Rate" if x == "units_rate" else "Direct Amount",
                    index=0 if item.get("kind") == "units_rate" else 1,
                    key=f"kind_{department}_{item_id}",
                )
                if c3.button("Delete", key=f"del_{department}_{item_id}"):
                    to_delete.append((section, item_id))

                if item["kind"] == "amount":
                    amount_df = pd.DataFrame(
                        [[_to_float(item["amount"].get(m, 0.0)) for m in MONTHS]],
                        index=["Amount"],
                        columns=MONTHS,
                    )
                    edited_amount = st.data_editor(
                        amount_df,
                        key=f"amt_{department}_{item_id}",
                        use_container_width=True,
                        num_rows="fixed",
                    )
                    item["amount"] = _editor_to_month_map(edited_amount, "Amount")
                else:
                    input_df = pd.DataFrame(
                        [
                            [_to_float(item["units"].get(m, 0.0)) for m in MONTHS],
                            [_to_float(item["rate"].get(m, 0.0)) for m in MONTHS],
                        ],
                        index=["Units", "Rate"],
                        columns=MONTHS,
                    )
                    edited_ur = st.data_editor(
                        input_df,
                        key=f"ur_{department}_{item_id}",
                        use_container_width=True,
                        num_rows="fixed",
                    )
                    item["units"] = _editor_to_month_map(edited_ur, "Units")
                    item["rate"] = _editor_to_month_map(edited_ur, "Rate")
                    calc_vals = item_cost_by_month(item)
                    st.dataframe(
                        pd.DataFrame([calc_vals], index=["Value (auto)"], columns=MONTHS),
                        use_container_width=True,
                    )

                item["comment"] = st.text_input(
                    "Comment / Assumption",
                    value=item.get("comment", ""),
                    key=f"comment_{department}_{item_id}",
                )

                if is_travel_section(section):
                    with st.expander("ROI & Supporting Docs", expanded=True):
                        benefit_df = pd.DataFrame(
                            [[_to_float(item["benefit"].get(m, 0.0)) for m in MONTHS]],
                            index=["Expected Return"],
                            columns=MONTHS,
                        )
                        edited_benefit = st.data_editor(
                            benefit_df,
                            key=f"benefit_{department}_{item_id}",
                            use_container_width=True,
                            num_rows="fixed",
                        )
                        item["benefit"] = _editor_to_month_map(edited_benefit, "Expected Return")

                        fy_roi = item_fy_roi(item)
                        st.caption(
                            f"FY Cost: {item_fy_cost(item):,.2f} | FY Return: {item_fy_benefit(item):,.2f} | "
                            f"FY ROI: {(f'{fy_roi:.2f}%' if fy_roi is not None else 'N/A')}"
                        )

                        uploaded = st.file_uploader(
                            "Attach supporting document",
                            type=["pdf", "xlsx", "docx", "png", "jpg", "jpeg"],
                            key=f"up_{department}_{item_id}",
                        )
                        err = add_attachment(item, uploaded)
                        if err:
                            st.warning(err)

                        docs = item.get("attachments", [])
                        if docs:
                            st.caption("Attached documents")
                            for doc in docs:
                                d1, d2 = st.columns([4, 1])
                                d1.download_button(
                                    f"Download {attachment_label(doc)}",
                                    data=base64.b64decode(doc.get("content_b64", "")),
                                    file_name=doc.get("name", "supporting_doc"),
                                    mime=doc.get("mime", "application/octet-stream"),
                                    key=f"dl_{department}_{item_id}_{doc.get('id')}",
                                )
                                if d2.button("Remove", key=f"rm_{department}_{item_id}_{doc.get('id')}"):
                                    item["attachments"] = [a for a in docs if a.get("id") != doc.get("id")]
                                    st.rerun()

                st.divider()

    for section, item_id in to_delete:
        payload["sections"][section] = [it for it in payload["sections"][section] if it["id"] != item_id]
    if to_delete:
        st.rerun()

    return payload


def build_travel_roi_report(payload: dict) -> pd.DataFrame:
    rows = []
    for section in SECTION_NAMES:
        if not is_travel_section(section):
            continue
        sec_label = section.split(". ", 1)[1] if ". " in section else section
        for item in payload.get("sections", {}).get(section, []):
            fy_roi = item_fy_roi(item)
            rows.append(
                {
                    "Section": sec_label,
                    "Subhead": item.get("name", "Subhead"),
                    "FY Cost": round(item_fy_cost(item), 2),
                    "FY Return": round(item_fy_benefit(item), 2),
                    "FY ROI %": round(fy_roi, 2) if fy_roi is not None else None,
                    "Docs": len(item.get("attachments", [])),
                }
            )
    return pd.DataFrame(rows)


def master_documents_rows(all_payloads: dict) -> list[dict]:
    rows = []
    for dept, payload in all_payloads.items():
        for section in SECTION_NAMES:
            if not is_travel_section(section):
                continue
            for item in payload.get("sections", {}).get(section, []):
                for att in item.get("attachments", []):
                    rows.append(
                        {
                            "Department": dept,
                            "Section": section,
                            "Subhead": item.get("name", "Subhead"),
                            "name": att.get("name", "supporting_doc"),
                            "mime": att.get("mime", "application/octet-stream"),
                            "size": int(_to_float(att.get("size", 0))),
                            "content_b64": att.get("content_b64", ""),
                            "id": att.get("id", uuid.uuid4().hex),
                        }
                    )
    return rows


def login_view(auth_map: dict, master_pw: str) -> None:
    st.title(APP_TITLE)
    st.write("Secure budget entry without VBA. Login with your assigned password.")
    if USE_SUPABASE:
        st.caption("Storage backend: Supabase (persistent across restarts/deploys)")
    else:
        st.caption("Storage backend: Local SQLite (not durable on Streamlit Cloud redeploys)")

    departments = sorted(auth_map.keys())
    with st.form("login_form", clear_on_submit=False):
        login_as = st.selectbox("Login As", options=["MASTER"] + departments)
        pw = st.text_input("Password", type="password")
        submit = st.form_submit_button("Login")

    if not submit:
        return

    if login_as == "MASTER" and pw == master_pw:
        st.session_state.authenticated = True
        st.session_state.role = "master"
        st.session_state.department = None
        st.rerun()

    if login_as != "MASTER" and auth_map.get(login_as) == pw:
        st.session_state.authenticated = True
        st.session_state.role = "department"
        st.session_state.department = login_as
        st.rerun()

    st.error("Invalid password")


def app_view(auth_map: dict, master_pw: str) -> None:
    all_departments = sorted(auth_map.keys())

    with st.sidebar:
        st.write(f"Role: {st.session_state.role}")
        if st.session_state.role == "department":
            st.write(f"Department: {st.session_state.department}")
        if st.button("Logout"):
            st.session_state.authenticated = False
            st.session_state.role = None
            st.session_state.department = None
            st.rerun()

    st.title(APP_TITLE)

    if st.session_state.role == "department":
        dept = st.session_state.department
        work_key = f"work_payload_{dept}"
        if work_key not in st.session_state:
            st.session_state[work_key] = load_payload(dept)

        current_payload = st.session_state[work_key]
        current_payload = render_department_form(dept, current_payload)
        st.session_state[work_key] = current_payload

        st.subheader("Travel ROI Snapshot")
        roi_df = build_travel_roi_report(current_payload)
        if not roi_df.empty:
            st.dataframe(roi_df, use_container_width=True)
        else:
            st.info("No travel subheads found yet.")

        col1, col2 = st.columns([1, 1])
        with col1:
            if st.button("Save Department Data", type="primary"):
                save_payload(dept, current_payload)
                st.success("Department data saved.")

        with col2:
            xlsx = workbook_bytes({dept: current_payload}, [dept], include_summary=False)
            st.download_button(
                "Download My Department Excel",
                data=xlsx,
                file_name=f"Budget_{dept.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    if st.session_state.role == "master":
        all_payloads = load_all_payloads(all_departments)

        status_rows = []
        for dept in all_departments:
            upd = all_payloads[dept].get("updated_at")
            status_rows.append({"Department": dept, "Last Updated": upd or "Not submitted"})

        st.subheader("Submission Status")
        st.dataframe(pd.DataFrame(status_rows), use_container_width=True)

        st.subheader("Detailed Consolidated Summary")
        df_sum = build_summary_dataframe(all_payloads, all_departments)
        st.dataframe(df_sum, use_container_width=True, height=600)

        st.subheader("Travel ROI Summary")
        travel_rows = []
        for dept in all_departments:
            roi_df = build_travel_roi_report(all_payloads[dept])
            if roi_df.empty:
                continue
            roi_df.insert(0, "Department", dept)
            travel_rows.append(roi_df)
        if travel_rows:
            st.dataframe(pd.concat(travel_rows, ignore_index=True), use_container_width=True)
        else:
            st.info("No travel ROI data submitted yet.")

        st.subheader("Travel Supporting Documents")
        doc_rows = master_documents_rows(all_payloads)
        if doc_rows:
            doc_df = pd.DataFrame(
                [{
                    "Department": r["Department"],
                    "Section": r["Section"],
                    "Subhead": r["Subhead"],
                    "Document": r["name"],
                    "Size (KB)": int((r["size"] + 1023) // 1024),
                } for r in doc_rows]
            )
            st.dataframe(doc_df, use_container_width=True)
            for r in doc_rows:
                st.download_button(
                    f"Download {r['Department']} - {r['Subhead']} - {r['name']}",
                    data=base64.b64decode(r["content_b64"]),
                    file_name=r["name"],
                    mime=r["mime"],
                    key=f"mdoc_{r['id']}",
                )
        else:
            st.info("No supporting documents uploaded yet.")

        xlsx = workbook_bytes(all_payloads, all_departments, include_summary=True)
        st.download_button(
            "Download Master Consolidated Excel",
            data=xlsx,
            file_name="Budget_Master_Consolidated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    ensure_db()
    auth_map, master_pw = load_auth_map()

    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
        st.session_state.role = None
        st.session_state.department = None

    if not st.session_state.authenticated:
        login_view(auth_map, master_pw)
        return

    app_view(auth_map, master_pw)


if __name__ == "__main__":
    main()
