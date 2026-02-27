import os
from datetime import datetime

import pandas as pd
from sqlalchemy import create_engine, text

import utils


def read_sheet_or_empty(excel_path: str, sheet_name: str, columns: list[str]) -> pd.DataFrame:
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl")
    except Exception:
        df = pd.DataFrame(columns=columns)

    for col in columns:
        if col not in df.columns:
            df[col] = None

    return df[columns]


def migrate_audit_logs(excel_path: str, db_url: str) -> None:
    df_audit = read_sheet_or_empty(excel_path, utils.SHEET_AUDIT, utils.AUDIT_COLUMNS)
    if df_audit.empty:
        print("Audit log kosong, skip.")
        return

    engine = create_engine(db_url, pool_pre_ping=True)
    rows = []
    for _, row in df_audit.iterrows():
        ts = row.get("Timestamp")
        if pd.isna(ts):
            parsed_ts = datetime.now()
        else:
            parsed_ts = pd.to_datetime(ts, errors="coerce")
            parsed_ts = parsed_ts.to_pydatetime() if not pd.isna(parsed_ts) else datetime.now()

        rows.append(
            {
                "timestamp": parsed_ts,
                "username": row.get("User"),
                "action": row.get("Action"),
                "doc_id": row.get("Doc ID"),
                "details": row.get("Details"),
            }
        )

    with engine.begin() as conn:
        conn.execute(text("DELETE FROM audit_logs;"))
        conn.execute(
            text(
                """
                INSERT INTO audit_logs (timestamp, username, action, doc_id, details)
                VALUES (:timestamp, :username, :action, :doc_id, :details);
                """
            ),
            rows,
        )

    print(f"Audit logs migrated: {len(rows)} rows")


def main() -> None:
    db_url = os.getenv("SUPABASE_DB_URL")
    if not db_url:
        raise RuntimeError("SUPABASE_DB_URL belum diset. Set env var dulu lalu jalankan ulang.")

    excel_path = utils.get_excel_path()
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel file tidak ditemukan: {excel_path}")

    print(f"Source Excel: {excel_path}")

    df_docs = read_sheet_or_empty(excel_path, utils.SHEET_DOCUMENT, utils.DOC_COLUMNS)
    ok_docs, msg_docs = utils.save_documents(df_docs)
    print(msg_docs)
    if not ok_docs:
        raise RuntimeError("Migrasi documents gagal")

    df_users = read_sheet_or_empty(excel_path, utils.SHEET_USERS, utils.USER_COLUMNS)
    ok_users, msg_users = utils.save_users(df_users)
    print(msg_users)
    if not ok_users:
        raise RuntimeError("Migrasi users gagal")

    df_fungsi = read_sheet_or_empty(excel_path, utils.SHEET_FUNGSI, utils.FUNGSI_COLUMNS)
    ok_fungsi, msg_fungsi = utils.save_fungsi(df_fungsi)
    print(msg_fungsi)
    if not ok_fungsi:
        raise RuntimeError("Migrasi fungsi gagal")

    migrate_audit_logs(excel_path, db_url)

    print("\n✅ Migrasi Excel -> Supabase selesai.")


if __name__ == "__main__":
    main()
