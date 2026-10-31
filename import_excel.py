"""One-time: copy rows from an existing payments.xlsx into the configured storage
(Google Sheets when STORAGE_BACKEND = "sheets"). Rows already present (same Receipt No) are skipped.

    python import_excel.py payments.xlsx
"""
import sys
from pathlib import Path

import store


def main(path: str):
    src = store.ExcelBackend(Path(path))
    rows = src.read(store.SHEET)
    if len(rows) < 2:
        raise SystemExit("Nothing to import.")
    headers, body = rows[0], rows[1:]
    dest = store.get_backend()
    store.ensure_schema(dest)
    existing = set(store.load_payments()["Receipt No"].str.strip())
    added = skipped = 0
    for row in body:
        rec = {h: (row[i] if i < len(row) else "") for i, h in enumerate(headers) if h}
        no = str(rec.get("Receipt No", "")).strip()
        if not no or no in existing:
            skipped += 1
            continue
        rec.setdefault("Source", store.SOURCE_ADMIN)
        rec["Source"] = rec["Source"] or store.SOURCE_ADMIN
        store.save_payment(rec)
        existing.add(no)
        added += 1
    print(f"Imported {added} row(s), skipped {skipped}.")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit("usage: python import_excel.py payments.xlsx")
    main(sys.argv[1])
