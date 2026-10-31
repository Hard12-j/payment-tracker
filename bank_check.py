"""Bank-statement matching (admin only). Pure logic, no UI."""
import pandas as pd

import verify


def read_statement(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file, dtype=str, encoding_errors="replace")
    return pd.read_excel(uploaded_file, dtype=str)


def match_statement(sdf: pd.DataFrame, online_rows: list) -> list:
    """Return the Receipt Nos (transaction IDs) found on the same statement row as a
    matching amount. `online_rows` are dicts with 'Receipt No' and 'Amount'."""
    row_texts = sdf.fillna("").astype(str).agg(" ".join, axis=1).tolist()
    row_keys = [verify.normalize_id(t) for t in row_texts]
    confirmed = []
    for p in online_rows:
        key = verify.normalize_id(p["Receipt No"])
        if len(key) < 6:
            continue
        try:
            amt = float(p["Amount"])
        except (TypeError, ValueError):
            continue
        for rk, txt in zip(row_keys, row_texts):
            if key in rk and verify.check_amount(txt, amt):
                confirmed.append(p["Receipt No"])
                break
    return confirmed
