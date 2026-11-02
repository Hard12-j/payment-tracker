import io

import pandas as pd

import verify


class WrongPassword(ValueError):
    """The statement is locked and the password was missing or wrong."""


def _bytes_of(uploaded_file) -> bytes:
    return uploaded_file.getvalue() if hasattr(uploaded_file, "getvalue") else uploaded_file.read()


def _read_excel(data: bytes, name: str, password: str) -> pd.DataFrame:
    import msoffcrypto
    buf = io.BytesIO(data)
    try:
        office = msoffcrypto.OfficeFile(buf)
        locked = office.is_encrypted()
    except Exception:
        locked = False                      # not an Office container (plain xlsx is a zip) -> read as is
        buf.seek(0)
    if locked:
        if not password:
            raise WrongPassword("This Excel statement is password protected. Enter its password.")
        try:
            office.load_key(password=password)
            out = io.BytesIO()
            office.decrypt(out)
        except Exception:
            raise WrongPassword("Wrong password for this Excel statement.")
        out.seek(0)
        buf = out
    return pd.read_excel(buf, dtype=str)


def _read_pdf(data: bytes, password: str) -> pd.DataFrame:
    """Every table row / text line becomes one 'row' of text. Wrapped lines are also joined with
    the next line, because the transaction ID and the amount are sometimes on neighbouring lines.
    Scanned (image-only) PDFs cannot be read."""
    import pdfplumber
    try:
        pdf = pdfplumber.open(io.BytesIO(data), password=password or None)
    except Exception as e:
        chain = [e, *getattr(e, "args", ()), getattr(e, "__context__", None), getattr(e, "__cause__", None)]
        text = " ".join(f"{type(x).__name__} {x}" for x in chain if x is not None).lower()
        if "password" in text or "encrypt" in text:
            raise WrongPassword("Wrong password for this PDF statement." if password
                                else "This PDF statement is password protected. Enter its password.")
        raise
    rows = []
    with pdf:
        for page in pdf.pages:
            for table in page.extract_tables() or []:
                for r in table:
                    cells = [str(c).replace("\n", " ").strip() for c in r if c]
                    if cells:
                        rows.append(" ".join(cells))
            lines = [ln.strip() for ln in (page.extract_text() or "").splitlines() if ln.strip()]
            for i, ln in enumerate(lines):
                rows.append(ln)
                if i + 1 < len(lines):
                    rows.append(ln + " " + lines[i + 1])
    if not rows:
        raise ValueError("No text found in this PDF (it may be a scanned image). "
                         "Download the Excel/CSV version from your bank instead.")
    return pd.DataFrame({"row": rows})


def read_statement(uploaded_file, password: str = "") -> pd.DataFrame:
    """CSV, Excel (.xlsx/.xls) or PDF. `password` is only used if the file is locked and is
    never stored anywhere."""
    name = uploaded_file.name.lower()
    data = _bytes_of(uploaded_file)
    if name.endswith(".csv"):
        return pd.read_csv(io.BytesIO(data), dtype=str, encoding_errors="replace")
    if name.endswith(".pdf"):
        return _read_pdf(data, password)
    return _read_excel(data, name, password)


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
