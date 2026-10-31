"""Pure verification logic for the public payment form (no UI, no OCR, no storage).

NOTE: there is deliberately NO check on status words ("successful", "failed", "pending",
"processing"...). A submission is never accepted or rejected because of such words.
"""
import hashlib
import os
import re

# Characters OCR commonly confuses. Applied to BOTH sides before comparing IDs.
_LOOKALIKE = str.maketrans({"O": "0", "Q": "0", "I": "1", "L": "1", "S": "5", "Z": "2", "B": "8"})


def normalize_id(txn_id: str) -> str:
    """Uppercase, drop non-alphanumerics, fold look-alike characters."""
    return re.sub(r"[^A-Z0-9]", "", str(txn_id).upper()).translate(_LOOKALIKE)


def clean_txn_id(txn_id: str) -> str:
    """The transaction ID as it is stored as the Receipt No (whitespace removed)."""
    return re.sub(r"\s+", "", str(txn_id))


def image_hash(image_bytes: bytes) -> str:
    return hashlib.sha256(image_bytes).hexdigest()


def amounts_in_text(text: str):
    """All numeric amounts found in the OCR text, as floats."""
    found = set()
    for m in re.finditer(r"(?<![\d.])(\d{1,3}(?:,\d{2,3})*|\d+)(?:\.(\d{1,2}))?(?![\d])", text):
        whole = m.group(1).replace(",", "")
        frac = m.group(2) or "0"
        try:
            found.add(float(f"{whole}.{frac}"))
        except ValueError:
            pass
    return found


def check_amount(text: str, expected: float) -> bool:
    return any(abs(a - float(expected)) < 0.005 for a in amounts_in_text(text))


def check_txn_in_text(text: str, typed_id: str) -> bool:
    return normalize_id(typed_id) in normalize_id(text)


def check_receiver(text: str, hints_csv: str = ""):
    """Optional. Only active if RECEIVER_HINTS is set (comma-separated UPI IDs / names)."""
    hints = [h.strip() for h in (hints_csv or os.getenv("RECEIVER_HINTS", "")).split(",") if h.strip()]
    if not hints:
        return True, ""
    low = re.sub(r"\s+", "", text.lower())
    if any(re.sub(r"\s+", "", h.lower()) in low for h in hints):
        return True, ""
    return False, "Screenshot does not show the payment going to our account"


def run_checks(text, typed_id, amount, receiver_hints: str = ""):
    """Returns a list of rejection reasons (empty list = passed).
    Duplicate checks (txn id / screenshot) live in store.py because they need the Excel file."""
    errors = []
    if len(normalize_id(typed_id)) < 10:
        errors.append("Transaction ID looks too short")
    if not check_amount(text, amount):
        errors.append("Amount in screenshot does not match the amount entered")
    if not check_txn_in_text(text, typed_id):
        errors.append("Transaction ID not found in the screenshot")
    ok, reason = check_receiver(text, receiver_hints)
    if not ok:
        errors.append(reason)
    return errors
