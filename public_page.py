"""PUBLIC page: the online-payment form only. Nothing else is shown to the public
(no records, no rejected attempts, no bank check)."""
import mimetypes

import streamlit as st

import ocr_utils
import store
import verify
from settings import ENABLE_DRIVE_UPLOAD, get_setting

SAVE_ERROR = "We could not save your payment right now. Please try again in a minute."


def render():
    st.title("🧾 Online payment")
    st.caption("Pay to our account, then fill this form with your payment screenshot.")

    with st.form("form"):
        name = st.text_input("Payer name")
        plot_no = st.text_input("Plot no.")
        amount = st.number_input("Amount paid (₹)", min_value=1.0, step=100.0, format="%.2f")
        txn_id = st.text_input("Transaction ID / UTR")
        shot = st.file_uploader("Payment screenshot", type=["png", "jpg", "jpeg", "webp"])
        go = st.form_submit_button("Submit")

    if not go:
        return

    if not (name.strip() and plot_no.strip() and txn_id.strip() and shot):
        st.error("All fields and the screenshot are required.")
        return

    data = shot.getvalue()
    try:
        with st.spinner("Reading screenshot..."):
            text = ocr_utils.extract_text(data)
    except Exception:
        st.error("Could not read that image. Please upload a clear, full screenshot (PNG/JPG).")
        return

    img_hash = verify.image_hash(data)
    errors = verify.run_checks(text, txn_id, amount, get_setting("RECEIVER_HINTS"))

    try:
        dup = store.find_duplicate(verify.normalize_id(txn_id), img_hash)
        if dup:
            errors.append(dup)
        if errors:
            store.log_attempt(name, plot_no, amount, txn_id, "; ".join(errors))
            _show_rejection(errors, text)
            return

        link, remark = _save_screenshot(data, shot.name, txn_id, img_hash)
        try:
            rec = store.add_online_payment(name.strip(), plot_no.strip(), amount, txn_id,
                                           img_hash, receipt_link=link, remarks=remark)
        except store.DuplicateError as e:      # two people submitted the same thing at once
            store.log_attempt(name, plot_no, amount, txn_id, str(e))
            _show_rejection([str(e)], text)
            return
    except Exception:
        st.error(SAVE_ERROR)
        return

    st.success("Payment accepted ✅")
    st.write(f"**Receipt No (Transaction ID):** {rec['Receipt No']}")
    st.write(f"**Amount:** ₹{float(rec['Amount']):,.2f}  \n**Plot no.:** {rec['Plot No']}")


def _save_screenshot(data: bytes, original_name: str, txn_id: str, img_hash: str):
    """Upload the screenshot to Google Drive; its link becomes the record's Receipt Link.
    If Drive is unavailable the payment is STILL recorded (the payer already paid) and the
    problem is noted in Remarks for the admin."""
    if not ENABLE_DRIVE_UPLOAD:
        return "", ""
    try:
        import drive_uploader
        ext = "." + original_name.rsplit(".", 1)[-1].lower() if "." in original_name else ".png"
        safe_id = "".join(ch for ch in verify.clean_txn_id(txn_id) if ch.isalnum() or ch in "-_")
        filename = f"{safe_id}_{img_hash[:8]}{ext}"
        mime = mimetypes.guess_type(filename)[0] or "image/png"
        return drive_uploader.upload_screenshot(data, filename, mime), ""
    except Exception:
        return "", "Screenshot could not be saved to Drive"


def _show_rejection(errors, ocr_text):
    st.error("Submission rejected:")
    for e in errors:
        st.write(f"• {e}")
    with st.expander("What the system read from your screenshot"):
        st.text(ocr_text or "(no text found)")
