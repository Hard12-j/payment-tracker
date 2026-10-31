"""ADMIN page (password protected): cash entries + receipts, history/resend,
bank check, rejected attempts. Opened at  http://<host>:8501/?admin"""
import hmac
import re
import time
from datetime import datetime
from decimal import Decimal
from pathlib import Path

import pandas as pd
import streamlit as st

import bank_check
import receipts
import store
from drive_uploader import upload_receipt
from settings import ENABLE_DRIVE_UPLOAD, get_setting


# --------------------------------------------------------------------------
# Login (fails closed: no ADMIN_PASSWORD configured = admin area disabled)
# --------------------------------------------------------------------------
def _logged_in() -> bool:
    password = get_setting("ADMIN_PASSWORD")
    if not password:
        st.error("Admin area is disabled. Set ADMIN_PASSWORD in .streamlit/secrets.toml.")
        return False
    if st.session_state.get("is_admin"):
        return True
    st.title("🔒 Admin login")
    with st.form("login"):
        typed = st.text_input("Password", type="password")
        ok = st.form_submit_button("Log in")
    if ok:
        if hmac.compare_digest(typed.encode(), password.encode()):
            st.session_state["is_admin"] = True
            st.rerun()
        else:
            time.sleep(1)  # slow down password guessing
            st.error("Wrong password.")
    return False


def render():
    if not _logged_in():
        return

    head_l, head_r = st.columns([6, 1])
    head_l.title("🧾 Payment Receipts — Admin")
    if head_r.button("Log out"):
        st.session_state.clear()
        st.rerun()

    sheet_url = store.storage_url()
    if sheet_url:
        st.caption(f"Data is stored in Google Sheets: [open the sheet]({sheet_url})")
    else:
        st.caption("Data is stored in the local file payments.xlsx")

    tab_new, tab_history, tab_bank, tab_rej = st.tabs(
        ["New payment", "History / Resend", "Bank check", "Rejected attempts"])
    with tab_new:
        _tab_new_payment()
    with tab_history:
        _tab_history()
    with tab_bank:
        _tab_bank_check()
    with tab_rej:
        _tab_rejected()


# --------------------------------------------------------------------------
# Shared widget: download / WhatsApp / email / drive for one receipt
# --------------------------------------------------------------------------
def show_result(rec: dict, pdf_path: Path, key_prefix: str):
    st.download_button("⬇️ Download PDF receipt", pdf_path.read_bytes(),
                       file_name=pdf_path.name, mime="application/pdf",
                       key=f"{key_prefix}_dl")

    link = str(rec.get("Receipt Link") or "").strip()
    online = rec.get("Source") == store.SOURCE_ONLINE
    if online:
        # For online payments the Receipt Link is the payer's SCREENSHOT (kept in Drive).
        if link:
            st.markdown(f"📎 **Payment screenshot (Drive):** [open]({link})")
        msg_rec = dict(rec, **{"Receipt Link": ""})      # never send the screenshot link to payers
        wa_label = "💬 Send via WhatsApp (opens chat, attach the PDF)"
    else:
        msg_rec = rec
        if link:
            st.markdown(f"🔗 **Drive Receipt Link:** [{link}]({link})")
            wa_label = "💬 Send via WhatsApp (includes Drive link)"
        else:
            wa_label = "💬 Send via WhatsApp (opens chat, attach the PDF)"
            if st.button("☁️ Upload to Drive and get link", key=f"{key_prefix}_up_drive"):
                try:
                    new_link = upload_receipt(pdf_path)
                    store.update_receipt_link(rec["Receipt No"], new_link)
                    rec["Receipt Link"] = new_link
                    if "last" in st.session_state and st.session_state["last"][0].get("Receipt No") == rec["Receipt No"]:
                        st.session_state["last"] = (rec, str(pdf_path))
                    st.success("Uploaded to Google Drive and saved!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Google Drive upload failed: {e}")

    if receipts.clean_phone(rec.get("Phone", "")):
        st.link_button(wa_label, receipts.whatsapp_link(msg_rec))
    if rec.get("Email") and st.button("📧 Send email again", key=f"{key_prefix}_mail"):
        try:
            receipts.send_email(msg_rec, pdf_path)
            st.success(f"Email sent to {rec['Email']}")
        except Exception as e:
            st.error(f"Email failed: {e}")


# --------------------------------------------------------------------------
# Tab 1: new (cash) payment
# --------------------------------------------------------------------------
def _tab_new_payment():
    receipt_no = store.next_receipt_no()
    st.caption(f"Next receipt no: **{receipt_no}**")

    with st.form("payment_form", clear_on_submit=False):
        name = st.text_input("Name *")
        plot_no = st.text_input("Plot no. *")
        amount = st.number_input("Amount (Rs.) *", min_value=0.0, step=100.0, format="%.2f")
        mode = st.selectbox("Payment mode", ["Cash", "UPI", "Cheque", "Bank transfer", "Card"])
        c1, c2 = st.columns(2)
        email = c1.text_input("Email")
        phone = c2.text_input("Phone / WhatsApp (10 digits)")
        remarks = st.text_input("Remarks (optional)")
        auto_email = st.checkbox("Email the receipt automatically", value=True)
        drive_upload = st.checkbox("Upload receipt to Google Drive and include link",
                                   value=ENABLE_DRIVE_UPLOAD)
        submitted = st.form_submit_button("Save & generate receipt", type="primary")

    if submitted:
        errors = []
        if not name.strip():
            errors.append("Name is required.")
        if not plot_no.strip():
            errors.append("Plot no. is required.")
        if amount <= 0:
            errors.append("Amount must be greater than zero.")
        if email and not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email.strip()):
            errors.append("Email address looks invalid.")
        if phone and len(receipts.clean_phone(phone)) < 11:
            errors.append("Phone number looks invalid.")
        if auto_email and not email:
            errors.append("Enter an email address, or untick automatic email.")

        if errors:
            for e in errors:
                st.error(e)
        else:
            rec = {
                "Receipt No": receipt_no,
                "Date": datetime.now().strftime("%d-%m-%Y %H:%M"),
                "Name": name.strip(),
                "Plot No": plot_no.strip(),
                "Amount": str(Decimal(str(amount)).quantize(Decimal("0.01"))),
                "Payment Mode": mode,
                "Email": email.strip(),
                "Phone": phone.strip(),
                "Remarks": remarks.strip(),
                "Receipt Link": "",
                "Source": store.SOURCE_ADMIN,
            }
            pdf_path = receipts.generate_receipt_pdf(rec)

            if drive_upload:
                try:
                    rec["Receipt Link"] = upload_receipt(pdf_path)
                    st.success("Uploaded receipt to Google Drive.")
                except Exception as e:
                    st.warning(f"Google Drive upload failed: {e}")

            try:
                store.add_admin_payment(rec)
            except PermissionError:
                st.error("Could not save. If you use a local Excel file, close it in Excel and try again.")
                st.stop()

            st.session_state["last"] = (rec, str(pdf_path))
            st.success(f"Saved. Receipt {receipt_no} created.")

            if auto_email and rec["Email"]:
                try:
                    receipts.send_email(rec, pdf_path)
                    st.success(f"Email sent to {rec['Email']}")
                except Exception as e:
                    st.warning(f"Payment saved, but email failed: {e}")

    if "last" in st.session_state:
        rec, p = st.session_state["last"]
        st.divider()
        st.subheader(f"Receipt {rec['Receipt No']}")
        show_result(rec, Path(p), "last")


# --------------------------------------------------------------------------
# Tab 2: history (cash + online) / resend
# --------------------------------------------------------------------------
def _tab_history():
    df = store.load_payments()
    if df.empty:
        st.info("No payments recorded yet.")
        return

    c1, c2 = st.columns([3, 1])
    q = c1.text_input("Search by name, plot no. or receipt no. / transaction ID")
    src = c2.selectbox("Source", ["All", store.SOURCE_ADMIN, store.SOURCE_ONLINE])
    view = df
    if src != "All":
        view = view[view["Source"] == src]
    if q:
        mask = view[["Receipt No", "Name", "Plot No"]].apply(
            lambda col: col.str.contains(q, case=False, na=False, regex=False)).any(axis=1)
        view = view[mask]
    shown = view.drop(columns=["Screenshot Hash"], errors="ignore")
    st.dataframe(shown.iloc[::-1], use_container_width=True, hide_index=True)

    st.download_button(
        "⬇️ Download as Excel", store.export_xlsx_bytes(), file_name="payments.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.divider()
    st.subheader("Open a record / resend receipt")
    options = view["Receipt No"].tolist()[::-1] or ["-"]
    choice = st.selectbox("Receipt No", options)
    if choice != "-":
        rec = df[df["Receipt No"] == choice].iloc[0].to_dict()
        pdf_path = receipts.receipt_path(choice)
        if not pdf_path.exists():
            pdf_path = receipts.generate_receipt_pdf(rec)
        show_result(rec, pdf_path, f"hist_{choice}")


# --------------------------------------------------------------------------
# Tab 3: bank check
# --------------------------------------------------------------------------
def _tab_bank_check():
    st.write("Upload your bank statement (CSV/Excel). Online payments whose **transaction ID "
             "(Receipt No)** and **amount** appear on the same statement row are marked "
             "**Confirmed**. Screenshots can be edited, so this is the check you can fully trust.")
    stmt = st.file_uploader("Bank statement", type=["csv", "xlsx"], key="stmt")
    if stmt and st.button("Match against statement"):
        try:
            sdf = bank_check.read_statement(stmt)
        except Exception as e:
            st.error(f"Could not read that statement: {e}")
            return
        df = store.load_payments()
        online = df[df["Source"] == store.SOURCE_ONLINE].to_dict("records")
        confirmed = bank_check.match_statement(sdf, online)
        try:
            changed = store.mark_bank_confirmed(confirmed)
        except PermissionError:
            st.error("Could not save. If you use a local Excel file, close it in Excel and try again.")
            return
        st.success(f"{len(confirmed)} online payment(s) found on the statement "
                   f"({changed} newly marked Confirmed).")

    df = store.load_payments()
    online = df[df["Source"] == store.SOURCE_ONLINE]
    pending = online[online["Bank Status"] != "Confirmed"]
    st.subheader(f"Online payments not yet confirmed by the bank ({len(pending)})")
    if pending.empty:
        st.info("Nothing pending." if len(online) else "No online payments yet.")
    else:
        st.dataframe(pending[["Receipt No", "Date", "Name", "Plot No", "Amount", "Bank Status"]],
                     use_container_width=True, hide_index=True)


# --------------------------------------------------------------------------
# Tab 4: rejected attempts (audit log)
# --------------------------------------------------------------------------
def _tab_rejected():
    att = store.load_attempts()
    if att.empty:
        st.info("No rejected attempts.")
    else:
        st.dataframe(att.iloc[::-1], use_container_width=True, hide_index=True)
