"""
Gayatri Nagar payments  (merged project)

  Public  :  streamlit run app.py          -> online-payment form ONLY
  Admin   :  same server, open  /?admin    -> password login (ADMIN_PASSWORD)

All data lives in one Excel file, payments.xlsx
(online payments use the transaction ID as the Receipt No).
"""
import streamlit as st

is_admin_url = "admin" in st.query_params

st.set_page_config(
    page_title="Payment Receipts — Admin" if is_admin_url else "Online payment",
    page_icon="🧾",
    layout="wide" if is_admin_url else "centered",
    initial_sidebar_state="collapsed",
)

if is_admin_url:
    import admin_page
    admin_page.render()
else:
    import public_page
    public_page.render()
