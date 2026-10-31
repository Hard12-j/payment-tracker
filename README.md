# Gayatri Nagar – Payments (merged)

One Streamlit app. Data lives in **Google Sheets** (or a local Excel file for offline use).

| Who | URL | What they see |
|---|---|---|
| **Public** | `http://<host>:8501/` | **Only the online-payment form.** |
| **Admin** | `http://<host>:8501/?admin` | Password login, then: **New payment** (cash etc. + PDF receipt/email/WhatsApp/Drive), **History / Resend**, **Bank check**, **Rejected attempts**. |

## The data (tab `Payments`)
`Receipt No, Date, Name, Plot No, Amount, Payment Mode, Email, Phone, Remarks, Receipt Link`
+ `Source, Bank Status, Screenshot Hash`

* **Online (public form):** `Receipt No` = the **transaction ID**; `Receipt Link` = the payer's
  **screenshot in Google Drive** (folder "Payment Screenshots"); `Source` = Online.
  Nothing is stored on the server's disk.
* **Admin entries:** `Receipt No` = `RCP-0001, RCP-0002 ...`; `Receipt Link` = Drive link of the PDF receipt.
* Tab `Rejected Attempts` – public submissions that failed a check (admin only).
* Screenshots are **private** (only your Google account can open them). If other people must
  open the links from the sheet, set `SCREENSHOT_LINK_PUBLIC = true`.
* If Drive is down when someone pays, the payment is **still recorded**, with a note in
  Remarks, so the payer is never turned away.

## Setup – local PC
```
pip install -r requirements.txt
copy .streamlit\secrets.toml.example .streamlit\secrets.toml   (cp on Mac/Linux) -> edit it
streamlit run app.py
```
Use `STORAGE_BACKEND = "excel"` to keep it all in a local `payments.xlsx` (no Google needed for
the data; Drive is still used for screenshots/receipts unless `ENABLE_DRIVE_UPLOAD = false`).

## Setup – Google Sheets + Drive (needed for deployment)
1. In your existing Google Cloud project also **enable "Google Sheets API"** (Drive API is already on).
2. *OAuth consent screen* -> set **Publishing status = "In production"**. (In "Testing" Google
   expires the login every 7 days and the deployed app would stop saving.)
3. Delete the old `token.json` and run once on your PC: `python authorize_google.py`
   (log in and accept both Drive and Sheets permissions).
4. **Deployed app:** paste the `GOOGLE_TOKEN_JSON = '''...'''` line it prints into your host's
   Secrets (with the rest of `secrets.toml`). `credentials.json` is not needed on the server.
5. First run creates a spreadsheet **"Gayatri Nagar Payments"** in your Drive (the admin page shows
   a link to it). Optionally put its ID in `SHEETS_SPREADSHEET_ID`.
6. Move your old rows once: `python import_excel.py payments.xlsx` (safe to repeat – duplicates skipped).

## What the public form checks
1. Amount entered appears on the screenshot
2. Transaction ID typed appears on the screenshot (OCR-tolerant)
3. Transaction ID not used before (also checked against every Receipt No)
4. Same screenshot not submitted before
5. Optional: payment went to your account (`RECEIVER_HINTS`; skipped if empty)

**No check is made on status words** ("successful", "failed", "pending", "processing"...).

## Bank check (admin)
Upload your bank statement (CSV/XLSX). Online payments whose transaction ID **and** amount appear
on the same statement row get `Bank Status = Confirmed`. Screenshots can be faked; the bank
statement cannot – run this regularly.

## Notes
* Run the app as **one process** (Streamlit Community Cloud / one Render instance is fine).
* Admin: "Download as Excel" exports both tabs. Google Sheets also keeps version history.
* Local Excel mode only: close `payments.xlsx` in Excel while people pay.
* Tests: `python test_merged.py`, `python test_drive.py`.
