"""
Test Script for Google Drive Uploader & App Features
-----------------------------------------------------
1. Pure Tests (Offline / No credentials needed):
   - Excel header migration (old 9-column file -> current columns)
   - Existing data preservation
   - Row link update
   - Message building with/without link
   - WhatsApp link generation

2. Live Google Drive Test (Requires credentials.json):
   - Generates a dummy PDF
   - Uploads to Google Drive folder ('Payment Receipts' or custom)
   - Checks deduplication / retrieval
   - Prints the shareable link
   - Deletes the uploaded file from Drive to clean up

Usage:
    python test_drive.py          # Runs pure tests cleanly without blocking
    python test_drive.py --live   # Runs pure tests + live Google Drive upload/delete
"""
import argparse
import os
import shutil
import sys
import tempfile
import unittest
from pathlib import Path
from decimal import Decimal
from openpyxl import Workbook, load_workbook
import pandas as pd

from store import (
    HEADERS,
    migrate_excel_if_needed,
    _ensure_workbook,
    save_payment,
    update_receipt_link,
)
from receipts import (
    receipt_message,
    whatsapp_link,
    generate_receipt_pdf,
)
from drive_uploader import (
    upload_receipt,
    get_drive_service,
    CREDENTIALS_FILE,
    TOKEN_FILE,
)


class TestPureLogic(unittest.TestCase):
    """Test pure functions without requiring network access or credentials."""

    def setUp(self):
        self.test_dir = tempfile.mkdtemp()
        self.test_excel = Path(self.test_dir) / "test_payments.xlsx"

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_excel_migration_adds_column_and_preserves_data(self):
        """Test migrating an older 9-column Excel file to 10 columns."""
        old_headers = [
            "Receipt No", "Date", "Name", "Plot No", "Amount",
            "Payment Mode", "Email", "Phone", "Remarks",
        ]
        wb = Workbook()
        ws = wb.active
        ws.title = "Payments"
        ws.append(old_headers)
        ws.append([
            "RCP-0001", "30-10-2026 12:00", "Ramesh Kumar", "Plot 101",
            "5000.00", "UPI", "ramesh@example.com", "9876543210", "Initial deposit"
        ])
        wb.save(self.test_excel)
        wb.close()

        # Run migration
        migrate_excel_if_needed(self.test_excel)

        # Inspect modified Excel file
        wb_migrated = load_workbook(self.test_excel)
        ws_migrated = wb_migrated["Payments"]
        headers = [cell.value for cell in ws_migrated[1] if cell.value is not None]

        # old 9 columns stay first and unchanged; "Receipt Link" and the new merged-app
        # columns (Source, Bank Status, ...) are appended after them
        self.assertEqual(headers[:9], old_headers)
        self.assertEqual(headers[9], "Receipt Link")
        self.assertEqual(headers, HEADERS)

        # Verify old row retained all original values
        row2 = [cell.value for cell in ws_migrated[2]]
        self.assertEqual(row2[0], "RCP-0001")
        self.assertEqual(row2[2], "Ramesh Kumar")
        self.assertEqual(row2[4], "5000.00")
        self.assertIsNone(row2[9])  # Receipt Link is empty for old record
        wb_migrated.close()

        # Verify read with pandas
        df = pd.read_excel(self.test_excel, sheet_name="Payments", dtype=str).fillna("")
        self.assertIn("Receipt Link", df.columns)
        self.assertEqual(df.iloc[0]["Receipt Link"], "")
        self.assertEqual(df.iloc[0]["Name"], "Ramesh Kumar")
        self.assertEqual(df.iloc[0]["Source"], "Admin")  # pre-merge rows are admin entries

    def test_update_receipt_link(self):
        """Test updating Receipt Link for an existing row in Excel."""
        _ensure_workbook(self.test_excel)
        record = {
            "Receipt No": "RCP-0002",
            "Date": "31-10-2026 14:00",
            "Name": "Suresh Sharma",
            "Plot No": "Plot 102",
            "Amount": "10000.00",
            "Payment Mode": "Cash",
            "Email": "suresh@example.com",
            "Phone": "9811122233",
            "Remarks": "Full payment",
            "Receipt Link": "",
        }
        save_payment(record, file_path=self.test_excel)

        # Update the link
        test_link = "https://drive.google.com/file/d/test123abc/view?usp=drivesdk"
        update_receipt_link("RCP-0002", test_link, file_path=self.test_excel)

        df = pd.read_excel(self.test_excel, sheet_name="Payments", dtype=str).fillna("")
        self.assertEqual(df.iloc[0]["Receipt Link"], test_link)

    def test_receipt_message_without_link(self):
        """Receipt message should not include download line if link is absent."""
        rec = {
            "Receipt No": "RCP-0001",
            "Date": "31-10-2026 10:00",
            "Name": "Anita Desai",
            "Plot No": "45",
            "Amount": "2500.00",
            "Receipt Link": "",
        }
        msg = receipt_message(rec)
        self.assertIn("Dear Anita Desai,", msg)
        self.assertIn("Receipt No: RCP-0001", msg)
        self.assertNotIn("Download your receipt:", msg)

    def test_receipt_message_with_link(self):
        """Receipt message should include download line if link is present."""
        test_link = "https://drive.google.com/file/d/xyz789/view"
        rec = {
            "Receipt No": "RCP-0001",
            "Date": "31-10-2026 10:00",
            "Name": "Anita Desai",
            "Plot No": "45",
            "Amount": "2500.00",
            "Receipt Link": test_link,
        }
        msg = receipt_message(rec)
        self.assertIn("Dear Anita Desai,", msg)
        self.assertIn(f"Download your receipt: {test_link}", msg)

    def test_whatsapp_link_includes_drive_link(self):
        """WhatsApp click-to-chat URL must properly encode the receipt download link."""
        test_link = "https://drive.google.com/file/d/xyz789/view"
        rec = {
            "Receipt No": "RCP-0001",
            "Date": "31-10-2026 10:00",
            "Name": "Anita Desai",
            "Plot No": "45",
            "Amount": "2500.00",
            "Phone": "9876543210",
            "Receipt Link": test_link,
        }
        wa_url = whatsapp_link(rec)
        self.assertTrue(wa_url.startswith("https://wa.me/919876543210?text="))
        # Check URL encoded link presence
        self.assertIn("Download%20your%20receipt%3A%20https%3A//drive.google.com", wa_url)


def test_live_drive():
    """
    Live test against Google Drive API.
    Uploads a dummy PDF, verifies the public link, and cleans up by deleting the file from Drive.
    """
    print("\n" + "=" * 60)
    print("LIVE GOOGLE DRIVE TEST")
    print("=" * 60)

    if not CREDENTIALS_FILE.exists() and not TOKEN_FILE.exists():
        print("Skipping live Google Drive upload test.")
        print(f"Reason: credentials.json not found in {CREDENTIALS_FILE.parent}.")
        print("To run the live test, follow the Google Drive setup steps in README.md.")
        return

    # Generate a dummy test PDF
    dummy_pdf_path = Path("receipts") / "DUMMY-TEST-RECEIPT.pdf"
    dummy_pdf_path.parent.mkdir(exist_ok=True)
    rec = {
        "Receipt No": "DUMMY-TEST-RECEIPT",
        "Date": "31-10-2026 16:00",
        "Name": "Google Drive Test User",
        "Plot No": "TEST-00",
        "Amount": "100.00",
        "Payment Mode": "UPI",
        "Email": "",
        "Phone": "",
        "Remarks": "Temporary PDF for test_drive.py",
        "Receipt Link": "",
    }
    generate_receipt_pdf(rec)
    print(f"Created temporary PDF: {dummy_pdf_path}")

    try:
        print("Authenticating and uploading to Google Drive...")
        link = upload_receipt(dummy_pdf_path)
        print("Upload successful!")
        print(f"Shareable Link: {link}")

        # Test deduplication
        print("Testing duplicate upload handling...")
        link_dup = upload_receipt(dummy_pdf_path)
        print(f"Reused existing file link: {link_dup}")
        assert link == link_dup, "Deduplication did not return matching link"

        # Cleanup uploaded file from Google Drive
        service = get_drive_service()
        query = f"name = '{dummy_pdf_path.name}' and trashed = false"
        res = service.files().list(q=query, spaces="drive", fields="files(id, name)").execute()
        files = res.get("files", [])
        for f in files:
            print(f"Deleting test file from Drive (ID: {f['id']})...")
            service.files().delete(fileId=f["id"]).execute()
        print("Drive cleanup complete.")

    except Exception as e:
        print(f"Live Drive test failed with error: {e}")
        raise
    finally:
        if dummy_pdf_path.exists():
            dummy_pdf_path.unlink()
            print(f"Removed local temporary PDF: {dummy_pdf_path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Test payment receipts and Google Drive upload")
    parser.add_argument("--live", action="store_true", help="Run live Google Drive upload test")
    args, _ = parser.parse_known_args()

    print("Running Pure Unit Tests (Offline / Mocked)...")
    suite = unittest.TestLoader().loadTestsFromTestCase(TestPureLogic)
    # Output to sys.stdout instead of sys.stderr so PowerShell doesn't treat test output as NativeCommandError
    runner = unittest.TextTestRunner(stream=sys.stdout, verbosity=2)
    result = runner.run(suite)

    if not result.wasSuccessful():
        sys.exit(1)

    print("\nAll pure tests passed successfully.")

    # Determine whether to run live test
    if args.live:
        test_live_drive()
    elif TOKEN_FILE.exists():
        test_live_drive()
    elif CREDENTIALS_FILE.exists():
        print("\n" + "=" * 60)
        print("LIVE GOOGLE DRIVE SETUP NOTICE")
        print("=" * 60)
        print("Found credentials.json! To authorize the app via browser consent and")
        print("run the live Google Drive test, run:")
        print("    python test_drive.py --live")
        print("=" * 60)
    else:
        print("\nNote: Place credentials.json in this folder to enable live Drive uploads.")
