"""PDF receipt, e-mail and WhatsApp helpers (ported unchanged from the admin app,
except that file names are made safe because online Receipt Nos are transaction IDs)."""
import re
import smtplib
from decimal import Decimal, ROUND_HALF_UP
from email.message import EmailMessage
from pathlib import Path
from urllib.parse import quote

from reportlab.lib import colors
from reportlab.lib.pagesizes import A5
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from settings import (DEFAULT_COUNTRY_CODE, ORG_ADDRESS, ORG_NAME, ORG_PHONE, RECEIPT_DIR,
                      SMTP_HOST, SMTP_PASSWORD, SMTP_PORT, SMTP_USER)

# ---- amount in words (Indian numbering: lakh / crore) ----
_ONES = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
         "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
         "Seventeen", "Eighteen", "Nineteen"]
_TENS = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]


def _below_hundred(n: int) -> str:
    return _ONES[n] if n < 20 else (_TENS[n // 10] + (" " + _ONES[n % 10] if n % 10 else ""))


def _below_thousand(n: int) -> str:
    parts = []
    if n >= 100:
        parts.append(_ONES[n // 100] + " Hundred")
        n %= 100
    if n:
        parts.append(_below_hundred(n))
    return " ".join(parts)


def _int_to_words(n: int) -> str:
    if n == 0:
        return "Zero"
    parts = []
    for value, name in ((10_000_000, "Crore"), (100_000, "Lakh"), (1000, "Thousand")):
        if n >= value:
            parts.append(_below_hundred(n // value) + " " + name)
            n %= value
    if n:
        parts.append(_below_thousand(n))
    return " ".join(parts)


def amount_in_words(amount) -> str:
    amount = Decimal(str(amount)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    rupees = int(amount)
    paise = int((amount - rupees) * 100)
    words = "Rupees " + _int_to_words(rupees)
    if paise:
        words += " and " + _below_hundred(paise) + " Paise"
    return words + " Only"


# ---- PDF ----
def receipt_path(receipt_no: str) -> Path:
    safe = re.sub(r"[^A-Za-z0-9._-]", "_", str(receipt_no))
    return RECEIPT_DIR / f"{safe}.pdf"


def generate_receipt_pdf(rec: dict) -> Path:
    path = receipt_path(rec["Receipt No"])
    doc = SimpleDocTemplate(
        str(path), pagesize=A5,
        leftMargin=14 * mm, rightMargin=14 * mm, topMargin=12 * mm, bottomMargin=12 * mm,
        title=f"Receipt {rec['Receipt No']}",
    )
    ss = getSampleStyleSheet()
    org = ParagraphStyle("org", parent=ss["Title"], fontSize=16, spaceAfter=2)
    small = ParagraphStyle("small", parent=ss["Normal"], fontSize=8.5, alignment=1,
                           textColor=colors.HexColor("#555555"))
    head = ParagraphStyle("head", parent=ss["Heading2"], alignment=1, spaceBefore=8, spaceAfter=8)
    body = ParagraphStyle("body", parent=ss["Normal"], fontSize=10, leading=14)

    amount_fmt = f"Rs. {Decimal(rec['Amount']):,.2f}"
    rows = [
        ["Receipt No.", Paragraph(str(rec["Receipt No"]), body)],
        ["Date", rec["Date"]],
        ["Received from", rec["Name"]],
        ["Plot No.", rec["Plot No"]],
        ["Payment mode", rec["Payment Mode"]],
        ["Amount", amount_fmt],
    ]
    if rec.get("Remarks"):
        rows.append(["Remarks", Paragraph(rec["Remarks"], body)])

    table = Table(rows, colWidths=[38 * mm, 78 * mm])
    table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#999999")),
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#F1F1F1")),
        ("FONTNAME", (1, 5), (1, 5), "Helvetica-Bold"),
        ("FONTSIZE", (1, 5), (1, 5), 12),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))

    story = [
        Paragraph(ORG_NAME, org),
        Paragraph(ORG_ADDRESS + (f" | Ph: {ORG_PHONE}" if ORG_PHONE else ""), small),
        Paragraph("PAYMENT RECEIPT", head),
        table,
        Spacer(1, 8),
        Paragraph(f"<b>Amount in words:</b> {amount_in_words(rec['Amount'])}", body),
        Spacer(1, 26),
        Paragraph("Authorised Signature", ParagraphStyle("sig", parent=body, alignment=2)),
        Spacer(1, 10),
        Paragraph("This is a computer-generated receipt.", small),
    ]
    doc.build(story)
    return path


# ---- e-mail / WhatsApp ----
def receipt_message(rec: dict) -> str:
    msg = (
        f"Dear {rec['Name']},\n\n"
        f"We have received Rs. {Decimal(rec['Amount']):,.2f} towards Plot No. {rec['Plot No']} "
        f"on {rec['Date']}.\nReceipt No: {rec['Receipt No']}\n\n"
    )
    link = str(rec.get("Receipt Link") or "").strip()
    if link:
        msg += f"Download your receipt: {link}\n\n"
    msg += f"Thank you.\n{ORG_NAME}"
    return msg


def send_email(rec: dict, pdf_path: Path):
    if not (SMTP_USER and SMTP_PASSWORD):
        raise RuntimeError("Email is not configured. Set SMTP_USER and SMTP_PASSWORD (see README).")
    msg = EmailMessage()
    msg["Subject"] = f"Payment Receipt {rec['Receipt No']} - {ORG_NAME}"
    msg["From"] = f"{ORG_NAME} <{SMTP_USER}>"
    msg["To"] = rec["Email"]
    msg.set_content(receipt_message(rec))
    msg.add_attachment(pdf_path.read_bytes(), maintype="application", subtype="pdf",
                       filename=pdf_path.name)
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
        server.login(SMTP_USER, SMTP_PASSWORD)
        server.send_message(msg)


def clean_phone(phone: str) -> str:
    digits = re.sub(r"\D", "", phone or "")
    if len(digits) == 10:
        digits = DEFAULT_COUNTRY_CODE + digits
    return digits


def whatsapp_link(rec: dict) -> str:
    return f"https://wa.me/{clean_phone(rec['Phone'])}?text={quote(receipt_message(rec))}"
