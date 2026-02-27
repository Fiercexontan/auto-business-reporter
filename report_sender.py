"""
====================================================
  AUTO BUSINESS REPORT SENDER
  Built by: Lincoln Adura
  Project:  Portfolio Project #1
  What it does: Reads data from Google Sheets,
                builds a clean HTML report,
                and emails it automatically.
====================================================
"""

import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
import gspread # pyright: ignore[reportMissingImports]
from google.oauth2.service_account import Credentials # pyright: ignore[reportMissingImports]
from dotenv import load_dotenv # pyright: ignore[reportMissingImports]

# ── Load environment variables from .env file ──
load_dotenv()

# ── CONFIG (reads from your .env file) ──
SENDER_EMAIL    = os.getenv("SENDER_EMAIL")       # Your Gmail address
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")    # Your Gmail App Password
RECIPIENT_EMAIL = os.getenv("RECIPIENT_EMAIL")    # Who receives the report
SHEET_NAME      = os.getenv("SHEET_NAME")         # Your Google Sheet name
BUSINESS_NAME   = os.getenv("BUSINESS_NAME", "Your Business")


# ══════════════════════════════════════════
#  STEP 1 — CONNECT TO GOOGLE SHEETS
# ══════════════════════════════════════════

def get_sheet_data():
    """
    Connects to Google Sheets using a service account
    and returns all rows from the first sheet.
    """
    print("📊 Connecting to Google Sheets...")

    # Tell Google which permissions we need
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly"
    ]

    # Load credentials from the JSON file you downloaded from Google Cloud
    creds = Credentials.from_service_account_file(
        "credentials.json",
        scopes=scopes
    )

    # Authorize and open the sheet
    client = gspread.authorize(creds)
    sheet  = client.open(SHEET_NAME).sheet1

    # Get all values as a list of lists
    data = sheet.get_all_records()  # Returns list of dicts using first row as keys

    print(f"✅ Found {len(data)} rows of data.")
    return data


# ══════════════════════════════════════════
#  STEP 2 — BUILD THE REPORT
# ══════════════════════════════════════════

def build_report(data):
    """
    Takes the sheet data and builds a clean HTML email report.
    This is what the client sees in their inbox.
    """
    print("📝 Building report...")

    today      = datetime.now().strftime("%B %d, %Y")
    week_start = datetime.now().strftime("%A, %B %d")

    # ── Calculate summary stats (adjust column names to match your sheet) ──
    # We assume your sheet has columns: Item, Sales, Revenue, Status
    # Change these to match whatever YOUR sheet has
    try:
        total_sales   = sum(int(row.get("Sales", 0))   for row in data)
        total_revenue = sum(float(row.get("Revenue", 0)) for row in data)
        completed     = sum(1 for row in data if str(row.get("Status", "")).lower() == "completed")
    except Exception:
        total_sales, total_revenue, completed = 0, 0, 0

    # ── Build the table rows from sheet data ──
    table_rows = ""
    for i, row in enumerate(data):
        bg = "#f9f7f4" if i % 2 == 0 else "#ffffff"
        cells = "".join(f"<td style='padding:10px 14px; border-bottom:1px solid #eee;'>{v}</td>"
                        for v in row.values())
        table_rows += f"<tr style='background:{bg};'>{cells}</tr>"

    # ── Column headers ──
    if data:
        headers = "".join(
            f"<th style='padding:12px 14px; text-align:left; background:#1a1a2e; color:#fff; font-weight:600;'>{col}</th>"
            for col in data[0].keys()
        )
    else:
        headers = "<th>No data found</th>"

    # ── Full HTML email template ──
    html = f"""
    <!DOCTYPE html>
    <html>
    <head><meta charset="UTF-8"/></head>
    <body style="margin:0; padding:0; background:#f0ede8; font-family: 'Helvetica Neue', Arial, sans-serif;">

      <div style="max-width:640px; margin:32px auto; background:#fff; border-radius:8px; overflow:hidden; box-shadow:0 2px 12px rgba(0,0,0,0.08);">

        <!-- HEADER -->
        <div style="background:#1a1a2e; padding:32px 36px;">
          <div style="font-size:12px; letter-spacing:0.2em; color:#7a8aaa; text-transform:uppercase; margin-bottom:8px;">Automated Weekly Report</div>
          <div style="font-size:26px; font-weight:700; color:#ffffff;">{BUSINESS_NAME}</div>
          <div style="font-size:13px; color:#9aa3b8; margin-top:6px;">Generated on {today}</div>
        </div>

        <!-- SUMMARY STATS -->
        <div style="padding:28px 36px; background:#f9f7f4; border-bottom:1px solid #eee;">
          <div style="font-size:11px; letter-spacing:0.15em; color:#999; text-transform:uppercase; margin-bottom:16px;">Summary</div>
          <div style="display:flex; gap:20px; flex-wrap:wrap;">

            <div style="flex:1; min-width:140px; background:#fff; border-radius:6px; padding:18px 20px; border:1px solid #eee;">
              <div style="font-size:28px; font-weight:700; color:#1a1a2e;">{total_sales}</div>
              <div style="font-size:12px; color:#999; margin-top:4px;">Total Sales</div>
            </div>

            <div style="flex:1; min-width:140px; background:#fff; border-radius:6px; padding:18px 20px; border:1px solid #eee;">
              <div style="font-size:28px; font-weight:700; color:#2d6a4f;">${total_revenue:,.2f}</div>
              <div style="font-size:12px; color:#999; margin-top:4px;">Total Revenue</div>
            </div>

            <div style="flex:1; min-width:140px; background:#fff; border-radius:6px; padding:18px 20px; border:1px solid #eee;">
              <div style="font-size:28px; font-weight:700; color:#c8531a;">{completed}</div>
              <div style="font-size:12px; color:#999; margin-top:4px;">Completed Orders</div>
            </div>

          </div>
        </div>

        <!-- DATA TABLE -->
        <div style="padding:28px 36px;">
          <div style="font-size:11px; letter-spacing:0.15em; color:#999; text-transform:uppercase; margin-bottom:16px;">Full Data Breakdown</div>
          <div style="overflow-x:auto;">
            <table style="width:100%; border-collapse:collapse; font-size:13px;">
              <thead><tr>{headers}</tr></thead>
              <tbody>{table_rows}</tbody>
            </table>
          </div>
        </div>

        <!-- FOOTER -->
        <div style="padding:20px 36px; background:#f9f7f4; border-top:1px solid #eee; text-align:center;">
          <div style="font-size:12px; color:#bbb;">This report was generated automatically · Built by Lincoln Adura</div>
          <div style="font-size:11px; color:#ccc; margin-top:4px;">Automation Developer · Web Developer</div>
        </div>

      </div>
    </body>
    </html>
    """

    print("✅ Report built.")
    return html


# ══════════════════════════════════════════
#  STEP 3 — SEND THE EMAIL
# ══════════════════════════════════════════

def send_email(html_content):
    """
    Sends the HTML report via Gmail using SMTP.
    """
    print(f"📧 Sending report to {RECIPIENT_EMAIL}...")

    today   = datetime.now().strftime("%B %d, %Y")
    subject = f"📊 Weekly Business Report — {today} | {BUSINESS_NAME}"

    # Create the email
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = SENDER_EMAIL
    msg["To"]      = RECIPIENT_EMAIL

    # Attach the HTML content
    msg.attach(MIMEText(html_content, "html"))

    # Connect to Gmail's SMTP server and send
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, RECIPIENT_EMAIL, msg.as_string())

    print("✅ Email sent successfully!")


# ══════════════════════════════════════════
#  MAIN — Run everything
# ══════════════════════════════════════════

if __name__ == "__main__":
    print("\n🚀 Starting Auto Report Sender...")
    print("=" * 45)

    try:
        data         = get_sheet_data()
        html_report  = build_report(data)
        send_email(html_report)

        print("=" * 45)
        print("🎉 Done! Report sent successfully.\n")

    except Exception as e:
        print(f"\n❌ Error: {e}")
        print("Check your .env file and credentials.json.\n")