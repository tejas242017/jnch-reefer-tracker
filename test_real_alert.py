import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Uses credentials from GitHub Actions or local environment
GMAIL_SENDER = os.getenv("GMAIL_SENDER", "")
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD", "")
ALERT_RECEIVER = os.getenv("ALERT_RECEIVER", GMAIL_SENDER)

def send_test_email():
    if not GMAIL_SENDER or not GMAIL_APP_PASSWORD:
        print("[!] GMAIL credentials missing. Ensure GMAIL_SENDER and GMAIL_APP_PASSWORD secrets are set on GitHub.")
        return

    subject = "🚨 JNCH DPD Reefer Alert: MSC Mandarins Detected [MEDUW5386315]"
    
    # Real test data from your historical discovery
    containers = [
        {"cntr": "MEDU9839122", "iso": "40RH", "weight": "24,156 KGS"},
        {"cntr": "BMOU9286052", "iso": "40RH", "weight": "24,157 KGS"},
        {"cntr": "MEDU9895717", "iso": "40RH", "weight": "24,157 KGS"}
    ]
    
    master_bl = "MEDUW5386315"
    vessel = "MSC VOYAGER"
    voyage = "ZF634R"
    pol = "DURBAN, SOUTH AFRICA (ZA)"
    total_gross = "72,470.8 KGS"
    party = "DIRECT PORT DELIVERY (DPD)"
    
    rows_html = ""
    for c in containers:
        ldb_link = f"https://ldb.co.in/ldb/containersearch/39/{c['cntr']}/"
        bmct_link = "https://eportal.bmctpl.com/eform/transactions/ContainerTracking.aspx"
        
        rows_html += f"""
        <tr style="border-bottom: 1px solid #e0e0e0;">
            <td style="padding: 10px; font-weight: bold; font-family: monospace; font-size: 14px;">
                <a href="{ldb_link}" target="_blank" style="color: #1a73e8; text-decoration: none;">{c['cntr']}</a>
            </td>
            <td style="padding: 10px; color: #d93025; font-weight: bold;">{c['iso']}</td>
            <td style="padding: 10px;">{c['weight']}</td>
            <td style="padding: 10px;">
                <a href="{ldb_link}" style="color: #1a73e8; text-decoration: none;">LDB Status</a> | 
                <a href="{bmct_link}" style="color: #1a73e8; text-decoration: none;">BMCT CFS</a>
            </td>
        </tr>
        """

    icegate_url = "https://foservices.icegate.gov.in/#/public-enquiries/document-status/sea-igm"

    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
    </head>
    <body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; line-height: 1.5; color: #202124; background-color: #f8f9fa; margin: 0; padding: 20px;">
        <div style="max-width: 650px; margin: 0 auto; background: #ffffff; border-radius: 8px; border: 1px solid #dadce0; overflow: hidden; padding: 24px;">
            
            <div style="border-bottom: 2px solid #1a73e8; padding-bottom: 12px; margin-bottom: 20px;">
                <h2 style="color: #1a73e8; margin: 0; font-size: 20px;">🚨 Reefer Import Detected at Nhava Sheva (JNCH)</h2>
                <p style="margin: 4px 0 0 0; color: #5f6368; font-size: 13px;">Master B/L resolved via automated BMCT terminal query</p>
            </div>

            <!-- Master Consignment Card -->
            <div style="background-color: #f1f3f4; border-radius: 6px; padding: 14px; margin-bottom: 20px; font-size: 13px;">
                <table style="width: 100%; border-collapse: collapse;">
                    <tr>
                        <td style="padding: 4px 0; color: #5f6368; width: 35%;"><strong>Master B/L No:</strong></td>
                        <td style="padding: 4px 0; font-weight: bold; color: #1a73e8; font-family: monospace; font-size: 15px;">{master_bl}</td>
                    </tr>
                    <tr>
                        <td style="padding: 4px 0; color: #5f6368;"><strong>Shipping Line / Vessel:</strong></td>
                        <td style="padding: 4px 0;">MSC | {vessel} ({voyage})</td>
                    </tr>
                    <tr>
                        <td style="padding: 4px 0; color: #5f6368;"><strong>Port of Loading (POL):</strong></td>
                        <td style="padding: 4px 0;">{pol}</td>
                    </tr>
                    <tr>
                        <td style="padding: 4px 0; color: #5f6368;"><strong>Total Consignment Wt:</strong></td>
                        <td style="padding: 4px 0; font-weight: bold;">{total_gross}</td>
                    </tr>
                    <tr>
                        <td style="padding: 4px 0; color: #5f6368;"><strong>Delivery Nomination:</strong></td>
                        <td style="padding: 4px 0;">{party}</td>
                    </tr>
                </table>
            </div>

            <h3 style="font-size: 15px; margin: 0 0 10px 0; color: #202124;">Linked Reefer Containers in Manifest:</h3>
            <table style="width: 100%; border-collapse: collapse; font-size: 13px; margin-bottom: 24px;">
                <thead>
                    <tr style="background-color: #f8f9fa; border-bottom: 2px solid #dadce0; text-align: left;">
                        <th style="padding: 8px 10px;">Container No</th>
                        <th style="padding: 8px 10px;">Equipment</th>
                        <th style="padding: 8px 10px;">Gross Weight</th>
                        <th style="padding: 8px 10px;">Actions</th>
                    </tr>
                </thead>
                <tbody>
                    {rows_html}
                </tbody>
            </table>

            <!-- Quick Action for Invoices & Cartons -->
            <div style="background-color: #e8f0fe; border-left: 4px solid #1a73e8; padding: 14px; border-radius: 4px;">
                <strong style="color: #1a73e8; font-size: 14px;">Verify Invoices, Cartons & Importer on ICEGATE:</strong>
                <p style="margin: 6px 0 12px 0; font-size: 12px; color: #3c4043;">
                    Because ICEGATE protects invoice declarations with a security CAPTCHA, click below to open the inquiry pre-loaded for <strong>INNSA1</strong>:
                </p>
                <div style="text-align: center; margin-bottom: 10px;">
                    <a href="{icegate_url}" target="_blank" style="background-color: #1a73e8; color: #ffffff; padding: 10px 20px; font-weight: bold; text-decoration: none; border-radius: 4px; font-size: 13px; display: inline-block;">Open ICEGATE Sea-IGM Enquiry</a>
                </div>
                <p style="margin: 0; font-size: 11px; color: #5f6368; text-align: center;">
                    Enter Location: <strong>INNSA1</strong> | Master BL: <strong>{master_bl}</strong>
                </p>
            </div>

            <p style="font-size: 11px; color: #9aa0a6; margin-top: 20px; text-align: center;">
                Generated automatically by JNCH Reefer Tracker (GitHub Cloud Runner).
            </p>
        </div>
    </body>
    </html>
    """

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = GMAIL_SENDER
    msg["To"] = ALERT_RECEIVER
    msg.attach(MIMEText(html_content, "html"))

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(GMAIL_SENDER, GMAIL_APP_PASSWORD)
            server.sendmail(GMAIL_SENDER, ALERT_RECEIVER, msg.as_string())
        print(f"[+] SUCCESS: Test alert sent directly to {ALERT_RECEIVER}!")
    except Exception as e:
        print(f"[!] Email dispatch failed: {e}")

if __name__ == "__main__":
    send_test_email()