import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

GMAIL_SENDER = (os.getenv("GMAIL_SENDER") or "").strip()
GMAIL_APP_PASSWORD = (os.getenv("GMAIL_APP_PASSWORD") or "").strip()
ALERT_RECEIVER = (os.getenv("ALERT_RECEIVER") or GMAIL_SENDER).strip()

def send_full_intelligence_email():
    if not GMAIL_SENDER or not GMAIL_APP_PASSWORD:
        print("[!] GMAIL credentials missing.")
        return

    subject = "🚨 [VASHI APMC REPORT] Fresh Mandarins: 3 Reefer(s) | 6,296 Cartons @ Ameya CFS"

    containers_data = [
        {
            "cntr": "MEDU9839122",
            "iso": "4532 (40' High Cube)",
            "temp": "2.0°C",
            "weight": "29,070.80 KGS",
            "terminal": "BMCT (PSA Mumbai)",
            "term_gate_out": "02-Oct-2026 15:22 (Truck: MH43BP7288 / MH46BB6315)",
            "cfs_yard": "Ameya Logistics CFS, Navi Mumbai",
            "cfs_status": "CFS OUT (07-Oct-2026)",
            "market_verdict": "APMC ARRIVED / DISPATCHED"
        },
        {
            "cntr": "BMOU9286052",
            "iso": "4532 (40' High Cube)",
            "temp": "2.0°C",
            "weight": "28,380.00 KGS",
            "terminal": "BMCT (PSA Mumbai)",
            "term_gate_out": "02-Oct-2026 15:22",
            "cfs_yard": "Ameya Logistics CFS, Navi Mumbai",
            "cfs_status": "CFS IN (Customs / PQ Hold)",
            "market_verdict": "HOLDING AT CFS"
        },
        {
            "cntr": "MEDU9895717",
            "iso": "4532 (40' High Cube)",
            "temp": "2.0°C",
            "weight": "28,560.00 KGS",
            "terminal": "BMCT (PSA Mumbai)",
            "term_gate_out": "02-Oct-2026 15:22",
            "cfs_yard": "Ameya Logistics CFS, Navi Mumbai",
            "cfs_status": "CFS OUT (07-Oct-2026)",
            "market_verdict": "APMC ARRIVED / DISPATCHED"
        }
    ]

    cards_html = ""
    for c in containers_data:
        ldb_link = f"https://ldb.co.in/ldb/containersearch/39/{c['cntr']}"
        status_color = "#d93025" if "CFS OUT" in c['cfs_status'] else "#137333"

        cards_html += f"""
        <div style="background: #ffffff; border: 1px solid #e0e0e0; border-radius: 6px; margin-bottom: 16px; padding: 16px; border-left: 5px solid {status_color};">
            <div style="display: flex; justify-content: space-between; border-bottom: 1px solid #f1f3f4; padding-bottom: 8px; margin-bottom: 10px;">
                <span style="font-size: 16px; font-weight: bold; font-family: monospace; color: #1a73e8;">
                    <a href="{ldb_link}" target="_blank" style="text-decoration: none; color: #1a73e8;">{c['cntr']}</a>
                </span>
                <span style="background-color: #f1f3f4; color: #202124; padding: 3px 8px; border-radius: 4px; font-size: 12px; font-weight: bold;">
                    ISO: {c['iso']} | {c['temp']}
                </span>
            </div>
            
            <table style="width: 100%; border-collapse: collapse; font-size: 13px; line-height: 1.6;">
                <tr>
                    <td style="color: #5f6368; width: 32%;"><strong>Gross Cargo Weight:</strong></td>
                    <td style="color: #202124; font-weight: bold;">{c['weight']}</td>
                </tr>
                <tr>
                    <td style="color: #5f6368;"><strong>Berth & Terminal Gate:</strong></td>
                    <td style="color: #202124;">{c['terminal']} &rarr; Out: {c['term_gate_out']}</td>
                </tr>
                <tr>
                    <td style="color: #5f6368;"><strong>Nominated CFS Yard:</strong></td>
                    <td style="color: #202124; font-weight: bold;">{c['cfs_yard']}</td>
                </tr>
                <tr>
                    <td style="color: #5f6368;"><strong>Current CFS Status:</strong></td>
                    <td style="color: {status_color}; font-weight: bold;">{c['cfs_status']}</td>
                </tr>
                <tr>
                    <td style="color: #5f6368;"><strong>Market Decision Verdict:</strong></td>
                    <td style="color: {status_color}; font-weight: bold;">{c['market_verdict']}</td>
                </tr>
            </table>
        </div>
        """

    html_content = f"""
    <!DOCTYPE html>
    <html>
    <body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; background-color: #f8f9fa; padding: 15px; margin: 0; color: #202124;">
        <div style="max-width: 720px; margin: 0 auto; background: #ffffff; border-radius: 8px; border: 1px solid #dadce0; overflow: hidden;">
            <div style="background-color: #1a73e8; color: white; padding: 20px 24px;">
                <h2 style="margin: 0; font-size: 20px;">🍎 Inbound Fruit Intelligence — Container Detail Audit</h2>
                <p style="margin: 4px 0 0 0; font-size: 13px; opacity: 0.95;">Nhava Sheva (JNCH) &bull; Sea-IGM: 1214081 &bull; Master B/L: MEDUW5386315</p>
            </div>
            <div style="padding: 20px;">
                <div style="background-color: #e8f0fe; border-left: 4px solid #1a73e8; padding: 14px 18px; border-radius: 4px; margin-bottom: 22px;">
                    <h3 style="margin: 0 0 10px 0; font-size: 15px; color: #1a73e8;">Consignment Overview (ICEGATE & Manifest)</h3>
                    <table style="width: 100%; border-collapse: collapse; font-size: 13px; line-height: 1.5;">
                        <tr>
                            <td style="color: #5f6368; width: 35%;"><strong>Cargo Declaration:</strong></td>
                            <td style="color: #d93025; font-weight: bold;">FRESH MANDARINS</td>
                        </tr>
                        <tr>
                            <td style="color: #5f6368;"><strong>Total Consignment Volume:</strong></td>
                            <td style="font-weight: bold; color: #202124;">6,296 Cartons (72,470.8 KGS)</td>
                        </tr>
                        <tr>
                            <td style="color: #5f6368;"><strong>Invoices Declared:</strong></td>
                            <td style="font-family: monospace; font-weight: bold;">C042946, C042853, C042854 (Dated 11-Sep-2026)</td>
                        </tr>
                        <tr>
                            <td style="color: #5f6368;"><strong>Vessel & Voyage:</strong></td>
                            <td>MSC BARBARA (Voyage: ZF634R) &bull; Discharged BMCT</td>
                        </tr>
                        <tr>
                            <td style="color: #5f6368;"><strong>Port of Loading (POL):</strong></td>
                            <td>Durban, South Africa (ZA)</td>
                        </tr>
                        <tr>
                            <td style="color: #5f6368;"><strong>Customs Delivery Mode:</strong></td>
                            <td>Direct Port Delivery (DPD - Client Code: 4QF)</td>
                        </tr>
                    </table>
                </div>
                <div style="background-color: #fef7e0; border-left: 4px solid #f9ab00; padding: 12px 16px; border-radius: 4px; margin-bottom: 22px; font-size: 13px;">
                    <strong style="color: #b06000;">Market Supply Insight:</strong>
                    2 out of 3 containers completed <strong>CFS OUT on 07-Oct-2026</strong>[cite: 4]. Wholesale arrivals will hit the Vashi market tonight. 
                    <strong>Recommendation:</strong> If holding sister consignments, avoid immediate liquidation to bypass price suppression from this delivery wave.
                </div>
                <h3 style="font-size: 15px; margin: 0 0 12px 0; color: #202124;">Physical Movement Breakdown (Container Wise):</h3>
                {cards_html}
                <div style="text-align: center; margin-top: 20px; font-size: 11px; color: #80868b; border-top: 1px solid #f1f3f4; padding-top: 15px;">
                    Automated JNCH Reefer Intelligence &bull; Continuous 4-Hour Polling Cycle
                </div>
            </div>
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
        print(f"[+] SUCCESS: Detailed container report delivered to {ALERT_RECEIVER}!")
    except Exception as e:
        print(f"[!] Email dispatch error: {e}")

if __name__ == "__main__":
    send_full_intelligence_email()
