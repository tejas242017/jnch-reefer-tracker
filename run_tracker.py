import os
import re
import sys
import smtplib
import asyncio
import subprocess
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd

# ----------------- NEW DYNAMIC CODE -----------------
# Auto-detect Environment: Local E: Drive vs GitHub Actions Cloud
IS_WINDOWS = sys.platform == "win32"

if IS_WINDOWS:
    BASE_DIR = r"E:\DPD_Tracker_Sandbox"
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(BASE_DIR, "pw-browsers")
else:
    # Linux (GitHub Actions)
    BASE_DIR = os.getcwd()

from playwright.async_api import async_playwright

DOWNLOAD_DIR = os.path.join(BASE_DIR, "manifests")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

MASTER_LOG_PATH = os.path.join(BASE_DIR, "detected_reefers.csv")
SEEN_CONTAINERS_FILE = os.path.join(BASE_DIR, "seen_containers.txt")

# Gmail Alerts Settings (reads from GitHub Secrets in cloud, or falls back to local strings)
GMAIL_SENDER = os.getenv("GMAIL_SENDER", "YOUR_GMAIL@gmail.com")
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD", "YOUR_16_CHAR_APP_PASSWORD")
ALERT_RECEIVER = os.getenv("ALERT_RECEIVER", GMAIL_SENDER)

# Carriers handling Perishables (Vietnam / South Africa)
TARGET_LINES = ["WAN HAI", "ONE", "CMA CGM", "MAERSK", "RCL", "SAMUDERA", "COSCO", "MSC", "HYUNDAI", "HMM"]
REEFER_CODES = ["45R1", "4532", "42R1", "22R1", "40RH", "40RF", "20RF", "RF", "RH", "REEF"]
# -------------------------------------------------

def get_seen_containers():
    if not os.path.exists(SEEN_CONTAINERS_FILE):
        return set()
    with open(SEEN_CONTAINERS_FILE, "r") as f:
        return set(line.strip() for line in f if line.strip())

def mark_containers_seen(new_cntrs):
    with open(SEEN_CONTAINERS_FILE, "a") as f:
        for c in new_cntrs:
            f.write(f"{c}\n")

def read_any_format(filepath):
    try:
        return pd.read_excel(filepath, engine="openpyxl")
    except Exception:
        pass
    try:
        return pd.read_excel(filepath, engine="xlrd")
    except Exception:
        pass
    try:
        tables = pd.read_html(filepath, flavor="html5lib")
        if tables:
            return tables[0]
    except Exception:
        pass
    try:
        return pd.read_csv(filepath, sep=None, engine="python", on_bad_lines="skip")
    except Exception:
        pass
    return None

def normalize_manifest_df(df):
    if df is None or df.empty:
        return None

    first_col_val = str(df.iloc[0, 0]).strip().upper()
    if "HDR" in first_col_val or "HEADER" in first_col_val:
        for offset in range(1, 4):
            candidate_row = [str(x).strip().upper() for x in df.iloc[offset]]
            if any("CONTAINER" in c or "CNTR" in c for c in candidate_row):
                df.columns = candidate_row
                df = df.iloc[offset + 1:].copy()
                break

    df.columns = [str(c).strip().upper() for c in df.columns]
    return df

def parse_reefer_manifest(filepath, line, vessel, voyage):
    try:
        raw_df = read_any_format(filepath)
        df = normalize_manifest_df(raw_df)

        if df is None or df.empty:
            return None

        type_col = next((c for c in df.columns if any(k in c for k in ["ISO", "TYPE", "SIZE", "EQPTYPE", "EQ_TYPE"])), None)
        cntr_col = next((c for c in df.columns if any(k in c for k in ["CONTAINER", "CNTR", "EQ_NO"])), None)
        pol_col = next((c for c in df.columns if any(k in c for k in ["POL", "LOAD", "ORIGIN"])), None)
        party_col = next((c for c in df.columns if any(k in c for k in ["CONSIGNEE", "PARTY", "CLIENT", "IMPORTER", "DPD", "GROUPCODE"])), None)
        weight_col = next((c for c in df.columns if any(k in c for k in ["WEIGHT", "GROSS", "WT"])), None)

        if not type_col or not cntr_col:
            return None

        reefer_regex = "|".join([rf"\b{re.escape(code)}\b" for code in REEFER_CODES]) + r"|RH|RF|REEF"
        reefer_mask = df[type_col].astype(str).str.contains(reefer_regex, case=False, na=False)
        reefers = df[reefer_mask].copy()

        if reefers.empty:
            return None

        reefers["_LINE"] = line
        reefers["_VESSEL"] = vessel
        reefers["_VOYAGE"] = voyage
        reefers["_CNTR"] = reefers[cntr_col].astype(str).str.strip()
        reefers["_ISO"] = reefers[type_col].astype(str).str.strip()
        reefers["_POL"] = reefers[pol_col].astype(str).str.strip() if pol_col else "N/A"
        reefers["_WEIGHT"] = reefers[weight_col].astype(str).str.strip() if weight_col else "N/A"
        reefers["_PARTY"] = reefers[party_col].astype(str).str.strip() if party_col else "N/A"

        return reefers

    except Exception as e:
        print(f"    [!] Error parsing {filepath}: {e}")
        return None

def send_gmail_alert(new_reefers_df):
    """Sends a clean HTML email alert with direct tracking links."""
    if GMAIL_SENDER == "YOUR_GMAIL@gmail.com":
        print("[!] Gmail credentials not set. Skipping email dispatch.")
        return

    subject = f"🚨 JNCH DPD Reefer Alert: {len(new_reefers_df)} Inbound Reefer(s) Detected!"
    
    rows_html = ""
    for _, row in new_reefers_df.iterrows():
        c_no = row["_CNTR"]
        ldb_link = f"https://ldb.co.in/ldb/containersearch/39/{c_no}/"
        bmct_link = "https://eportal.bmctpl.com/eform/transactions/ContainerTracking.aspx"
        
        rows_html += f"""
        <tr style="border-bottom: 1px solid #ddd;">
            <td style="padding: 8px; font-weight: bold;"><a href="{ldb_link}" target="_blank">{c_no}</a></td>
            <td style="padding: 8px;">{row['_ISO']}</td>
            <td style="padding: 8px;">{row['_LINE']}</td>
            <td style="padding: 8px;">{row['_VESSEL']} ({row['_VOYAGE']})</td>
            <td style="padding: 8px;">{row['_POL']}</td>
            <td style="padding: 8px;">{row['_WEIGHT']}</td>
            <td style="padding: 8px;">{row['_PARTY']}</td>
            <td style="padding: 8px;"><a href="{ldb_link}">LDB Track</a> | <a href="{bmct_link}">BMCT</a></td>
        </tr>
        """

    html_content = f"""
    <html>
    <body style="font-family: Arial, sans-serif;">
        <h2 style="color: #1a73e8;">New Reefer Containers Inbound at Nhava Sheva (JNCH)</h2>
        <p>The automated DPD Advance Listing scanner detected new reefer filings:</p>
        <table style="width: 100%; border-collapse: collapse; text-align: left; font-size: 13px;">
            <thead>
                <tr style="background-color: #f2f2f2; border-bottom: 2px solid #ccc;">
                    <th style="padding: 8px;">Container No</th>
                    <th style="padding: 8px;">ISO</th>
                    <th style="padding: 8px;">Line</th>
                    <th style="padding: 8px;">Vessel</th>
                    <th style="padding: 8px;">POL</th>
                    <th style="padding: 8px;">Gross Wt</th>
                    <th style="padding: 8px;">Party / Code</th>
                    <th style="padding: 8px;">Live Track</th>
                </tr>
            </thead>
            <tbody>
                {rows_html}
            </tbody>
        </table>
        <br>
        <p style="font-size: 12px; color: #666;">Data synced to GitHub repo from E:\\DPD_Tracker_Sandbox.</p>
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
        print(f"[+] Successfully sent Gmail alert for {len(new_reefers_df)} reefers!")
    except Exception as e:
        print(f"[!] Failed to send email alert: {e}")

def git_commit_and_push():
    """Commits and pushes updated data directly from E: drive sandbox to GitHub."""
    try:
        subprocess.run(["git", "add", "detected_reefers.csv", "seen_containers.txt"], cwd=SANDBOX_DIR, check=True)
        # Check if there are changes to commit
        status = subprocess.run(["git", "status", "--porcelain"], cwd=SANDBOX_DIR, capture_output=True, text=True)
        if status.stdout.strip():
            subprocess.run(["git", "commit", "-m", "Auto-update: New reefer arrivals logged"], cwd=SANDBOX_DIR, check=True)
            subprocess.run(["git", "push", "origin", "main"], cwd=SANDBOX_DIR, check=True)
            print("[+] Successfully synced updates to GitHub repository.")
        else:
            print("[*] GitHub is already up to date.")
    except Exception as e:
        print(f"[!] Git sync failed (verify remote credentials): {e}")

async def run_tracker():
    seen_cntrs = get_seen_containers()
    all_reefers = []

    async with async_playwright() as p:
        print(f"[*] Launching Chromium ({'Windows Sandbox' if IS_WINDOWS else 'GitHub Cloud Runner'})...")
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()

        print("[*] Checking DPD JNCH Advance Listing...")
        await page.goto("https://dpdjnch.com/ShippingLine/AdvanceListing.aspx", wait_until="networkidle")

        rows = await page.locator("table tr").all()
        for row in rows[1:]:
            cols = await row.locator("td").all_text_contents()
            if len(cols) < 5:
                continue

            line_name = cols[2].strip()
            vessel_name = cols[3].strip()
            voyage_no = cols[4].strip()

            if any(target in line_name.upper() for target in TARGET_LINES):
                action_btn = row.locator("td:last-child a, td:last-child input[type='submit']")
                if await action_btn.count() > 0:
                    try:
                        async with page.expect_download(timeout=12000) as download_info:
                            await action_btn.first.click()

                        download = await download_info.value
                        clean_vessel = re.sub(r'[^A-Za-z0-9_]', '_', vessel_name)
                        clean_line = re.sub(r'[^A-Za-z0-9_]', '_', line_name)[:25]
                        save_path = os.path.join(DOWNLOAD_DIR, f"{clean_line}_{clean_vessel}_{voyage_no}.xlsx")
                        await download.save_as(save_path)

                        found = parse_reefer_manifest(save_path, line_name, vessel_name, voyage_no)
                        if found is not None and not found.empty:
                            all_reefers.append(found)

                    except Exception:
                        pass

        await browser.close()

    if all_reefers:
        master_df = pd.concat(all_reefers, ignore_index=True)
        
        # Isolate brand-new reefer boxes never seen before
        new_reefers = master_df[~master_df["_CNTR"].isin(seen_cntrs)].copy()

        # Update CSV log
        if os.path.exists(MASTER_LOG_PATH):
            existing_df = pd.read_csv(MASTER_LOG_PATH)
            combined = pd.concat([existing_df, master_df]).drop_duplicates(subset=["_CNTR"])
            combined.to_csv(MASTER_LOG_PATH, index=False)
        else:
            master_df.to_csv(MASTER_LOG_PATH, index=False)

        if not new_reefers.empty:
            print(f"[***] ALERT: {len(new_reefers)} NEW Reefer(s) discovered!")
            send_gmail_alert(new_reefers)
            mark_containers_seen(new_reefers["_CNTR"].tolist())
            git_commit_and_push()
        else:
            print("[*] Reefers found in queue, but already alerted previously.")
    else:
        print("[-] Scan complete. No active reefers identified across current uploaded filings.")

if __name__ == "__main__":
    asyncio.run(run_tracker())