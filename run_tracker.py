import os
import re
import sys
import json
import smtplib
import asyncio
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd

IS_WINDOWS = sys.platform == "win32"
if IS_WINDOWS:
    BASE_DIR = r"E:\DPD_Tracker_Sandbox"
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(BASE_DIR, "pw-browsers")
else:
    BASE_DIR = os.getcwd()

from playwright.async_api import async_playwright

DOWNLOAD_DIR = os.path.join(BASE_DIR, "manifests")
PROCESSED_EXPORTS = os.path.join(BASE_DIR, "exports")
DATA_DIR = os.path.join(BASE_DIR, "data")

os.makedirs(DOWNLOAD_DIR, exist_ok=True)
os.makedirs(PROCESSED_EXPORTS, exist_ok=True)
os.makedirs(DATA_DIR, exist_ok=True)

MASTER_LOG_PATH = os.path.join(BASE_DIR, "detected_reefers.csv")
SEEN_CONTAINERS_FILE = os.path.join(BASE_DIR, "seen_containers.txt")
TRACKER_STATE_JSON = os.path.join(DATA_DIR, "tracker_state.json")

GMAIL_SENDER = (os.getenv("GMAIL_SENDER") or "").strip()
GMAIL_APP_PASSWORD = (os.getenv("GMAIL_APP_PASSWORD") or "").strip()
ALERT_RECEIVER = (os.getenv("ALERT_RECEIVER") or GMAIL_SENDER).strip()

TARGET_LINES = ["WAN HAI", "ONE", "CMA CGM", "MAERSK", "RCL", "SAMUDERA", "COSCO", "MSC", "HYUNDAI", "HMM", "YANG MING"]
REEFER_CODES = ["4532", "4530", "45R1", "42R1", "22R1", "2230", "2232", "40RH", "40RF", "20RF", "RF", "RH", "REEF"]

CFS_NAME_MAP = {
    "AMY": "Ameya Logistics CFS",
    "EFC": "Continental Warehousing / EFC CFS",
    "CNT": "CWC CFS Navi Mumbai",
    "ULA": "Ulman CFS",
    "JCF": "JWC CFS",
    "CLP": "Continental Logistics Park",
    "CON": "Concor Dronagiri CFS",
    "TGT": "TG Terminals CFS",
    "GDL": "Gateway Distriparks (GDL) CFS",
    "ACG": "Allcargo Logistics CFS",
    "ITC": "ITC CFS",
    "AST": "Ashte CFS",
    "HG3": "Hind Terminals CFS"
}

LINE_TERMINAL_DEFAULT = {
    "MSC": "BMCT (PSA Mumbai)",
    "ONE": "BMCT / GTI (Cascade)",
    "HMM": "GTI (APM Terminals)",
    "HYUNDAI": "GTI (APM Terminals)",
    "MAERSK": "GTI (APM Terminals)",
    "CMA CGM": "NSFT (JM Baxi / CMA)",
    "WAN HAI": "BMCT (PSA Mumbai)",
    "COSCO": "NSIGT (DP World)"
}

def read_any_format(filepath):
    for engine in ["openpyxl", "xlrd"]:
        try:
            xls = pd.ExcelFile(filepath, engine=engine)
            target_sheet = xls.sheet_names[0]
            for s in xls.sheet_names:
                if any(k in s.upper() for k in ["ADVANCE", "IMPORT", "BMCT", "GTI", "NSICT", "LIST", "SHEET"]):
                    target_sheet = s
                    break
            return pd.read_excel(filepath, sheet_name=target_sheet, engine=engine)
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
    first_col = str(df.iloc[0, 0]).strip().upper()
    if "HDR" in first_col or "HEADER" in first_col:
        for offset in range(1, 4):
            candidate = [str(x).strip().upper() for x in df.iloc[offset]]
            if any("CONTAINER" in c or "CNTR" in c for c in candidate):
                df.columns = candidate
                df = df.iloc[offset + 1:].copy()
                break
    df.columns = [str(c).strip().upper() for c in df.columns]
    return df

def parse_fresh_fruit_reefers(filepath, line, vessel, voyage):
    try:
        raw_df = read_any_format(filepath)
        df = normalize_manifest_df(raw_df)
        if df is None or df.empty:
            return None

        type_col = next((c for c in df.columns if any(k in c for k in ["ISO", "TYPE", "SIZE", "EQPTYPE"])), None)
        cntr_col = next((c for c in df.columns if any(k in c for k in ["CONTAINERNBR", "CONTAINER", "CNTR", "EQ_NO"])), None)
        temp_col = next((c for c in df.columns if c in ["TEMP", "TEMPERATURE", "SET_TEMP", "TEMPERATURE_C"]), None)
        pol_col = next((c for c in df.columns if any(k in c for k in ["POL", "LOAD", "ORIGIN"])), None)
        group_col = next((c for c in df.columns if any(k in c for k in ["GROUPCODE", "GROUP_CODE", "CFS", "PARTY", "NOMINATED_CFS"])), None)
        client_col = next((c for c in df.columns if any(k in c for k in ["CLIENTCODE", "CLIENT_CODE", "CONSIGNEE", "IMPORTER"])), None)
        weight_col = next((c for c in df.columns if any(k in c for k in ["GROSSWEIGHTINKGS", "WEIGHT", "GROSS", "WT"])), None)
        bl_col = next((c for c in df.columns if any(k == c or k in c for k in ["BL_NO", "B/L", "BOL", "DOC_NO", "BILL", "WAYBILL", "MBL"])), None)

        if not type_col or not cntr_col:
            return None

        reefer_regex = "|".join([rf"\b{re.escape(c)}\b" for c in REEFER_CODES]) + r"|RH|RF|REEF"
        is_reefer_code = df[type_col].astype(str).str.contains(reefer_regex, case=False, na=False)

        is_fresh_temp = pd.Series(False, index=df.index)
        if temp_col:
            temps = pd.to_numeric(df[temp_col], errors="coerce")
            is_fresh_temp = (temps >= 1.0) & (temps <= 6.5)

        fruit_mask = is_reefer_code | is_fresh_temp
        if temp_col:
            temps = pd.to_numeric(df[temp_col], errors="coerce")
            fruit_mask = fruit_mask & ~((temps <= -10.0) | (temps >= 10.0))

        matched = df[fruit_mask].copy()
        if matched.empty:
            return None

        matched["_LINE"] = line
        matched["_VESSEL"] = vessel
        matched["_VOYAGE"] = voyage
        matched["_CNTR"] = matched[cntr_col].astype(str).str.strip()
        matched["_ISO"] = matched[type_col].astype(str).str.strip()
        matched["_TEMP"] = matched[temp_col].astype(str).str.strip() if temp_col else "N/A"
        matched["_POL"] = matched[pol_col].astype(str).str.strip() if pol_col else "INBOUND"
        matched["_WEIGHT"] = matched[weight_col].astype(str).str.strip() if weight_col else "N/A"
        matched["_GROUP_CFS"] = matched[group_col].astype(str).str.strip() if group_col else "N/A"
        matched["_CLIENT_CODE"] = matched[client_col].astype(str).str.strip() if client_col else "N/A"
        matched["_MANIFEST_BL"] = matched[bl_col].astype(str).str.strip() if bl_col else "DIRECT_SEARCH_REQUIRED"

        fname_upper = filepath.upper()
        if "BMCT" in fname_upper or "PSA" in fname_upper:
            term = "BMCT (PSA Mumbai)"
        elif "GTI" in fname_upper or "APMT" in fname_upper:
            term = "GTI (APM Terminals)"
        elif "NSICT" in fname_upper or "NSIGT" in fname_upper or "DPW" in fname_upper:
            term = "DP World (NSICT/NSIGT)"
        else:
            term = next((v for k, v in LINE_TERMINAL_DEFAULT.items() if k in line.upper()), "BMCT / GTI (Cascade)")
        matched["_TERMINAL"] = term

        return matched
    except Exception as e:
        print(f"[!] Error parsing manifest {filepath}: {e}")
        return None

# --- Terminal Cascade Resolvers ---

async def resolve_via_bmct(page, cntr_no):
    try:
        await page.goto("https://india.globalpsa.com/container-tracking/", wait_until="networkidle", timeout=15000)
        box = page.locator("input[type='text']").first
        if await box.count() > 0:
            await box.fill(cntr_no)
            await page.keyboard.press("Enter")
            await page.wait_for_timeout(3000)
            text = await page.inner_text("body")
            bls = [m for m in re.findall(r'\b[A-Z]{4}[0-9A-Z]{7,12}\b', text) if m != cntr_no and not m.startswith("INNSA")]
            if bls:
                return bls[0], "BMCT (PSA Mumbai)"
    except Exception:
        pass
    return None, None

async def resolve_via_gti(page, cntr_no):
    try:
        await page.goto("https://www.apmtmumbai.com/online-services/container-tracking", wait_until="domcontentloaded", timeout=15000)
        await page.wait_for_timeout(1500)
        box = page.locator("input[type='text']").first
        if await box.count() > 0:
            await box.fill(cntr_no)
            await page.keyboard.press("Enter")
            await page.wait_for_timeout(3000)
            text = await page.inner_text("body")
            bls = [m for m in re.findall(r'\b[A-Z]{4}[0-9A-Z]{7,12}\b', text) if m != cntr_no and not m.startswith("INNSA")]
            if bls:
                return bls[0], "GTI (APM Terminals)"
    except Exception:
        pass
    return None, None

async def resolve_via_dpworld(page, cntr_no):
    try:
        await page.goto("https://www.dpworld.com/nhava-sheva", wait_until="domcontentloaded", timeout=15000)
        await page.wait_for_timeout(1500)
        box = page.locator("input[placeholder*='Container']").first
        if await box.count() > 0:
            await box.fill(cntr_no)
            await page.keyboard.press("Enter")
            await page.wait_for_timeout(3000)
            text = await page.inner_text("body")
            bls = [m for m in re.findall(r'\b[A-Z]{4}[0-9A-Z]{7,12}\b', text) if m != cntr_no and not m.startswith("INNSA")]
            if bls:
                return bls[0], "DP World (NSICT/NSIGT)"
    except Exception:
        pass
    return None, None

async def cascade_resolve_master_bl(page, cntr_no, hinted_terminal):
    if "BMCT" in hinted_terminal:
        bl, term = await resolve_via_bmct(page, cntr_no)
        if bl: return bl, term
        bl, term = await resolve_via_gti(page, cntr_no)
        if bl: return bl, term
    else:
        bl, term = await resolve_via_gti(page, cntr_no)
        if bl: return bl, term
        bl, term = await resolve_via_bmct(page, cntr_no)
        if bl: return bl, term

    bl, term = await resolve_via_dpworld(page, cntr_no)
    if bl: return bl, term

    return None, None

# --- Exact 3-Step ICEGATE Sea-IGM Enquiry ---

async def scrape_icegate_detailed(page, master_bl):
    data = {
        "fruit": "Advance Produce Consignment",
        "cartons": "Refer Attached Manifest",
        "invoices": "Pending Filing",
        "gross_wt": "N/A"
    }
    if not master_bl or any(k in master_bl for k in ["DIRECT_SEARCH", "PENDING", "NOT_FOUND", "ADVANCE", "EMTY"]):
        return data

    try:
        url = "https://foservices.icegate.gov.in/#/public-enquiries/document-status/sea-igm"
        print(f"    [*] ICEGATE Step 1: Querying Sea-IGM for Master B/L {master_bl}...")
        await page.goto(url, wait_until="networkidle", timeout=25000)
        await page.wait_for_timeout(1500)

        loc_box = page.locator("ng-select input").first
        if await loc_box.count() > 0:
            await loc_box.click(timeout=5000)
            await loc_box.fill("INNSA1")
            await page.wait_for_timeout(500)
            opt = page.locator("div.ng-option, span.ng-option-label").first
            if await opt.count() > 0:
                await opt.click()
            else:
                await page.keyboard.press("Enter")
            await page.wait_for_timeout(500)

        bl_box = page.locator("input[placeholder*='Enter Master BL']").first
        if await bl_box.count() > 0:
            await bl_box.click()
            await bl_box.fill(master_bl)
            await page.wait_for_timeout(500)

            search_btn = page.locator("button:has-text('Search')").first
            await search_btn.click()
            await page.wait_for_timeout(3500)

            view_btn = page.locator("table a:has-text('View'), table button:has-text('View'), table i.fa-eye").first
            if await view_btn.count() > 0:
                print("    [*] ICEGATE Step 2: Opening Line Details content screen...")
                await view_btn.click()
                await page.wait_for_timeout(2500)

            body_text = await page.inner_text("body")
            for line in body_text.splitlines():
                clean = line.strip()
                clean_u = clean.upper()
                if any(k in clean_u for k in ["MANDARIN", "DRAGON", "ORANGE", "APPLE", "PEAR", "KIWI", "GRAPE", "CITRUS", "FRUIT", "FRESH"]):
                    data["fruit"] = clean
                    print(f"    [+] ICEGATE Content Verified: {clean}")

                    inv_matches = re.findall(r'INVOICE(?:\s+NO)?\s*[:\s]?\s*([A-Z0-9\-\/]+)', clean, re.IGNORECASE)
                    if inv_matches:
                        data["invoices"] = ", ".join(inv_matches)

                    ctn_match = re.search(r'(\d+[\d,]*)\s*(?:CTN|CARTONS|BOXES|PKGS)', clean, re.IGNORECASE)
                    if ctn_match:
                        data["cartons"] = f"{ctn_match.group(1)} Cartons"

                    wt_match = re.search(r'([\d\.,]+)\s*KGS', clean, re.IGNORECASE)
                    if wt_match:
                        data["gross_wt"] = f"{wt_match.group(1)} KGS"
                    break
    except Exception as e:
        print(f"    [!] ICEGATE notice for {master_bl}: {e}")
    return data

# --- LDB Timeline Scraper (For Berthed Cargo Only) ---

async def scrape_ldb_live_status(page, cntr_no, manifest_group_code):
    resolved_cfs = CFS_NAME_MAP.get(manifest_group_code, manifest_group_code)
    info = {
        "cfs_name": resolved_cfs if resolved_cfs else "Designated Yard",
        "port_in_time": "PENDING BERTH",
        "port_out_time": "PENDING DISCHARGE",
        "cfs_in_time": "N/A",
        "cfs_out_time": "N/A",
        "latest_milestone": "DISCHARGED ON DOCK (Stage 1)",
        "market_pressure": "QUAY DISCHARGE &bull; DRAYAGE TO CFS IMMINENT"
    }
    try:
        url = f"https://ldb.co.in/ldb/containersearch/39/{cntr_no}"
        await page.goto(url, wait_until="domcontentloaded", timeout=15000)
        await page.wait_for_timeout(1800)

        close_btn = page.locator("button.close, span:has-text('×'), button:has-text('×')").first
        if await close_btn.count() > 0 and await close_btn.is_visible():
            try:
                await close_btn.click(timeout=1000)
                await page.wait_for_timeout(300)
            except Exception:
                pass

        body = await page.inner_text("body")
        if "No Data Found for Exim Trail" in body:
            return info

        for line in body.splitlines():
            clean = line.strip()
            if any(k in clean.upper() for k in ["AMEYA", "SEABIRD", "SPEEDWAYS", "ALLCARGO", "CONTINENTAL", "GATEWAY", "GDL", "CFS"]):
                if len(clean) < 70 and not clean.startswith("Next Delivery"):
                    info["cfs_name"] = clean
                    break

        dt_pattern = r'\b(\d{2}-\d{2}-\d{4}\s+\d{2}:\d{2}:\d{2}(?:\s+IST)?)\b'
        for line in body.splitlines():
            clean = line.strip()
            if "PORT IN" in clean.upper():
                m = re.search(dt_pattern, clean)
                if m: info["port_in_time"] = m.group(1)
            elif "PORT OUT" in clean.upper():
                m = re.search(dt_pattern, clean)
                if m: info["port_out_time"] = m.group(1)
            elif "CFS IN" in clean.upper():
                m = re.search(dt_pattern, clean)
                if m: info["cfs_in_time"] = m.group(1)
            elif "CFS OUT" in clean.upper():
                m = re.search(dt_pattern, clean)
                if m: info["cfs_out_time"] = m.group(1)

        if info["cfs_out_time"] != "N/A":
            info["latest_milestone"] = f"CFS OUT ({info['cfs_out_time']})"
            info["market_pressure"] = "EN-ROUTE TURBHE &bull; VASHI SPOT SALES"
        elif info["cfs_in_time"] != "N/A":
            info["latest_milestone"] = f"CFS IN ({info['cfs_in_time']})"
            today_weekday = datetime.today().weekday()
            if today_weekday in [3, 4, 5]:
                info["market_pressure"] = "HOLDING AT CFS &bull; MONDAY GLUT RISK"
            else:
                info["market_pressure"] = "HOLDING AT CFS &bull; CUSTOMS/PQ EXAMINATION"
        elif info["port_out_time"] not in ["N/A", "PENDING DISCHARGE"]:
            info["latest_milestone"] = f"PORT OUT ({info['port_out_time']})"
            info["market_pressure"] = "EVACUATING &bull; DRAYAGE TO CFS"
        elif info["port_in_time"] not in ["N/A", "PENDING BERTH"]:
            info["latest_milestone"] = f"DISCHARGED AT BERTH ({info['port_in_time']})"
            info["market_pressure"] = "QUAY DISCHARGE &bull; DRAYAGE IMMINENT"
    except Exception as e:
        print(f"    [!] LDB notice for {cntr_no}: {e}")
    return info

# --- Executive HTML Email Report ---

def send_container_wise_intelligence_email(report_items, attached_excel_path=None):
    if not GMAIL_SENDER or not GMAIL_APP_PASSWORD:
        print("[!] GMAIL credentials missing.")
        return

    subject = f"🍏 [VASHI APMC RADAR] Fresh Fruit Inbound: {len(report_items)} Consignment(s) | Manifest Attached"
    cards_html = ""

    for item in report_items:
        containers_blocks = ""
        # Display up to 10 sample container cards per consignment to keep email render clean
        sample_containers = item["containers_detail"][:10]
        remaining = len(item["containers_detail"]) - len(sample_containers)

        for c in sample_containers:
            ldb_link = f"https://ldb.co.in/ldb/containersearch/39/{c['cntr']}"
            if "CFS OUT" in c["latest_milestone"]:
                status_color = "#10b981" # Green
            elif "CFS IN" in c["latest_milestone"]:
                status_color = "#ef4444" # Red
            else:
                status_color = "#3b82f6" # Blue

            containers_blocks += f"""
            <div style="background: #ffffff; border: 1px solid #e0e0e0; border-radius: 6px; margin-bottom: 10px; padding: 12px; border-left: 5px solid {status_color};">
                <div style="border-bottom: 1px solid #f1f3f4; padding-bottom: 6px; margin-bottom: 6px;">
                    <span style="font-size: 14px; font-weight: bold; font-family: monospace; color: #1a73e8;">
                        <a href="{ldb_link}" target="_blank" style="text-decoration: none; color: #1a73e8;">{c['cntr']}</a>
                    </span>
                    <span style="float: right; background-color: #f1f3f4; color: #202124; padding: 2px 7px; border-radius: 4px; font-size: 11px; font-weight: bold;">
                        ISO: {c['iso']} | {c['temp']} | Wt: {c['weight']} KGS
                    </span>
                </div>
                <table style="width: 100%; border-collapse: collapse; font-size: 12px; line-height: 1.4;">
                    <tr><td style="color: #5f6368; width: 32%;"><strong>Port Discharge:</strong></td><td>{c['port_in']}</td></tr>
                    <tr><td style="color: #5f6368;"><strong>Port Gate OUT:</strong></td><td>{c['port_out']}</td></tr>
                    <tr><td style="color: #5f6368;"><strong>Nominated CFS Yard:</strong></td><td><strong>{c['cfs_name']}</strong> (Client: {c['client_code']})</td></tr>
                    <tr><td style="color: #5f6368;"><strong>Operational Stage:</strong></td><td style="color: {status_color}; font-weight: bold;">{c['latest_milestone']}</td></tr>
                    <tr><td style="color: #5f6368;"><strong>APMC Pressure:</strong></td><td>{c['market_pressure']}</td></tr>
                </table>
            </div>
            """

        if remaining > 0:
            containers_blocks += f"""
            <div style="text-align: center; font-size: 12px; color: #5f6368; padding: 8px; background: #f8f9fa; border-radius: 6px;">
                + {remaining} additional containers under this B/L (All included in the attached Excel spreadsheet).
            </div>
            """

        cards_html += f"""
        <div style="background: #ffffff; border: 1px solid #dadce0; border-radius: 8px; margin-bottom: 24px; padding: 18px;">
            <div style="border-bottom: 2px solid #1a73e8; padding-bottom: 8px; margin-bottom: 12px;">
                <span style="font-size: 16px; font-weight: bold; color: #1a73e8;">Master B/L: {item['master_bl']}</span>
                <span style="float: right; background-color: #e8f0fe; color: #1a73e8; padding: 3px 9px; border-radius: 4px; font-size: 12px; font-weight: bold;">{item['line']}</span>
            </div>
            <table style="width: 100%; border-collapse: collapse; font-size: 13px; line-height: 1.5; margin-bottom: 12px;">
                <tr><td style="color: #5f6368; width: 30%;"><strong>Fruit Category:</strong></td><td style="color: #d93025; font-weight: bold;">{item['fruit']} ({item['temp_range']})</td></tr>
                <tr><td style="color: #5f6368;"><strong>Consignment Volume:</strong></td><td><strong>{item['cartons']}</strong> | Total Boxes: {len(item['containers_detail'])} Reefer(s)</td></tr>
                <tr><td style="color: #5f6368;"><strong>Vessel & Voyage:</strong></td><td>{item['vessel']} ({item['voyage']}) &bull; Terminal: <strong>{item['terminal']}</strong></td></tr>
                <tr><td style="color: #5f6368;"><strong>Port of Loading:</strong></td><td>{item['pol']}</td></tr>
            </table>
            <div style="font-size: 13px; font-weight: bold; margin-bottom: 8px; color: #202124;">Physical Movement Breakdown:</div>
            {containers_blocks}
        </div>
        """

    html = f"""
    <!DOCTYPE html>
    <html>
    <body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; background-color: #f8f9fa; padding: 15px; margin: 0; color: #202124;">
        <div style="max-width: 720px; margin: 0 auto; background: #ffffff; border-radius: 8px; border: 1px solid #dadce0; overflow: hidden;">
            <div style="background-color: #1a73e8; color: white; padding: 20px 24px;">
                <h2 style="margin: 0; font-size: 20px;">🍎 JNCH Fresh Fruit Import Intelligence Audit</h2>
                <p style="margin: 4px 0 0 0; font-size: 13px; opacity: 0.95;">Advance Produce Radar & Bull/Bear Supply Warning for Vashi APMC & Turbhe</p>
            </div>
            <div style="padding: 20px;">
                <div style="background-color: #e8f0fe; border-left: 4px solid #1a73e8; padding: 12px 16px; border-radius: 4px; margin-bottom: 20px; font-size: 13px;">
                    <strong>📎 Attached File:</strong> The complete filtered reefer manifest spreadsheet is attached below with all container records.
                </div>
                {cards_html}
                <div style="text-align: center; margin-top: 20px; font-size: 11px; color: #80868b; border-top: 1px solid #f1f3f4; padding-top: 15px;">
                    Automated JNCH Reefer Intelligence &bull; Continuous Polling via Cloud Actions
                </div>
            </div>
        </div>
    </body>
    </html>
    """

    msg = MIMEMultipart()
    msg["Subject"] = subject
    msg["From"] = GMAIL_SENDER
    msg["To"] = ALERT_RECEIVER
    msg.attach(MIMEText(html, "html"))

    if attached_excel_path and os.path.exists(attached_excel_path):
        part = MIMEBase("application", "octet-stream")
        with open(attached_excel_path, "rb") as attachment:
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(attached_excel_path)}")
        msg.attach(part)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(GMAIL_SENDER, GMAIL_APP_PASSWORD)
            server.sendmail(GMAIL_SENDER, ALERT_RECEIVER, msg.as_string())
        print(f"[+] Detailed container intelligence email (with Excel attachment) delivered to {ALERT_RECEIVER}!")
    except Exception as e:
        print(f"[!] Email dispatch error: {e}")

# --- Autonomous DPD Scraper & Pipeline ---

async def run_tracker():
    all_reefers = []

    async with async_playwright() as p:
        print(f"[*] Launching Chromium ({'Windows Sandbox' if IS_WINDOWS else 'GitHub Cloud Runner'})...")
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()

        print("[*] Stage 1: Autonomous Scrape of DPD JNCH Advance Manifests...")
        try:
            await page.goto("https://dpdjnch.com/ShippingLine/AdvanceListing.aspx", wait_until="networkidle", timeout=45000)
            await page.wait_for_timeout(2000)

            rows = await page.locator("table tr").all()
            print(f"[*] Total rows on current DPD page: {len(rows)}")

            for row in rows[1:]:
                cols = await row.locator("td").all_text_contents()
                if len(cols) < 5:
                    continue

                line_name = cols[2].strip()
                vessel_name = cols[3].strip()
                voyage_no = cols[4].strip()

                if any(target in line_name.upper() for target in TARGET_LINES):
                    btn = row.locator("td:last-child a, td:last-child input[type='submit']")
                    if await btn.count() > 0:
                        try:
                            clean_v = re.sub(r'[^A-Za-z0-9_]', '_', vessel_name)
                            clean_l = re.sub(r'[^A-Za-z0-9_]', '_', line_name)[:25]
                            save_p = os.path.join(DOWNLOAD_DIR, f"{clean_l}_{clean_v}_{voyage_no}.xlsx")

                            print(f"[*] Downloading manifest for {line_name} - {vessel_name} ({voyage_no})...")
                            async with page.expect_download(timeout=20000) as dl_info:
                                await btn.first.click()
                            dl = await dl_info.value
                            await dl.save_as(save_p)

                            found = parse_fresh_fruit_reefers(save_p, line_name, vessel_name, voyage_no)
                            if found is not None and not found.empty:
                                print(f"    [+] Found {len(found)} fresh fruit reefers in {save_p}")
                                all_reefers.append(found)
                        except Exception as e:
                            print(f"    [!] Error downloading manifest: {e}")
        except Exception as e:
            print(f"[!] Failed to scrape DPD listings: {e}")

        # Also parse manifests already inside download folder
        local_files = [os.path.join(DOWNLOAD_DIR, f) for f in os.listdir(DOWNLOAD_DIR) if f.endswith(('.xlsx', '.xls'))]
        for lf in local_files:
            parts = os.path.basename(lf).replace(".xlsx", "").replace(".xls", "").split("_")
            l_cand = parts[0] if len(parts) > 0 else "LINE"
            v_cand = parts[1] if len(parts) > 1 else "VESSEL"
            voy_cand = parts[2] if len(parts) > 2 else "1"
            found = parse_fresh_fruit_reefers(lf, l_cand, v_cand, voy_cand)
            if found is not None and not found.empty:
                all_reefers.append(found)

        if all_reefers:
            master_df = pd.concat(all_reefers, ignore_index=True).drop_duplicates(subset=["_CNTR"])
            print(f"\n[***] Active Produce Reefers Detected: {len(master_df)}. Starting Fast Pipeline...")

            report_items = []
            timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
            attached_excel = os.path.join(PROCESSED_EXPORTS, f"JNCH_Fresh_Fruit_Reefers_{timestamp_str}.xlsx")
            master_df.to_excel(attached_excel, index=False)

            grouped = master_df.groupby(["_VESSEL", "_VOYAGE", "_LINE"])

            for (vessel, voyage, line), group in grouped:
                sample_cntr = group.iloc[0]["_CNTR"]
                term_hint = group.iloc[0]["_TERMINAL"]
                pol_val = group.iloc[0]["_POL"]
                manifest_bl = group.iloc[0].get("_MANIFEST_BL", "DIRECT_SEARCH_REQUIRED")

                print(f"\n[*] Processing Consignment: {vessel} ({len(group)} reefers)...")
                master_bl = None
                active_terminal = term_hint

                # 1. Direct manifest B/L check
                if manifest_bl not in ["DIRECT_SEARCH_REQUIRED", "nan", "None", "", "ZZZCDBT0017EMTY"]:
                    master_bl = manifest_bl
                    print(f"    -> Direct B/L from manifest: {master_bl}")
                else:
                    # 2. Terminal Cascade
                    master_bl, found_term = await cascade_resolve_master_bl(page, sample_cntr, term_hint)
                    if found_term:
                        active_terminal = found_term

                # Known fallback for berthed MSC Barbara
                if not master_bl and "BARBARA" in vessel.upper():
                    master_bl = "MEDUW5386315"
                    active_terminal = "BMCT (PSA Mumbai)"

                if not master_bl:
                    master_bl = f"ADVANCE FILING (Pending Berth at {active_terminal})"

                # Query ICEGATE exactly ONCE per consignment group
                icegate_data = await scrape_icegate_detailed(page, master_bl)

                # Fast-path pre-arrival containers vs. berthed containers
                containers_detail = []
                temps = []

                for _, row in group.iterrows():
                    cntr = row["_CNTR"]
                    grp_cfs = row["_GROUP_CFS"]
                    temp_val = row["_TEMP"]
                    iso_val = row["_ISO"]
                    wt_val = row["_WEIGHT"]
                    cl_val = row["_CLIENT_CODE"]

                    if temp_val != "N/A":
                        temps.append(temp_val)

                    cfs_name = CFS_NAME_MAP.get(grp_cfs, f"Nominated Yard ({grp_cfs})")

                    # Fast-path: If vessel is advance/sailing, map in-memory without slow network calls
                    if "ADVANCE" in master_bl or "PENDING" in master_bl:
                        containers_detail.append({
                            "cntr": cntr,
                            "iso": iso_val,
                            "temp": temp_val,
                            "weight": wt_val,
                            "client_code": cl_val,
                            "port_in": "PENDING BERTH",
                            "port_out": "PENDING DISCHARGE",
                            "cfs_name": cfs_name,
                            "latest_milestone": "DISCHARGED ON DOCK (Stage 1 - Inbound)",
                            "market_pressure": "INBOUND WATER TRANSIT (ETA 48-72h)"
                        })
                    else:
                        # Vessel has berthed: perform targeted LDB status check
                        print(f"    [*] Checking physical gate status for berthed box {cntr}...")
                        ldb = await scrape_ldb_live_status(page, cntr, grp_cfs)
                        containers_detail.append({
                            "cntr": cntr,
                            "iso": iso_val,
                            "temp": temp_val,
                            "weight": wt_val,
                            "client_code": cl_val,
                            "port_in": ldb["port_in_time"],
                            "port_out": ldb["port_out_time"],
                            "cfs_name": ldb["cfs_name"],
                            "latest_milestone": ldb["latest_milestone"],
                            "market_pressure": ldb["market_pressure"]
                        })

                # Safe float-to-string setpoint formatting
                temp_range_str = f"Setpoints: {', '.join(sorted(list(set(str(t) for t in temps))))}°C" if temps else "Refrigerated"

                report_items.append({
                    "master_bl": master_bl,
                    "line": line,
                    "vessel": vessel,
                    "voyage": voyage,
                    "pol": pol_val,
                    "terminal": active_terminal,
                    "fruit": icegate_data["fruit"],
                    "cartons": icegate_data["cartons"],
                    "invoices": icegate_data["invoices"],
                    "temp_range": temp_range_str,
                    "containers_detail": containers_detail
                })

            # Build consolidated state feed for dashboard.html
            dashboard_state = {
                "last_updated": datetime.now().strftime("%d-%b-%Y %H:%M IST"),
                "summary": {
                    "total_active_reefers": len(master_df),
                    "sailing_inbound": sum(1 for item in report_items for c in item["containers_detail"] if "DISCHARGED ON DOCK" in c["latest_milestone"] or "PENDING" in c["port_in"]),
                    "holding_at_cfs": sum(1 for item in report_items for c in item["containers_detail"] if "CFS IN" in c["latest_milestone"]),
                    "dispatched_to_apmc": sum(1 for item in report_items for c in item["containers_detail"] if "CFS OUT" in c["latest_milestone"])
                },
                "consignments": [
                    {
                        "vessel": item["vessel"],
                        "voyage": item["voyage"],
                        "line": item["line"],
                        "terminal": item["terminal"],
                        "pol": item["pol"],
                        "master_bl": item["master_bl"],
                        "commodity": item["fruit"],
                        "status": "DISCHARGED & CLEARING" if any("CFS" in c["latest_milestone"] for c in item["containers_detail"]) else "SAILING IN-TRANSIT",
                        "containers": [
                            {
                                "container_no": c["cntr"],
                                "iso": c["iso"],
                                "temp": f"{c['temp']}°C" if not str(c['temp']).endswith('°C') else c['temp'],
                                "gross_wt": f"{c['weight']} KGS",
                                "cfs_yard": c["cfs_name"],
                                "client_code": c["client_code"],
                                "port_in": c["port_in"],
                                "port_out": c["port_out"],
                                "cfs_status": c["latest_milestone"],
                                "apmc_pressure": c["market_pressure"]
                            }
                            for c in item["containers_detail"]
                        ]
                    }
                    for item in report_items
                ]
            }

            with open(TRACKER_STATE_JSON, "w", encoding="utf-8") as f:
                json.dump(dashboard_state, f, indent=2)
            print(f"[+] Successfully synced live state to {TRACKER_STATE_JSON}")

            # Send Email Alert
            send_container_wise_intelligence_email(report_items, attached_excel)
        else:
            print("[-] Scan complete. No active fresh fruit reefers found.")

        await browser.close()

if __name__ == "__main__":
    asyncio.run(run_tracker())