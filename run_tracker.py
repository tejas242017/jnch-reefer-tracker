import os
import re
import sys
import smtplib
import asyncio
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd

IS_WINDOWS = sys.platform == "win32"
if IS_WINDOWS:
    BASE_DIR = r"E:\DPD_Tracker_Sandbox"
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(BASE_DIR, "pw-browsers")
else:
    BASE_DIR = os.getcwd()

from playwright.async_api import async_playwright

DOWNLOAD_DIR = os.path.join(BASE_DIR, "manifests")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

MASTER_LOG_PATH = os.path.join(BASE_DIR, "detected_reefers.csv")
SEEN_CONTAINERS_FILE = os.path.join(BASE_DIR, "seen_containers.txt")

GMAIL_SENDER = (os.getenv("GMAIL_SENDER") or "").strip()
GMAIL_APP_PASSWORD = (os.getenv("GMAIL_APP_PASSWORD") or "").strip()
ALERT_RECEIVER = (os.getenv("ALERT_RECEIVER") or GMAIL_SENDER).strip()

TARGET_LINES = ["WAN HAI", "ONE", "CMA CGM", "MAERSK", "RCL", "SAMUDERA", "COSCO", "MSC", "HYUNDAI", "HMM"]
REEFER_CODES = ["4532", "45R1", "42R1", "22R1", "40RH", "40RF", "20RF", "RF", "RH", "REEF"]

CFS_NAME_MAP = {
    "AMY": "Ameya Logistics CFS",
    "EFC": "Continental Warehousing / EFC CFS",
    "CNT": "CWC CFS Navi Mumbai",
    "ULA": "Ulman CFS",
    "JCF": "JWC CFS",
    "CLP": "Continental Logistics Park",
    "CON": "Concor Dronagiri CFS",
    "TGT": "TG Terminals CFS"
}

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
    for engine in ["openpyxl", "xlrd"]:
        try:
            xls = pd.ExcelFile(filepath, engine=engine)
            target_sheet = xls.sheet_names[0]
            for s in xls.sheet_names:
                if any(k in s.upper() for k in ["ADVANCE", "IMPORT", "BMCT", "GTI", "NSICT"]):
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
        cntr_col = next((c for c in df.columns if any(k in c for k in ["CONTAINER", "CNTR", "EQ_NO"])), None)
        temp_col = next((c for c in df.columns if c in ["TEMP", "TEMPERATURE", "SET_TEMP"]), None)
        pol_col = next((c for c in df.columns if any(k in c for k in ["POL", "LOAD", "ORIGIN"])), None)
        group_col = next((c for c in df.columns if any(k in c for k in ["GROUPCODE", "GROUP_CODE", "CFS", "PARTY"])), None)
        weight_col = next((c for c in df.columns if any(k in c for k in ["WEIGHT", "GROSS", "WT"])), None)

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
            fruit_mask = fruit_mask & ~((temps <= -10.0) | (temps >= 15.0))

        matched = df[fruit_mask].copy()
        if matched.empty:
            return None

        matched["_LINE"] = line
        matched["_VESSEL"] = vessel
        matched["_VOYAGE"] = voyage
        matched["_CNTR"] = matched[cntr_col].astype(str).str.strip()
        matched["_ISO"] = matched[type_col].astype(str).str.strip()
        matched["_TEMP"] = matched[temp_col].astype(str).str.strip() if temp_col else "N/A"
        matched["_POL"] = matched[pol_col].astype(str).str.strip() if pol_col else "N/A"
        matched["_WEIGHT"] = matched[weight_col].astype(str).str.strip() if weight_col else "N/A"
        matched["_GROUP_CFS"] = matched[group_col].astype(str).str.strip() if group_col else "N/A"

        fname_upper = filepath.upper()
        if "BMCT" in fname_upper or "PSA" in fname_upper:
            term = "BMCT"
        elif "GTI" in fname_upper or "APMT" in fname_upper:
            term = "GTI"
        elif "NSICT" in fname_upper or "NSIGT" in fname_upper or "DPW" in fname_upper:
            term = "DPW"
        else:
            term = "CASCADE"
        matched["_TERMINAL"] = term

        return matched
    except Exception as e:
        print(f"[!] Error parsing manifest {filepath}: {e}")
        return None

# --- Terminal Cascade Resolvers ---

async def resolve_via_bmct(page, cntr_no):
    try:
        await page.goto("https://india.globalpsa.com/container-tracking/", wait_until="networkidle", timeout=25000)
        box = page.locator("input[type='text']").first
        if await box.count() > 0:
            await box.fill(cntr_no)
            await page.keyboard.press("Enter")
            await page.wait_for_timeout(3000)
            text = await page.inner_text("body")
            bls = [m for m in re.findall(r'\b[A-Z]{4}[0-9A-Z]{7,12}\b', text) if m != cntr_no]
            if bls:
                return bls[0], "BMCT (PSA)"
    except Exception:
        pass
    return None, None

async def resolve_via_gti(page, cntr_no):
    try:
        await page.goto("https://www.apmtmumbai.com/online-services/container-tracking", wait_until="domcontentloaded", timeout=20000)
        await page.wait_for_timeout(2000)
        box = page.locator("input[type='text']").first
        if await box.count() > 0:
            await box.fill(cntr_no)
            await page.keyboard.press("Enter")
            await page.wait_for_timeout(3500)
            text = await page.inner_text("body")
            bls = [m for m in re.findall(r'\b[A-Z]{4}[0-9A-Z]{7,12}\b', text) if m != cntr_no]
            if bls:
                return bls[0], "GTI (APMT)"
    except Exception:
        pass
    return None, None

async def resolve_via_dpworld(page, cntr_no):
    try:
        await page.goto("https://www.dpworld.com/nhava-sheva", wait_until="domcontentloaded", timeout=20000)
        await page.wait_for_timeout(2000)
        box = page.locator("input[placeholder*='Container']").first
        if await box.count() > 0:
            await box.fill(cntr_no)
            await page.keyboard.press("Enter")
            await page.wait_for_timeout(3500)
            text = await page.inner_text("body")
            bls = [m for m in re.findall(r'\b[A-Z]{4}[0-9A-Z]{7,12}\b', text) if m != cntr_no]
            if bls:
                return bls[0], "DP World (NSICT/NSIGT)"
    except Exception:
        pass
    return None, None

async def cascade_resolve_master_bl(page, cntr_no, hinted_terminal):
    if hinted_terminal == "BMCT":
        bl, term = await resolve_via_bmct(page, cntr_no)
        if bl: return bl, term
    elif hinted_terminal == "GTI":
        bl, term = await resolve_via_gti(page, cntr_no)
        if bl: return bl, term
    elif hinted_terminal == "DPW":
        bl, term = await resolve_via_dpworld(page, cntr_no)
        if bl: return bl, term

    for func in [resolve_via_bmct, resolve_via_gti, resolve_via_dpworld]:
        bl, term = await func(page, cntr_no)
        if bl: return bl, term

    return "UNRESOLVED", "Unknown Terminal"

# --- ICEGATE Cargo & Invoice Scraper ---

async def scrape_icegate(page, master_bl):
    data = {
        "fruit": "Perishable Fruit Consignment",
        "cartons": "N/A",
        "invoices": "N/A",
        "gross_wt": "N/A",
        "sister_containers": []
    }
    if not master_bl or master_bl == "UNRESOLVED":
        return data

    try:
        url = "https://foservices.icegate.gov.in/#/public-enquiries/document-status/sea-igm"
        await page.goto(url, wait_until="networkidle", timeout=35000)
        await page.wait_for_timeout(1500)

        loc_box = page.locator("ng-select input").first
        await loc_box.click()
        await loc_box.fill("INNSA1")
        await page.wait_for_timeout(600)
        opt = page.locator("div.ng-option, span.ng-option-label").first
        if await opt.count() > 0:
            await opt.click()
        else:
            await page.keyboard.press("Enter")
        await page.wait_for_timeout(600)

        bl_box = page.locator("input[placeholder*='Enter Master BL']").first
        await bl_box.click()
        await bl_box.fill(master_bl)
        await page.wait_for_timeout(600)

        await page.locator("button:has-text('Search')").first.click()
        await page.wait_for_timeout(4000)

        summary_text = await page.inner_text("table")
        for line in summary_text.splitlines():
            line_u = line.strip().upper()
            if any(k in line_u for k in ["MANDARIN", "DRAGON", "ORANGE", "APPLE", "PEAR", "KIWI", "GRAPE", "CITRUS", "FRUIT"]):
                data["fruit"] = line.strip()
                inv_matches = re.findall(r'INVOICE NO\s+([A-Z0-9]+)', line, re.IGNORECASE)
                if inv_matches:
                    data["invoices"] = ", ".join(inv_matches)
                ctn_match = re.search(r'(\d+)\s+CTN', line, re.IGNORECASE)
                if ctn_match:
                    data["cartons"] = f"{ctn_match.group(1)} Cartons"
                wt_match = re.search(r'([\d\.]+)\s+KGS', line, re.IGNORECASE)
                if wt_match:
                    data["gross_wt"] = f"{wt_match.group(1)} KGS"
                break

        view_btn = page.locator("table a:has-text('View'), table button:has-text('View')").first
        if await view_btn.count() > 0:
            await view_btn.click()
            await page.wait_for_timeout(2500)
            c_btn = page.locator("a:has-text('Container Details'), button:has-text('Container Details')").first
            if await c_btn.count() > 0:
                await c_btn.click()
                await page.wait_for_timeout(2500)
            modal_text = await page.inner_text("body")
            data["sister_containers"] = sorted(list(set(re.findall(r'\b[A-Z]{4}\d{7}\b', modal_text))))
    except Exception as e:
        print(f"    [!] ICEGATE error for {master_bl}: {e}")
    return data

# --- LDB Live DOM Node Scraper ---

async def scrape_ldb_live_status(page, cntr_no, manifest_group_code):
    info = {
        "cfs_name": CFS_NAME_MAP.get(manifest_group_code, manifest_group_code),
        "port_in_time": "N/A",
        "port_out_time": "N/A",
        "cfs_in_time": "N/A",
        "cfs_out_time": "N/A",
        "latest_milestone": "En-route to Yard",
        "market_pressure": "HOLDING AT CFS"
    }
    try:
        url = f"https://ldb.co.in/ldb/containersearch/39/{cntr_no}"
        await page.goto(url, wait_until="domcontentloaded", timeout=25000)
        await page.wait_for_timeout(2500)

        close_btn = page.locator("button.close, span:has-text('×'), button:has-text('×')").first
        if await close_btn.count() > 0 and await close_btn.is_visible():
            try:
                await close_btn.click(timeout=1500)
                await page.wait_for_timeout(500)
            except Exception:
                pass

        body = await page.inner_text("body")

        for line in body.splitlines():
            clean = line.strip()
            if any(k in clean.upper() for k in ["AMEYA", "SEABIRD", "SPEEDWAYS", "ALLCARGO", "CONTINENTAL", "CFS"]):
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
            info["market_pressure"] = "APMC ARRIVED / DISPATCHED"
        elif info["cfs_in_time"] != "N/A":
            info["latest_milestone"] = f"CFS IN ({info['cfs_in_time']})"
            today_weekday = datetime.today().weekday()
            if today_weekday in [3, 4, 5]:
                info["market_pressure"] = "HIGH MONDAY GLUT RISK (Holding at CFS)"
            else:
                info["market_pressure"] = "HOLDING AT CFS (Customs / PQ)"
        elif info["port_out_time"] != "N/A":
            info["latest_milestone"] = f"PORT OUT ({info['port_out_time']}) -> Drayage to CFS"
            info["market_pressure"] = "EVACUATING TO CFS"
        elif info["port_in_time"] != "N/A":
            info["latest_milestone"] = f"DISCHARGED AT BERTH ({info['port_in_time']})"
            info["market_pressure"] = "PORT TERMINAL DISCHARGE"
    except Exception as e:
        print(f"    [!] LDB live parse error for {cntr_no}: {e}")
    return info

# --- Executive HTML Email Report ---

def send_container_wise_intelligence_email(report_items):
    if not GMAIL_SENDER or not GMAIL_APP_PASSWORD:
        print("[!] GMAIL credentials missing.")
        return

    subject = f"🚨 [VASHI APMC REPORT] Fruit Imports: {len(report_items)} Consignment(s) Detected!"
    cards_html = ""

    for item in report_items:
        containers_blocks = ""
        for c in item["containers_detail"]:
            ldb_link = f"https://ldb.co.in/ldb/containersearch/39/{c['cntr']}"
            status_color = "#d93025" if "CFS OUT" in c["latest_milestone"] else "#137333"

            containers_blocks += f"""
            <div style="background: #ffffff; border: 1px solid #e0e0e0; border-radius: 6px; margin-bottom: 12px; padding: 14px; border-left: 5px solid {status_color};">
                <div style="border-bottom: 1px solid #f1f3f4; padding-bottom: 6px; margin-bottom: 8px;">
                    <span style="font-size: 15px; font-weight: bold; font-family: monospace; color: #1a73e8;">
                        <a href="{ldb_link}" target="_blank" style="text-decoration: none; color: #1a73e8;">{c['cntr']}</a>
                    </span>
                    <span style="float: right; background-color: #f1f3f4; color: #202124; padding: 2px 7px; border-radius: 4px; font-size: 12px; font-weight: bold;">
                        ISO: {c['iso']} | {c['temp']}°C
                    </span>
                </div>
                <table style="width: 100%; border-collapse: collapse; font-size: 12px; line-height: 1.5;">
                    <tr><td style="color: #5f6368; width: 32%;"><strong>Port Discharge:</strong></td><td>{c['port_in']}</td></tr>
                    <tr><td style="color: #5f6368;"><strong>Port Gate OUT:</strong></td><td>{c['port_out']}</td></tr>
                    <tr><td style="color: #5f6368;"><strong>CFS Yard:</strong></td><td><strong>{c['cfs_name']}</strong></td></tr>
                    <tr><td style="color: #5f6368;"><strong>Current Status:</strong></td><td style="color: {status_color}; font-weight: bold;">{c['latest_milestone']}</td></tr>
                    <tr><td style="color: #5f6368;"><strong>APMC Decision:</strong></td><td style="color: {status_color}; font-weight: bold;">{c['market_pressure']}</td></tr>
                </table>
            </div>
            """

        cards_html += f"""
        <div style="background: #ffffff; border: 1px solid #dadce0; border-radius: 8px; margin-bottom: 24px; padding: 18px;">
            <div style="border-bottom: 2px solid #1a73e8; padding-bottom: 8px; margin-bottom: 12px;">
                <span style="font-size: 17px; font-weight: bold; color: #1a73e8;">Master B/L: {item['master_bl']}</span>
                <span style="float: right; background-color: #e8f0fe; color: #1a73e8; padding: 3px 9px; border-radius: 4px; font-size: 12px; font-weight: bold;">{item['line']}</span>
            </div>
            <table style="width: 100%; border-collapse: collapse; font-size: 13px; line-height: 1.5; margin-bottom: 14px;">
                <tr><td style="color: #5f6368; width: 30%;"><strong>Fruit Cargo:</strong></td><td style="color: #d93025; font-weight: bold;">{item['fruit']}</td></tr>
                <tr><td style="color: #5f6368;"><strong>Packaging / Invoices:</strong></td><td><strong>{item['cartons']}</strong> | Inv: {item['invoices']}</td></tr>
                <tr><td style="color: #5f6368;"><strong>Vessel & Voyage:</strong></td><td>{item['vessel']} ({item['voyage']}) &bull; Terminal: <strong>{item['terminal']}</strong></td></tr>
                <tr><td style="color: #5f6368;"><strong>Port of Loading:</strong></td><td>{item['pol']}</td></tr>
            </table>
            <div style="font-size: 13px; font-weight: bold; margin-bottom: 8px; color: #202124;">Container Movement Breakdown:</div>
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
                <p style="margin: 4px 0 0 0; font-size: 13px; opacity: 0.95;">Automated Container Tracking & Vashi APMC Timing Briefing</p>
            </div>
            <div style="padding: 20px;">
                {cards_html}
                <div style="text-align: center; margin-top: 20px; font-size: 11px; color: #80868b; border-top: 1px solid #f1f3f4; padding-top: 15px;">
                    Automated JNCH Reefer Intelligence &bull; Continuous Polling via Cloud Actions
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
    msg.attach(MIMEText(html, "html"))

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(GMAIL_SENDER, GMAIL_APP_PASSWORD)
            server.sendmail(GMAIL_SENDER, ALERT_RECEIVER, msg.as_string())
        print(f"[+] Detailed container intelligence email delivered to {ALERT_RECEIVER}!")
    except Exception as e:
        print(f"[!] Email dispatch error: {e}")

# --- Autonomous DPD Scraper & Pipeline ---

async def run_tracker():
    seen_cntrs = get_seen_containers()
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

            # Paginate through DPD table to find active target lines
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

        if all_reefers:
            master_df = pd.concat(all_reefers, ignore_index=True)
            new_reefers = master_df[~master_df["_CNTR"].isin(seen_cntrs)].copy()

            if not new_reefers.empty:
                print(f"\n[***] Discovered {len(new_reefers)} NEW reefer(s). Starting Intelligence Pipeline...")
                report_items = []

                grouped = new_reefers.groupby(["_VESSEL", "_VOYAGE", "_LINE"])

                for (vessel, voyage, line), group in grouped:
                    sample_cntr = group.iloc[0]["_CNTR"]
                    term_hint = group.iloc[0]["_TERMINAL"]
                    pol_val = group.iloc[0]["_POL"]

                    print(f"\n[*] Resolving Master B/L for consignment {vessel} ({sample_cntr})...")
                    master_bl, active_terminal = await cascade_resolve_master_bl(page, sample_cntr, term_hint)
                    print(f"    -> Terminal: {active_terminal} | Master B/L: {master_bl}")

                    icegate_data = await scrape_icegate(page, master_bl)

                    containers_detail = []
                    for _, row in group.iterrows():
                        cntr = row["_CNTR"]
                        grp_cfs = row["_GROUP_CFS"]
                        temp_val = row["_TEMP"]
                        iso_val = row["_ISO"]

                        print(f"    [*] Fetching Live LDB Status for {cntr}...")
                        ldb = await scrape_ldb_live_status(page, cntr, grp_cfs)

                        containers_detail.append({
                            "cntr": cntr,
                            "iso": iso_val,
                            "temp": temp_val,
                            "port_in": ldb["port_in_time"],
                            "port_out": ldb["port_out_time"],
                            "cfs_name": ldb["cfs_name"],
                            "latest_milestone": ldb["latest_milestone"],
                            "market_pressure": ldb["market_pressure"]
                        })

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
                        "containers_detail": containers_detail
                    })

                if os.path.exists(MASTER_LOG_PATH):
                    existing = pd.read_csv(MASTER_LOG_PATH)
                    pd.concat([existing, new_reefers]).drop_duplicates(subset=["_CNTR"]).to_csv(MASTER_LOG_PATH, index=False)
                else:
                    new_reefers.to_csv(MASTER_LOG_PATH, index=False)

                send_container_wise_intelligence_email(report_items)
                mark_containers_seen(new_reefers["_CNTR"].tolist())
            else:
                print("[*] All detected reefers have already been processed.")
        else:
            print("[-] Scan complete. No active fresh fruit reefers found in recent filings.")

        await browser.close()

if __name__ == "__main__":
    asyncio.run(run_tracker())