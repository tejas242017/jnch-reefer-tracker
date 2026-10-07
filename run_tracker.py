import os
import re
import sys
import smtplib
import asyncio
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd

# Environment detection: Local E: Drive vs GitHub Actions Cloud Runner
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
    # Check all sheets for Excel
    for engine in ["openpyxl", "xlrd"]:
        try:
            xls = pd.ExcelFile(filepath, engine=engine)
            # Prefer sheets with Advance List or BMCTPL/GTI
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

        # Filter 1: Reefer Equipment Code
        reefer_regex = "|".join([rf"\b{re.escape(c)}\b" for c in REEFER_CODES]) + r"|RH|RF|REEF"
        is_reefer_code = df[type_col].astype(str).str.contains(reefer_regex, case=False, na=False)

        # Filter 2: Temperature Setpoint (Fresh Produce: +1.0°C to +6.0°C)
        is_fresh_temp = pd.Series(False, index=df.index)
        if temp_col:
            temps = pd.to_numeric(df[temp_col], errors="coerce")
            is_fresh_temp = (temps >= 1.0) & (temps <= 6.5)

        # Container is targeted if it matches reefer ISO or fresh fruit temp setpoint
        fruit_mask = is_reefer_code | is_fresh_temp
        if temp_col:
            # Explicitly exclude frozen cargo (-23°C) and ambient/pharma (+20°C)
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

        # Detect Discharging Terminal from manifest/filepath
        term = "CASCADE"
        fname_upper = filepath.upper()
        if "BMCT" in fname_upper or "PSA" in fname_upper:
            term = "BMCT"
        elif "GTI" in fname_upper or "APMT" in fname_upper:
            term = "GTI"
        elif "NSICT" in fname_upper or "NSIGT" in fname_upper or "DPW" in fname_upper:
            term = "DPW"
        elif "NSFT" in fname_upper or "JNPCT" in fname_upper:
            term = "NSFT"
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
    # 1. Targeted check if hinted
    if hinted_terminal == "BMCT":
        bl, term = await resolve_via_bmct(page, cntr_no)
        if bl: return bl, term
    elif hinted_terminal == "GTI":
        bl, term = await resolve_via_gti(page, cntr_no)
        if bl: return bl, term
    elif hinted_terminal == "DPW":
        bl, term = await resolve_via_dpworld(page, cntr_no)
        if bl: return bl, term

    # 2. Sequential Cascade Fallback across all terminals
    bl, term = await resolve_via_bmct(page, cntr_no)
    if bl: return bl, term

    bl, term = await resolve_via_gti(page, cntr_no)
    if bl: return bl, term

    bl, term = await resolve_via_dpworld(page, cntr_no)
    if bl: return bl, term

    return "UNRESOLVED", "Unknown Terminal"

# --- ICEGATE Cargo & Invoice Scraper ---

async def scrape_icegate(page, master_bl):
    data = {
        "fruit": "Perishable Cargo",
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

        # Location INNSA1
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

        # Master B/L
        bl_box = page.locator("input[placeholder*='Enter Master BL']").first
        await bl_box.click()
        await bl_box.fill(master_bl)
        await page.wait_for_timeout(600)

        # Search
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

        # Click View for Container Details
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

# --- LDB CFS & Movement Milestone Scraper ---

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

async def scrape_ldb_status(page, cntr_no, manifest_group_code):
    info = {
        "cfs_name": CFS_NAME_MAP.get(manifest_group_code, manifest_group_code),
        "milestone": "En-route to Yard",
        "market_pressure": "NORMAL",
        "last_update": datetime.now().strftime("%d-%b-%Y")
    }
    try:
        url = f"https://ldb.co.in/ldb/containersearch/39/{cntr_no}"
        await page.goto(url, wait_until="domcontentloaded", timeout=20000)
        await page.wait_for_timeout(2000)

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
                if len(clean) < 70:
                    info["cfs_name"] = clean
                    break

        for line in body.splitlines():
            clean = line.strip()
            if any(k in clean.upper() for k in ["CFS OUT", "CFS IN", "GATE OUT", "GATE IN"]):
                if len(clean) < 50:
                    info["milestone"] = clean
                    break

        # Strategic Vashi APMC Timing Analysis
        today_weekday = datetime.today().weekday()  # 3=Thursday, 4=Friday, 5=Saturday, 0=Monday
        if "CFS OUT" in info["milestone"].upper():
            info["market_pressure"] = "IMMEDIATE (Landing at Vashi APMC Tonight)"
        elif "CFS IN" in info["milestone"].upper():
            if today_weekday in [3, 4, 5]: # Thu, Fri, Sat hold
                info["market_pressure"] = "HIGH MONDAY GLUT RISK (Dumped Mon Night)"
            else:
                info["market_pressure"] = "HOLDING AT CFS (Customs/PQ)"
    except Exception:
        pass
    return info

# --- Executive HTML Email Report ---

def send_market_intelligence_report(report_data):
    if not GMAIL_SENDER or not GMAIL_APP_PASSWORD:
        print("[!] GMAIL credentials missing. Skipping email.")
        return

    subject = f"🍏 Vashi APMC Fruit Supply Intelligence: {len(report_data)} Consignment(s) Detected!"
    cards_html = ""

    for item in report_data:
        ldb_link = f"https://ldb.co.in/ldb/containersearch/39/{item['primary_cntr']}"
        pressure_color = "#d93025" if "GLUT" in item['market_pressure'] or "IMMEDIATE" in item['market_pressure'] else "#137333"

        cards_html += f"""
        <div style="background: #ffffff; border: 1px solid #dadce0; border-radius: 8px; margin-bottom: 22px; padding: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
            <div style="border-bottom: 2px solid #1a73e8; padding-bottom: 10px; margin-bottom: 14px;">
                <span style="font-size: 17px; font-weight: bold; color: #1a73e8;">Master B/L: {item['master_bl']}</span>
                <span style="float: right; background-color: #e8f0fe; color: #1a73e8; padding: 4px 10px; border-radius: 4px; font-size: 12px; font-weight: bold;">{item['line']}</span>
            </div>
            
            <table style="width: 100%; border-collapse: collapse; font-size: 13px; line-height: 1.6;">
                <tr>
                    <td style="color: #5f6368; width: 32%;"><strong>Fruit Cargo:</strong></td>
                    <td style="font-weight: bold; color: #d93025; font-size: 14px;">{item['fruit']}</td>
                </tr>
                <tr>
                    <td style="color: #5f6368;"><strong>Packaging & Volume:</strong></td>
                    <td style="font-weight: bold; color: #202124;">{item['cartons']} ({item['gross_wt']})</td>
                </tr>
                <tr>
                    <td style="color: #5f6368;"><strong>Temperature Setpoint:</strong></td>
                    <td style="color: #1a73e8; font-weight: bold;">{item['temp']}°C (Fresh Fruit Protocol)</td>
                </tr>
                <tr>
                    <td style="color: #5f6368;"><strong>Vessel & Origin (POL):</strong></td>
                    <td>{item['vessel']} ({item['voyage']}) | POL: {item['pol']}</td>
                </tr>
                <tr>
                    <td style="color: #5f6368;"><strong>Berthing Terminal:</strong></td>
                    <td><strong>{item['terminal']}</strong></td>
                </tr>
                <tr>
                    <td style="color: #5f6368;"><strong>Current CFS Yard:</strong></td>
                    <td><strong>{item['cfs_name']}</strong> ({item['milestone']})</td>
                </tr>
                <tr>
                    <td style="color: #5f6368;"><strong>Vashi APMC Pressure:</strong></td>
                    <td style="color: {pressure_color}; font-weight: bold;">{item['market_pressure']}</td>
                </tr>
                <tr>
                    <td style="color: #5f6368;"><strong>Total Reefer Boxes:</strong></td>
                    <td style="font-family: monospace; font-size: 12px;">{', '.join(item['sister_containers']) if item['sister_containers'] else item['primary_cntr']}</td>
                </tr>
            </table>

            <div style="margin-top: 14px; padding-top: 10px; border-top: 1px solid #eee; font-size: 12px;">
                <a href="{ldb_link}" target="_blank" style="color: #1a73e8; text-decoration: none; font-weight: bold;">🔍 Live LDB Container Tracker</a>
            </div>
        </div>
        """

    html = f"""
    <!DOCTYPE html>
    <html>
    <body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; background-color: #f8f9fa; padding: 20px; margin: 0; color: #202124;">
        <div style="max-width: 720px; margin: 0 auto;">
            <div style="background: #1a73e8; color: white; padding: 18px 24px; border-radius: 8px 8px 0 0;">
                <h2 style="margin: 0; font-size: 21px;">🍎 Daily Nhava Sheva Fruit Import Intelligence</h2>
                <p style="margin: 4px 0 0 0; font-size: 13px; opacity: 0.9;">Advance Cargo Volume, CFS Cold-Chain Holds & Vashi APMC Timing</p>
            </div>
            
            <div style="background: white; padding: 20px; border-radius: 0 0 8px 8px; border: 1px solid #dadce0; border-top: none;">
                <div style="background-color: #fef7e0; border-left: 4px solid #f9ab00; padding: 12px; margin-bottom: 20px; font-size: 12px; line-height: 1.5;">
                    <strong>Decision Framework:</strong> Weekend shipping line DO blocks mean reefers clearing Friday/Saturday accumulate in CFS yards. Use the carton volumes below to decide whether to stay put at the CFS on cold-plug power or dispatch to Vashi APMC.
                </div>
                {cards_html}
                <p style="font-size: 11px; color: #9aa0a6; text-align: center; margin-top: 25px;">
                    Automated JNCH Reefer Intelligence • Generated autonomously via GitHub Actions Cloud Runner
                </p>
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
        print(f"[+] Intelligence report successfully dispatched to {ALERT_RECEIVER}!")
    except Exception as e:
        print(f"[!] Email dispatch error: {e}")

# --- Master Execution Routine ---

async def run_tracker():
    seen_cntrs = get_seen_containers()
    all_reefers = []

    async with async_playwright() as p:
        print(f"[*] Launching Chromium ({'Windows Sandbox' if IS_WINDOWS else 'GitHub Cloud Runner'})...")
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()

        print("[*] Stage 1: Scanning DPD JNCH Advance Manifests...")
        await page.goto("https://dpdjnch.com/ShippingLine/AdvanceListing.aspx", wait_until="networkidle")

        rows = await page.locator("table tr").all()
        for row in rows[1:]:
            cols = await row.locator("td").all_text_contents()
            if len(cols) < 5:
                continue

            line_name, vessel_name, voyage_no = cols[2].strip(), cols[3].strip(), cols[4].strip()

            if any(target in line_name.upper() for target in TARGET_LINES):
                btn = row.locator("td:last-child a, td:last-child input[type='submit']")
                if await btn.count() > 0:
                    try:
                        async with page.expect_download(timeout=15000) as dl_info:
                            await btn.first.click()
                        dl = await dl_info.value
                        clean_v = re.sub(r'[^A-Za-z0-9_]', '_', vessel_name)
                        clean_l = re.sub(r'[^A-Za-z0-9_]', '_', line_name)[:25]
                        save_p = os.path.join(DOWNLOAD_DIR, f"{clean_l}_{clean_v}_{voyage_no}.xlsx")
                        await dl.save_as(save_p)

                        found = parse_fresh_fruit_reefers(save_p, line_name, vessel_name, voyage_no)
                        if found is not None and not found.empty:
                            all_reefers.append(found)
                    except Exception:
                        pass

        if all_reefers:
            master_df = pd.concat(all_reefers, ignore_index=True)
            new_reefers = master_df[~master_df["_CNTR"].isin(seen_cntrs)].copy()

            if not new_reefers.empty:
                print(f"[***] Discovered {len(new_reefers)} NEW fresh fruit reefer(s). Starting Intelligence Pipeline...")
                report_items = []

                # Group by Vessel/Voyage to avoid redundant B/L resolutions
                for _, row in new_reefers.iterrows():
                    cntr = row["_CNTR"]
                    line = row["_LINE"]
                    term_hint = row["_TERMINAL"]
                    grp_cfs = row["_GROUP_CFS"]
                    temp_val = row["_TEMP"]

                    print(f"\n[*] Processing Container: {cntr} ({line} | Setpoint: {temp_val}°C)...")

                    # Phase 2 & 3: Cascaded Terminal Resolver (PSA -> GTI -> DPW)
                    master_bl, active_terminal = await cascade_resolve_master_bl(page, cntr, term_hint)
                    print(f"    -> Resolved Terminal: {active_terminal} | Master B/L: {master_bl}")

                    # Phase 4: Customs ICEGATE Public Enquiry
                    icegate_data = await scrape_icegate(page, master_bl)
                    print(f"    -> ICEGATE Cargo: {icegate_data['fruit']} | Volume: {icegate_data['cartons']}")

                    # Phase 5: LDB Physical CFS Location & APMC Timing Analysis
                    ldb_data = await scrape_ldb_status(page, cntr, grp_cfs)
                    print(f"    -> Nominated CFS: {ldb_data['cfs_name']} | Milestone: {ldb_data['milestone']}")
                    print(f"    -> Vashi APMC Timing: {ldb_data['market_pressure']}")

                    report_items.append({
                        "primary_cntr": cntr,
                        "master_bl": master_bl,
                        "line": line,
                        "vessel": row["_VESSEL"],
                        "voyage": row["_VOYAGE"],
                        "pol": row["_POL"],
                        "temp": temp_val,
                        "terminal": active_terminal,
                        "fruit": icegate_data["fruit"],
                        "cartons": icegate_data["cartons"],
                        "gross_wt": icegate_data["gross_wt"],
                        "invoices": icegate_data["invoices"],
                        "sister_containers": icegate_data["sister_containers"] or [cntr],
                        "cfs_name": ldb_data["cfs_name"],
                        "milestone": ldb_data["milestone"],
                        "market_pressure": ldb_data["market_pressure"]
                    })

                # Persist to local/cloud CSV log
                if os.path.exists(MASTER_LOG_PATH):
                    existing = pd.read_csv(MASTER_LOG_PATH)
                    pd.concat([existing, new_reefers]).drop_duplicates(subset=["_CNTR"]).to_csv(MASTER_LOG_PATH, index=False)
                else:
                    new_reefers.to_csv(MASTER_LOG_PATH, index=False)

                # Send executive briefing to your Gmail
                send_market_intelligence_report(report_items)
                mark_containers_seen(new_reefers["_CNTR"].tolist())
            else:
                print("[*] All detected reefers in current queue have already been processed.")
        else:
            print("[-] Scan complete. No active fresh fruit reefers found in recent filings.")

        await browser.close()

if __name__ == "__main__":
    asyncio.run(run_tracker())