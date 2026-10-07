import os
import glob
import pandas as pd

MANIFEST_DIR = r"E:\DPD_Tracker_Sandbox\manifests"

def read_any_format(filepath):
    """Tries openpyxl -> xlrd -> html -> csv in sequence."""
    # 1. Modern Excel (.xlsx)
    try:
        return pd.read_excel(filepath, engine="openpyxl")
    except Exception:
        pass

    # 2. Legacy Excel (.xls / OLE2 binary with 0xD0 magic byte)
    try:
        return pd.read_excel(filepath, engine="xlrd")
    except Exception:
        pass

    # 3. HTML table disguised as Excel
    try:
        tables = pd.read_html(filepath, flavor="html5lib")
        if tables:
            return tables[0]
    except Exception:
        pass

    # 4. CSV disguised as Excel
    try:
        return pd.read_csv(filepath, sep=None, engine="python", on_bad_lines="skip")
    except Exception:
        pass

    return None

def inspect_all():
    files = glob.glob(os.path.join(MANIFEST_DIR, "*.xlsx"))
    print(f"[*] Found {len(files)} downloaded files to inspect.\n")

    for filepath in files:
        fname = os.path.basename(filepath)
        print(f"--- Inspecting: {fname} ---")

        df = read_any_format(filepath)
        if df is None or df.empty:
            print("    [!] Could not parse file format with any engine.")
            continue

        # Clean headers
        df.columns = [str(c).strip().upper() for c in df.columns]
        print(f"    Rows: {len(df)} | Columns: {list(df.columns[:8])}")

        # Find Type & Container Columns
        type_col = next((c for c in df.columns if any(k in c for k in ["TYPE", "ISO", "SIZE", "EQPTYPE", "EQ_TYPE"])), None)
        cntr_col = next((c for c in df.columns if any(k in c for k in ["CONTAINER", "CNTR", "EQ_NO"])), None)
        pol_col = next((c for c in df.columns if any(k in c for k in ["POL", "LOAD", "ORIGIN"])), None)
        pod_col = next((c for c in df.columns if any(k in c for k in ["POD", "DISCH"])), None)
        party_col = next((c for c in df.columns if any(k in c for k in ["CONSIGNEE", "PARTY", "CLIENT", "IMPORTER", "DPD", "OPR"])), None)

        if not type_col or not cntr_col:
            # Check row 1 or 2 for header offset
            for skip in range(1, 4):
                try:
                    t_df = pd.read_excel(filepath, skiprows=skip)
                    t_df.columns = [str(c).strip().upper() for c in t_df.columns]
                    type_col = next((c for c in t_df.columns if any(k in c for k in ["TYPE", "ISO", "SIZE", "EQPTYPE"])), None)
                    cntr_col = next((c for c in t_df.columns if any(k in c for k in ["CONTAINER", "CNTR"])), None)
                    if type_col and cntr_col:
                        df = t_df
                        break
                except Exception:
                    continue

        if not type_col:
            print("    [-] Container Type/ISO column missing.")
            continue

        # Reefer Filter (40RH, 40RF, 45R1, 22R1, REEFER, RF, RH)
        reefer_mask = df[type_col].astype(str).str.contains(r"RH|RF|45R1|22R1|REEF", case=False, na=False)
        reefers = df[reefer_mask]

        if not reefers.empty:
            print(f"    [***] ALERT! Found {len(reefers)} REEFER Containers!")
            for _, r in reefers.head(10).iterrows():
                cntr = r.get(cntr_col, "N/A")
                iso = r.get(type_col, "N/A")
                pol = r.get(pol_col, "N/A") if pol_col else "N/A"
                pod = r.get(pod_col, "N/A") if pod_col else "N/A"
                party = r.get(party_col, "N/A") if party_col else "N/A"
                print(f"        -> Cntr: {cntr} | ISO: {iso} | POL: {pol} | POD: {pod} | Party: {party}")
        else:
            print("    [-] 0 Reefers detected.")
        print()

if __name__ == "__main__":
    inspect_all()