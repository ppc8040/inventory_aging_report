#!/usr/bin/env python3
"""
Inventory Aging Report  (SAP S/4HANA Cloud -> SharePoint Excel)

What it does
  1. Pulls Material Stock (MB52 equivalent) for ALL plants. Keeps only non-zero Unrestricted (01)
     stock with plant / storage location / batch.
  2. Pulls Material Documents (MB51 equivalent) for those materials and, for every
     Material + Plant + Batch, works out:
        - earliest Goods Receipt posting date in that plant
        - net cumulative quantity in that plant (receipts - issues, reversals netted)
     Batches whose net quantity in that plant is 0 are dropped.
  3. Writes three tabs to "Inventory Aging Report.xlsx" in the Indus_USA SharePoint site:
        Raw data              - the merged table
        Aging data            - Material x Plant, qty in 0-30 / 31-60 / 61-90 / 90+ day buckets
        Transit and Transfer  - non-zero stock in transit / transfer stock types (called out separately)

Usage
    python inventory_aging_report.py --limit 5 --no-upload     # quick trial, 5 materials, CSVs only
    python inventory_aging_report.py --no-upload               # full run, CSVs only
    python inventory_aging_report.py                           # full run + upload to SharePoint
"""

import argparse
import logging
import os
import re
import sys
import time
from datetime import datetime
from urllib.parse import quote

import msal
import numpy as np
import pandas as pd
import requests

try:                      # optional: lets you keep secrets in a local .env file
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass


def require_env(name):
    val = os.environ.get(name)
    if not val:
        sys.exit("Missing environment variable %s. Set it in your shell, a local .env file, or GitHub secrets." % name)
    return val


# ============================================================================
# CONFIGURATION
# ============================================================================
SAP_USERNAME = require_env("SAP_USERNAME")
SAP_PASSWORD = require_env("SAP_PASSWORD")

SAP_HOST = "https://my409486-api.s4hana.cloud.sap"
STOCK_URL = SAP_HOST + "/sap/opu/odata/sap/API_MATERIAL_STOCK_SRV/A_MatlStkInAcctMod"
DOC_ITEM_URL = SAP_HOST + "/sap/opu/odata/sap/API_MATERIAL_DOCUMENT_SRV/A_MaterialDocumentItem"
PRODUCT_URL = SAP_HOST + "/sap/opu/odata/sap/API_PRODUCT_SRV/A_Product"
PRODUCT_DESC_URL = SAP_HOST + "/sap/opu/odata/sap/API_PRODUCT_SRV/A_ProductDescription"
PRODUCT_GROUP_TEXT_URL = SAP_HOST + "/sap/opu/odata/sap/API_PRODUCTGROUP_SRV/A_ProductGroupText"
DESC_LANGUAGE = "EN"

# Product type descriptions: edit the wording to match your system (there is no standard read API for these).
PRODUCT_TYPE_NAMES = {
    "ROH": "Raw Material", "HALB": "Semifinished Product", "FERT": "Finished Product",
    "HAWA": "Trading Goods", "VERP": "Packaging", "ERSA": "Spare Parts",
    "HIBE": "Operating Supplies", "NLAG": "Non-Stock Material", "DIEN": "Service",
}

# Microsoft Graph (same app registration as the Open Order Report script)
TENANT_AUTHORITY = "https://login.microsoftonline.com/ed97e9bb-e119-4bd4-ab00-307d64bdf908"
CLIENT_ID = "a2239be7-12cd-442d-983f-ea7e316ef767"
CLIENT_SECRET = require_env("AZURE_CLIENT_SECRET")

SHAREPOINT_HOSTNAME = "indusair.sharepoint.com"
SHAREPOINT_SITE_PATH = "/sites/Indus_USA"
SHAREPOINT_FOLDER = ""            # sub-folder inside the site's Documents library; "" = library root
EXCEL_FILE_NAME = "Inventory Aging Report.xlsx"

RAW_SHEET = "Raw data"
AGING_SHEET = "Aging data"
TRANSIT_SHEET = "Transit and Transfer"

# SAP stock type codes (field InventoryStockType)
UNRESTRICTED = "01"
TRANSIT_TRANSFER_TYPES = {
    "04": "Stock transfer - storage location level",
    "05": "Stock transfer - plant level",
    "06": "Stock in transit",
}

# Movement types counted as a "Goods Receipt" for the earliest-date logic
GR_MOVEMENT_TYPES = {"101", "501", "561"}
# If a batch has no GR at all in a plant (e.g. it only arrived via a plant transfer),
# fall back to the first receipt of any kind and flag it in "Date Source".
USE_FALLBACK_RECEIPT = True

PAGE_SIZE = 1000
MATERIALS_PER_DOC_CALL = 10

# ============================================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler("inventory_aging.log", encoding="utf-8"), logging.StreamHandler()],
)
log = logging.getLogger("aging")


# ---------------------------------------------------------------------------
# HTTP helpers
# ---------------------------------------------------------------------------
def http(method, url, retries=4, **kwargs):
    """requests wrapper with retry on 429 / 5xx / network errors. Raises with the body on other errors."""
    kwargs.setdefault("timeout", 120)
    last = None
    for attempt in range(retries):
        try:
            resp = requests.request(method, url, **kwargs)
        except requests.RequestException as exc:
            last = exc
            time.sleep(2 * (2 ** attempt))
            continue
        if resp.status_code == 429 or resp.status_code >= 500:
            wait = int(resp.headers.get("Retry-After", 2 * (2 ** attempt)))
            log.warning("HTTP %s from %s - retrying in %ss", resp.status_code, url[:90], wait)
            last = RuntimeError("HTTP %s: %s" % (resp.status_code, resp.text[:300]))
            time.sleep(wait)
            continue
        return resp
    raise RuntimeError("Request failed after %d attempts: %s" % (retries, last))


def build_url(base, params):
    parts = []
    for k, v in params.items():
        parts.append("%s=%s" % (k, quote(str(v), safe="(),'/$")))
    return base + "?" + "&".join(parts)


def sap_get_all(base_url, params, label):
    """Fetch every row of an OData V2 entity set (JSON) using $top/$skip paging."""
    rows, skip = [], 0
    while True:
        p = dict(params)
        p["$format"] = "json"
        p["$top"] = PAGE_SIZE
        p["$skip"] = skip
        resp = http("GET", build_url(base_url, p), auth=(SAP_USERNAME, SAP_PASSWORD),
                    headers={"Accept": "application/json"})
        if resp.status_code != 200:
            raise RuntimeError("SAP %s returned HTTP %s: %s" % (label, resp.status_code, resp.text[:600]))
        data = resp.json()["d"]
        results = data["results"] if isinstance(data, dict) else data
        if not results:
            break
        rows.extend(results)
        skip += len(results)
    return rows


def parse_odata_date(value):
    if not value:
        return pd.NaT
    m = re.search(r"/Date\((-?\d+)", str(value))
    return pd.Timestamp(int(m.group(1)), unit="ms") if m else pd.NaT


# ---------------------------------------------------------------------------
# SAP extraction
# ---------------------------------------------------------------------------
def fetch_stock():
    types = [UNRESTRICTED] + list(TRANSIT_TRANSFER_TYPES.keys())
    flt = " or ".join("InventoryStockType eq '%s'" % t for t in types)
    params = {
        "$filter": flt,
        "$select": "Material,Plant,StorageLocation,Batch,InventoryStockType,InventorySpecialStockType,"
                   "MaterialBaseUnit,MatlWrhsStkQtyInMatlBaseUnit",
    }
    log.info("Fetching material stock (all plants) ...")
    rows = sap_get_all(STOCK_URL, params, "Material Stock")
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df = df.drop(columns=[c for c in df.columns if c == "__metadata"])
    for c in ["Material", "Plant", "StorageLocation", "Batch", "InventoryStockType",
              "InventorySpecialStockType", "MaterialBaseUnit"]:
        df[c] = df[c].fillna("").astype(str).str.strip()
    df["Qty"] = pd.to_numeric(df["MatlWrhsStkQtyInMatlBaseUnit"], errors="coerce").fillna(0.0)
    log.info("Stock rows returned: %d. Stock type mix (rows): %s", len(df),
             df["InventoryStockType"].value_counts().to_dict())
    return df


def fetch_movements(materials):
    cols = ("Material,Plant,Batch,GoodsMovementType,DebitCreditCode,QuantityInBaseUnit,"
            "GoodsMovementIsCancelled,MaterialDocument,MaterialDocumentYear,MaterialDocumentItem")
    select = cols + ",to_MaterialDocumentHeader/PostingDate"
    out = []
    chunks = [materials[i:i + MATERIALS_PER_DOC_CALL] for i in range(0, len(materials), MATERIALS_PER_DOC_CALL)]
    for n, chunk in enumerate(chunks, 1):
        flt = "(" + " or ".join("Material eq '%s'" % m.replace("'", "''") for m in chunk) + ")"
        params = {"$filter": flt, "$select": select, "$expand": "to_MaterialDocumentHeader"}
        rows = sap_get_all(DOC_ITEM_URL, params, "Material Document")
        for r in rows:
            hdr = r.get("to_MaterialDocumentHeader") or {}
            out.append({
                "Material": (r.get("Material") or "").strip(),
                "Plant": (r.get("Plant") or "").strip(),
                "Batch": (r.get("Batch") or "").strip(),
                "MovementType": (r.get("GoodsMovementType") or "").strip(),
                "DC": (r.get("DebitCreditCode") or "S").strip().upper(),
                "Qty": r.get("QuantityInBaseUnit"),
                "Cancelled": bool(r.get("GoodsMovementIsCancelled")),
                "PostingDate": parse_odata_date(hdr.get("PostingDate")),
            })
        log.info("Material documents: batch %d/%d done (%d item lines so far)", n, len(chunks), len(out))
    mv = pd.DataFrame(out)
    if not mv.empty:
        mv["Qty"] = pd.to_numeric(mv["Qty"], errors="coerce").fillna(0.0)
    return mv



def fetch_product_master(materials):
    """Material description, product type and product group (+ group text) for the given materials."""
    cols = ["Material", "Material Description", "Product Type", "Product Type Description",
            "Product Group", "Product Group Description"]
    if not materials:
        return pd.DataFrame(columns=cols)
    prod, desc = {}, {}
    chunks = [materials[i:i + 20] for i in range(0, len(materials), 20)]
    for n, chunk in enumerate(chunks, 1):
        mf = "(" + " or ".join("Product eq '%s'" % m.replace("'", "''") for m in chunk) + ")"
        for r in sap_get_all(PRODUCT_URL, {"$filter": mf, "$select": "Product,ProductType,ProductGroup"},
                             "Product"):
            prod[(r.get("Product") or "").strip()] = ((r.get("ProductType") or "").strip(),
                                                        (r.get("ProductGroup") or "").strip())
        for r in sap_get_all(PRODUCT_DESC_URL, {"$filter": mf + " and Language eq '%s'" % DESC_LANGUAGE,
                                                "$select": "Product,ProductDescription"}, "Product Description"):
            desc[(r.get("Product") or "").strip()] = (r.get("ProductDescription") or "").strip()
        if n % 10 == 0 or n == len(chunks):
            log.info("Product master: %d/%d batches", n, len(chunks))

    group_text = {}
    try:
        for g in sap_get_all(PRODUCT_GROUP_TEXT_URL, {"$filter": "Language eq '%s'" % DESC_LANGUAGE},
                             "Product Group Text"):
            code = (g.get("MaterialGroup") or g.get("ProductGroup") or "").strip()
            group_text[code] = (g.get("MaterialGroupText") or g.get("MaterialGroupName") or "").strip()
    except Exception as exc:
        log.warning("Product group descriptions unavailable (%s). Group description column will be blank. "
                    "The API user needs the Product Group read scenario (API_PRODUCTGROUP_SRV).", exc)

    rows = []
    for m in materials:
        ptype, pgroup = prod.get(m, ("", ""))
        rows.append({"Material": m, "Material Description": desc.get(m, ""), "Product Type": ptype,
                     "Product Type Description": PRODUCT_TYPE_NAMES.get(ptype, ""),
                     "Product Group": pgroup, "Product Group Description": group_text.get(pgroup, "")})
    return pd.DataFrame(rows, columns=cols)


def add_attributes(df, master):
    """Left-join the descriptive columns and place them right after Material."""
    if df.empty:
        for c in master.columns[1:]:
            df[c] = ""
    else:
        df = df.merge(master, on="Material", how="left")
        for c in master.columns[1:]:
            df[c] = df[c].fillna("")
    attrs = list(master.columns[1:])
    rest = [c for c in df.columns if c != "Material" and c not in attrs]
    return df[["Material"] + attrs + rest]


# ---------------------------------------------------------------------------
# Transformation
# ---------------------------------------------------------------------------
BUCKETS = ["0-30", "31-60", "61-90", "90+"]


def bucket_for(age):
    if pd.isna(age):
        return "Unknown"
    if age <= 30:
        return "0-30"
    if age <= 60:
        return "31-60"
    if age <= 90:
        return "61-90"
    return "90+"


def build_tables(stock, movements, run_ts):
    keys = ["Material", "Plant", "Batch"]

    # Unrestricted, regular (non-special) stock, non-zero
    unres = stock[(stock["InventoryStockType"] == UNRESTRICTED) & (stock["Qty"] != 0)].copy()
    special = (unres["InventorySpecialStockType"] != "").sum()
    if special:
        log.info("Ignoring %d unrestricted rows that are special stock (consignment / sales-order / project).", special)
    unres = unres[unres["InventorySpecialStockType"] == ""]
    unres = (unres.groupby(["Material", "Plant", "StorageLocation", "Batch", "MaterialBaseUnit"], as_index=False)["Qty"]
             .sum())
    unres = unres[unres["Qty"] != 0]
    log.info("Unrestricted non-zero stock lines: %d (%d materials)", len(unres), unres["Material"].nunique())

    # Transit & transfer, called out separately
    tt = stock[stock["InventoryStockType"].isin(TRANSIT_TRANSFER_TYPES.keys()) & (stock["Qty"] != 0)].copy()
    tt["Stock Type Description"] = tt["InventoryStockType"].map(TRANSIT_TRANSFER_TYPES)
    transit = tt[["Material", "Plant", "StorageLocation", "InventoryStockType", "Stock Type Description",
                  "Batch", "MaterialBaseUnit", "Qty"]].rename(columns={
        "StorageLocation": "Storage Location", "InventoryStockType": "Stock Type",
        "MaterialBaseUnit": "Base Unit", "Qty": "Quantity"})
    transit = transit.sort_values(["Material", "Plant", "Stock Type"]).reset_index(drop=True)

    # Movement-derived facts per Material + Plant + Batch
    if movements.empty:
        facts = pd.DataFrame(columns=keys + ["NetQty", "FirstGR", "FirstReceipt"])
    else:
        mv = movements.copy()
        mv["Signed"] = np.where(mv["DC"] == "H", -mv["Qty"], mv["Qty"])
        net = mv.groupby(keys)["Signed"].sum().rename("NetQty")
        live_receipts = mv[(mv["DC"] != "H") & (~mv["Cancelled"])]
        first_gr = (live_receipts[live_receipts["MovementType"].isin(GR_MOVEMENT_TYPES)]
                    .groupby(keys)["PostingDate"].min().rename("FirstGR"))
        first_any = live_receipts.groupby(keys)["PostingDate"].min().rename("FirstReceipt")
        facts = pd.concat([net, first_gr, first_any], axis=1).reset_index()

    raw = unres.merge(facts, on=keys, how="left")

    # Drop batches whose cumulative quantity in that plant nets to zero
    zero_mask = raw["NetQty"].notna() & (raw["NetQty"].abs() < 1e-9)
    if zero_mask.any():
        log.info("Dropping %d stock lines whose cumulative movement qty in the plant is 0.", zero_mask.sum())
    raw = raw[~zero_mask].copy()

    def pick_date(r):
        if pd.notna(r["FirstGR"]):
            return r["FirstGR"], "First GR"
        if USE_FALLBACK_RECEIPT and pd.notna(r["FirstReceipt"]):
            return r["FirstReceipt"], "Fallback: first receipt (no GR found)"
        return pd.NaT, "No receipt found"

    picked = raw.apply(pick_date, axis=1, result_type="expand")
    raw["Earliest GR Date"] = pd.to_datetime(picked[0])
    raw["Date Source"] = picked[1]
    raw["Age (days)"] = (run_ts.normalize() - raw["Earliest GR Date"].dt.normalize()).dt.days
    raw["Aging Bucket"] = raw["Age (days)"].apply(bucket_for)
    raw["Batch Managed"] = np.where(raw["Batch"] != "", "Yes", "No")
    raw["As Of Date"] = run_ts.strftime("%Y-%m-%d")

    raw = raw.rename(columns={"StorageLocation": "Storage Location", "MaterialBaseUnit": "Base Unit",
                              "Qty": "Unrestricted Qty", "NetQty": "Net Movement Qty (Plant+Batch)"})
    raw = raw[["Material", "Plant", "Storage Location", "Batch", "Batch Managed", "Base Unit", "Unrestricted Qty",
               "Earliest GR Date", "Date Source", "Net Movement Qty (Plant+Batch)", "Age (days)",
               "Aging Bucket", "As Of Date"]]
    raw = raw.sort_values(["Material", "Plant", "Storage Location", "Batch"]).reset_index(drop=True)

    # Aging table: Material x Plant
    if raw.empty:
        aging = pd.DataFrame(columns=["Material", "Plant", "Base Unit"] + BUCKETS + ["Unknown", "Total"])
    else:
        piv = raw.pivot_table(index=["Material", "Plant", "Base Unit"], columns="Aging Bucket",
                              values="Unrestricted Qty", aggfunc="sum", fill_value=0.0)
        for b in BUCKETS + ["Unknown"]:
            if b not in piv.columns:
                piv[b] = 0.0
        piv = piv[BUCKETS + ["Unknown"]]
        piv["Total"] = piv.sum(axis=1)
        aging = piv.reset_index()
        aging.columns.name = None
        aging = aging.rename(columns={b: b + " days" for b in BUCKETS})
        aging = aging.rename(columns={"Unknown": "Unknown (no receipt date)"})
    return raw, aging, transit


# ---------------------------------------------------------------------------
# SharePoint / Excel via Microsoft Graph
# ---------------------------------------------------------------------------
GRAPH = "https://graph.microsoft.com/v1.0"


def graph_token():
    app = msal.ConfidentialClientApplication(CLIENT_ID, authority=TENANT_AUTHORITY, client_credential=CLIENT_SECRET)
    res = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in res:
        raise RuntimeError("Could not get Graph token: %s" % res.get("error_description", res))
    return res["access_token"]


def col_letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def to_py(v):
    if v is None or (isinstance(v, float) and np.isnan(v)) or v is pd.NaT:
        return ""
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.floating,)):
        return "" if np.isnan(v) else float(v)
    if isinstance(v, pd.Timestamp):
        return (v - pd.Timestamp("1899-12-30")).days   # Excel date serial
    return v


def upload_to_sharepoint(sheets):
    token = graph_token()
    auth = {"Authorization": "Bearer " + token}
    jhdr = dict(auth, **{"Content-Type": "application/json"})

    r = http("GET", "%s/sites/%s:%s" % (GRAPH, SHAREPOINT_HOSTNAME, SHAREPOINT_SITE_PATH), headers=auth)
    if r.status_code != 200:
        raise RuntimeError("Could not resolve SharePoint site (HTTP %s): %s. If 403, the Azure app needs access to "
                           "the Indus_USA site." % (r.status_code, r.text[:400]))
    site_id = r.json()["id"]
    log.info("SharePoint site id: %s", site_id)

    rel = (SHAREPOINT_FOLDER.strip("/") + "/" if SHAREPOINT_FOLDER.strip("/") else "") + EXCEL_FILE_NAME
    item = "%s/sites/%s/drive/root:/%s:" % (GRAPH, site_id, quote(rel))

    r = http("GET", item, headers=auth)
    if r.status_code == 404:
        log.info("Workbook not found - creating a blank one.")
        import io
        from openpyxl import Workbook
        buf = io.BytesIO()
        Workbook().save(buf)
        put = http("PUT", item + "/content", headers=dict(auth, **{
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}),
            data=buf.getvalue())
        if put.status_code not in (200, 201):
            raise RuntimeError("Could not create workbook: %s %s" % (put.status_code, put.text[:400]))
    elif r.status_code != 200:
        raise RuntimeError("Could not access workbook: %s %s" % (r.status_code, r.text[:400]))

    s = http("POST", item + "/workbook/createSession", headers=jhdr, json={"persistChanges": True})
    if s.status_code not in (200, 201):
        raise RuntimeError("Could not open Excel session: %s %s" % (s.status_code, s.text[:400]))
    session_id = s.json()["id"]
    wh = dict(jhdr, **{"workbook-session-id": session_id})

    try:
        existing = http("GET", item + "/workbook/worksheets", headers=wh).json().get("value", [])
        existing_names = [w["name"] for w in existing]

        for name, (df, date_cols) in sheets.items():
            if name not in existing_names:
                a = http("POST", item + "/workbook/worksheets/add", headers=wh, json={"name": name})
                if a.status_code not in (200, 201):
                    raise RuntimeError("Could not add sheet '%s': %s" % (name, a.text[:300]))
            ws = "%s/workbook/worksheets('%s')" % (item, name)
            http("POST", ws + "/usedRange/clear", headers=wh, json={"applyTo": "All"})

            headers = list(df.columns)
            ncols = len(headers)
            body = [[to_py(v) for v in row] for row in df.itertuples(index=False, name=None)]
            all_rows = [headers] + body

            CH = 2000
            for start in range(0, len(all_rows), CH):
                part = all_rows[start:start + CH]
                r1, r2 = start + 1, start + len(part)
                addr = "A%d:%s%d" % (r1, col_letter(ncols), r2)
                u = http("PATCH", "%s/range(address='%s')" % (ws, addr), headers=wh, json={"values": part})
                if u.status_code not in (200, 201):
                    raise RuntimeError("Write to '%s' failed: %s %s" % (name, u.status_code, u.text[:400]))

            # cosmetics (non-critical)
            try:
                hdr_addr = "A1:%s1" % col_letter(ncols)
                http("PATCH", "%s/range(address='%s')/format/font" % (ws, hdr_addr), headers=wh,
                     json={"bold": True, "color": "#FFFFFF"})
                http("PATCH", "%s/range(address='%s')/format/fill" % (ws, hdr_addr), headers=wh,
                     json={"color": "#2F4F4F"})
                for dc in date_cols:
                    if dc in headers and len(body) > 0:
                        L = col_letter(headers.index(dc) + 1)
                        fmt = [["dd-mmm-yy"]] * len(body)
                        http("PATCH", "%s/range(address='%s2:%s%d')" % (ws, L, L, len(body) + 1), headers=wh,
                             json={"numberFormat": fmt})
                http("POST", "%s/range(address='%s')/format/autofitColumns" % (ws, hdr_addr), headers=wh)
            except Exception as exc:   # formatting must never fail the run
                log.warning("Formatting skipped for '%s': %s", name, exc)
            log.info("Wrote %d data rows to tab '%s'", len(body), name)
    finally:
        http("POST", item + "/workbook/closeSession", headers=dict(auth, **{"workbook-session-id": session_id}))


# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=0, help="only process the first N materials (for a trial run)")
    ap.add_argument("--no-upload", action="store_true", help="write CSV files locally instead of uploading")
    args = ap.parse_args()

    run_ts = pd.Timestamp(datetime.now())
    stock = fetch_stock()
    if stock.empty:
        log.error("No stock rows returned from SAP - check the API access / filters.")
        sys.exit(1)

    materials = sorted(stock.loc[(stock["InventoryStockType"] == UNRESTRICTED) & (stock["Qty"] != 0), "Material"]
                       .unique().tolist())
    if args.limit:
        materials = materials[:args.limit]
        stock = stock[stock["Material"].isin(materials)]
        log.info("TRIAL MODE: limited to %d materials", len(materials))
    log.info("Materials with non-zero unrestricted stock: %d", len(materials))

    movements = fetch_movements(materials)
    log.info("Material document item lines fetched: %d", len(movements))

    raw, aging, transit = build_tables(stock, movements, run_ts)

    all_mats = sorted(set(raw["Material"]) | set(transit["Material"]))
    log.info("Fetching descriptions / product type / product group for %d materials ...", len(all_mats))
    master = fetch_product_master(all_mats)
    raw = add_attributes(raw, master)
    aging = add_attributes(aging, master)
    transit = add_attributes(transit, master)

    if args.no_upload:
        raw.to_csv("raw_data.csv", index=False)
        aging.to_csv("aging_data.csv", index=False)
        transit.to_csv("transit_and_transfer.csv", index=False)
        log.info("CSV files written (raw_data.csv, aging_data.csv, transit_and_transfer.csv).")
    else:
        upload_to_sharepoint({
            RAW_SHEET: (raw, ["Earliest GR Date"]),
            AGING_SHEET: (aging, []),
            TRANSIT_SHEET: (transit, []),
        })

    print("\nDONE  | Raw data rows: %d | Aging rows (Material x Plant): %d | Transit/Transfer rows: %d"
          % (len(raw), len(aging), len(transit)))
    if not raw.empty:
        print("Aging bucket split (qty):", raw.groupby("Aging Bucket")["Unrestricted Qty"].sum().round(2).to_dict())
        print("Date source split (lines):", raw["Date Source"].value_counts().to_dict())


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        log.error("FAILED: %s", exc)
        import traceback
        log.error(traceback.format_exc())
        print("\nFAILED - see inventory_aging.log for details. Message: %s" % exc)
        sys.exit(1)
