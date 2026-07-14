"""Import the 2026 Tradeplan Excel workbook into JSON for the tradeplan app.

Usage:
    python import_data.py                 (uses the default S: drive path)
    python import_data.py "path\to\file.xlsx"
    python import_data.py --reset-orders  (also overwrite data/orders.json with Excel's orders)

Reads cached (calculated) values from the workbook, so the file must have been
saved by Excel at least once. Re-run this whenever the weekly data imports in
the workbook are refreshed; your saved order edits in data/orders.json are kept
unless --reset-orders is passed.
"""
import json
import math
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

DEFAULT_EXCEL = Path(r"S:\Buying\BUDGETS & FORECASTING\Buying 2026 Planning Folder\2026 Tradeplan - V2 - Forecast Unlinked.xlsx")
DEFAULT_EXCEL_BY_YEAR = {
    2026: DEFAULT_EXCEL,
    2025: Path(r"S:\Buying\BUDGETS & FORECASTING\Buying 2025 Planning Folder\Tradeplan\2025 Tradeplan - Working Document.xlsx"),
}
OUT_DIR = Path(__file__).resolve().parent / "data"
WEEKS = 53
WCOL0 = 10          # Suppliers sheet: week 1 = column K (0-based index 10)
SHEET_W = 96        # Suppliers sheet width (col CR)

METRIC_ROWS = ["LY Sales", "Actual Sales", "Sales Forecast", "Sales Value",
               "Order Forecast", "Order FOB", "CBM", "Stock Closing",
               "Stock Holding Value", "Weeks Cover"]


def norm_status(*vals):
    """Normalise catalog status to Live / Not Live / Unknown (authoritative
    source first). Live = plan for rebuy; Not Live = discontinued, run down."""
    for v in vals:
        if isinstance(v, str):
            u = v.strip().upper()
            if u == "LIVE":
                return "Live"
            if u in ("NOT LIVE", "NOTLIVE"):
                return "Not Live"
    return "Unknown"


def num(v, default=0.0):
    if isinstance(v, bool):
        return default
    if isinstance(v, (int, float)):
        return float(v)
    return default


def weekvals(row, start=WCOL0):
    return [num(v) for v in row[start:start + WEEKS]]


def sheet_rows(ws, width):
    rows = []
    for r in ws.iter_rows(values_only=True):
        r = list(r)
        if len(r) < width:
            r += [None] * (width - len(r))
        rows.append(r)
    return rows


def parse_suppliers_sheet(rows):
    suppliers, skus = [], {}
    sup_index = {}
    card = {}
    cached = {}  # sku_id -> cached rows for verification
    i = 0
    while i < len(rows):
        row = rows[i]
        label = row[2] if isinstance(row[2], str) else None
        if label == "Supplier number:":
            card["number"] = row[5]
        elif label == "Contact:":
            card["contact"] = row[5]
        elif label == "Port:":
            card["port"] = row[5]
        elif label == "Email:":
            card["email"] = str(row[5]) if row[5] is not None else None

        if row[9] == "LY Sales" and row[2] and row[1]:
            block = rows[i:i + 10]
            sup_name = str(row[1]).strip()
            if sup_name not in sup_index:
                sup_index[sup_name] = len(suppliers)
                suppliers.append({"name": sup_name,
                                  "number": card.get("number"),
                                  "contact": card.get("contact"),
                                  "port": card.get("port"),
                                  "email": card.get("email"),
                                  "origin": None})
                card = {}
            code = str(row[2]).strip()
            sku_id = code
            n = 2
            while sku_id in skus:
                sku_id = f"{code}#{n}"
                n += 1
            skus[sku_id] = {
                "id": sku_id,
                "code": code,
                "supplier": sup_name,
                "category": block[0][4] if isinstance(block[0][4], str) else None,
                "status": block[0][8] if isinstance(block[0][8], str) else None,
                "fob": num(block[1][6]),
                "landed": num(block[3][6]),
                "asp": num(block[5][6]),
                "cbm": num(block[7][6]),
                "stock_now": num(block[7][7]),
                "ly": weekvals(block[0]),
                "actual": weekvals(block[1]),
                "orders": weekvals(block[4]),
            }
            cached[sku_id] = {
                "forecast": weekvals(block[2]),
                "sales_value": weekvals(block[3]),
                "stock": weekvals(block[7]),
                "cover": weekvals(block[9]),
            }
            i += 10
            continue
        i += 1
    return suppliers, skus, cached


def parse_suppliers_2025(rows):
    """2025 layout: left block A-L, metric labels in col M (idx 12), weekly data
    from col N (idx 13). Reference data (name/FOB/landed/ASP/CBM) on the block's
    6th row; free stock in col J. Returns cached metric rows too."""
    W0 = 13
    def wv(row):
        return [num(v) for v in row[W0:W0 + WEEKS]]
    suppliers, skus, cached = [], {}, {}
    sup_index, card = {}, {}
    i = 0
    while i < len(rows):
        row = rows[i]
        label = row[2] if isinstance(row[2], str) else None
        if label == "Supplier number:":
            card["number"] = row[3]
        elif label == "Contact:":
            card["contact"] = row[3]
        elif label == "Port:":
            card["port"] = row[3]
        elif label == "Email:":
            card["email"] = str(row[3]) if row[3] is not None else None

        if len(row) > 12 and row[12] == "LY Sales" and row[2] and row[1]:
            block = rows[i:i + 10]
            sup_name = str(row[1]).strip()
            if sup_name not in sup_index:
                sup_index[sup_name] = len(suppliers)
                suppliers.append({"name": sup_name, "number": card.get("number"),
                                  "contact": card.get("contact"), "port": card.get("port"),
                                  "email": card.get("email"), "origin": None})
                card = {}
            code = str(row[2]).strip()
            sku_id = code
            n = 2
            while sku_id in skus:
                sku_id = f"{code}#{n}"; n += 1
            ref = block[5]   # row carrying name/FOB/landed/ASP/CBM
            skus[sku_id] = {
                "id": sku_id, "code": code, "supplier": sup_name,
                "category": None,
                "status": block[0][11] if isinstance(block[0][11], str) else None,
                "name_x": str(ref[3]) if ref[3] is not None else "",
                "fob": num(ref[5]), "landed": num(ref[6]), "asp": num(ref[7]), "cbm": num(ref[8]),
                "stock_now": num(block[7][9]),       # free stock value (col J on Stock Closing row)
                "ly": wv(block[0]), "actual": wv(block[1]), "orders": wv(block[4]),
                "sales_forecast": wv(block[2]), "running_stock_x": wv(block[7]),
            }
            cached[sku_id] = {
                "forecast": wv(block[2]), "sales_value": wv(block[3]),
                "stock": wv(block[7]), "cover": wv(block[9]),
            }
            i += 10
            continue
        i += 1
    return suppliers, skus, cached


def sku_week_table(ws, max_col=54):
    """Sheets shaped SKU in col A, weeks 1..53 in cols B..BB."""
    out = {}
    for r in ws.iter_rows(values_only=True):
        r = list(r)
        if len(r) < max_col:
            r += [None] * (max_col - len(r))
        key = r[0]
        if isinstance(key, str) and key.strip() and key.strip().lower() != "sku":
            out[key.strip()] = [num(v) for v in r[1:1 + WEEKS]]
    return out


# ---------------------------------------------------------------------------
# Calculation engine (mirror of the Excel formulas; same logic as static/app.js)
# ---------------------------------------------------------------------------
def compute_sku(sku, base_forecast, running_stock, cur_week, multiplier):
    fc = [round(base_forecast[w] * multiplier) for w in range(WEEKS)]
    stock = [0.0] * WEEKS
    for w in range(WEEKS):
        wk = w + 1
        if wk == 1 and cur_week != 1:
            stock[w] = running_stock[w]
        elif wk == cur_week:
            stock[w] = sku["orders"][w] + sku["stock_now"]
        elif wk < cur_week:
            stock[w] = running_stock[w]
        else:
            stock[w] = max(0.0, stock[w - 1] + sku["orders"][w] - fc[w])
    value = [0.0] * WEEKS
    for w in range(WEEKS):
        if (w + 1) < cur_week:
            value[w] = sku["actual"][w] * sku["asp"]
        else:
            value[w] = min(fc[w], stock[w]) * sku["asp"]
    cover = [0.0] * WEEKS
    for w in range(WEEKS):
        seg = fc[w:w + 4]
        avg = sum(seg) / len(seg)
        cover[w] = stock[w] / avg if avg else stock[w]
    return fc, stock, value, cover


def verify(skus, cached, base, runstock, cur_week, multiplier):
    stats = {k: [0, 0] for k in ("forecast", "stock", "sales_value", "cover")}
    examples = []
    for sid, sku in skus.items():
        bf = base[sid]
        rs = runstock.get(sku["code"], [0.0] * WEEKS)
        fc, stock, value, cover = compute_sku(sku, bf, rs, cur_week, multiplier)
        c = cached[sid]
        for name, mine, tol in (("forecast", fc, 0.51), ("stock", stock, 0.51),
                                ("sales_value", value, 1.5), ("cover", cover, 0.05)):
            for w in range(WEEKS):
                stats[name][1] += 1
                if abs(mine[w] - c[name][w]) > tol:
                    stats[name][0] += 1
                    if len(examples) < 12:
                        examples.append(f"  {sid} {name} wk{w+1}: app={mine[w]:.2f} excel={c[name][w]:.2f}")
    return stats, examples


def build_2026(wb, src, year):
    sup_rows = sheet_rows(wb["Suppliers"], SHEET_W)
    suppliers, skus, cached = parse_suppliers_sheet(sup_rows)
    print(f"  {len(suppliers)} suppliers, {len(skus)} SKU lines")

    runstock = sku_week_table(wb["Running Stock"])

    vs_rows = sheet_rows(wb["Value Splits"], 59)
    origins = {}
    for r in vs_rows[3:445]:
        if isinstance(r[0], str) and r[0].strip() and r[2] is None and isinstance(r[1], str):
            origins[r[0].strip()] = r[1].strip()
    for s in suppliers:
        s["origin"] = origins.get(s["name"])
    multiplier = num(vs_rows[493][15], 1.0) or 1.0   # Value Splits P494
    history = {"quarters": {
        "2023": [num(vs_rows[i][5]) for i in range(499, 503)],
        "2024": [num(vs_rows[i][6]) for i in range(499, 503)],
    }}
    capacities = {
        "racking_websa": num(vs_rows[477][3], 3960),
        "racking_express": num(vs_rows[478][3], 2848),
        "racking_refit": num(vs_rows[479][3], 2184),
        "stillage_websa": num(vs_rows[484][3], 912),
        "stillage_lough": num(vs_rows[485][3], 243),
        "stillage_express": num(vs_rows[486][3], 0),
        "stillage_free": num(vs_rows[487][3], 0),
    }

    # Product names + catalog status from the Buying Report (the authoritative
    # source; col B = name, col D = Catalog Status).
    names = {}
    catalog = {}
    seasons = {}
    for r in wb["Buying Report"].iter_rows(min_row=2, values_only=True):
        r = list(r) + [None] * 6
        if isinstance(r[0], str) and r[0].strip():
            code = r[0].strip()
            names.setdefault(code, str(r[1]) if r[1] is not None else "")
            catalog.setdefault(code, r[3])
            seasons.setdefault(code, str(r[5]).strip() if r[5] is not None else "")

    # Pallet config (TIxHI): full pallet qty + pallet type per SKU
    pallets = {}
    for r in wb["TIxHI"].iter_rows(values_only=True):
        r = list(r) + [None] * 20
        if isinstance(r[0], str) and r[0].strip() and r[0].strip() != "Product":
            fpq = num(r[7])
            ptype = r[9] if isinstance(r[9], str) else None
            if r[0].strip() not in pallets:
                pallets[r[0].strip()] = {"fpq": fpq, "type": ptype}

    # Product images: first image per SKU, prefer type 'Product Image'
    images = {}
    for r in wb["Product Images"].iter_rows(min_row=2, values_only=True):
        r = list(r) + [None] * 8
        skucode, itype, url = r[1], r[3], r[7]
        if not isinstance(skucode, str) or not isinstance(url, str):
            continue
        cur = images.get(skucode)
        if cur is None or (itype == "Product Image" and cur[0] != "Product Image"):
            images[skucode] = (itype, url)

    cur_week_cached = sup_rows[2][9]
    cur_week = int(cur_week_cached) if isinstance(cur_week_cached, (int, float)) else 1
    print(f"  workbook saved at week {cur_week}, multiplier {multiplier}")

    # Base forecast = the Suppliers sheet's own (cached) Sales Forecast row,
    # un-multiplied. This captures the ~28 SKUs whose forecast cells were
    # manually overridden rather than linked to 'Base Forecast Units'.
    base = {sid: [v / multiplier for v in cached[sid]["forecast"]] for sid in skus}

    print("Verifying calculation engine against Excel's cached results ...")
    stats, examples = verify(skus, cached, base, runstock, cur_week, multiplier)
    for name, (bad, total) in stats.items():
        pct = 100.0 * (1 - bad / total) if total else 100.0
        print(f"  {name:<12} {total - bad}/{total} match ({pct:.2f}%)")
    if examples:
        print("  sample mismatches:")
        print("\n".join(examples))

    for sid, sku in skus.items():
        sku["name"] = names.get(sku["code"], "")
        sku["status"] = norm_status(catalog.get(sku["code"]), sku.get("status"))
        sku["season"] = seasons.get(sku["code"], "") or "No Defined Season"
        sku["base_forecast"] = base[sid]
        sku["running_stock"] = runstock.get(sku["code"], [0.0] * WEEKS)
        p = pallets.get(sku["code"], {})
        sku["fpq"] = p.get("fpq", 0)
        sku["pallet_type"] = p.get("type")
        img = images.get(sku["code"])
        sku["image"] = img[1] if img else None

    orders = {sid: sku.pop("orders") for sid, sku in skus.items()}

    master = {
        "generated": datetime.now().isoformat(timespec="seconds"),
        "source": str(src), "year": year,
        "weeks": WEEKS, "week1_start": "2025-12-29", "data_week": cur_week,
        "suppliers": suppliers, "skus": list(skus.values()),
        "history": history, "capacities": capacities,
        "container_cbm": 68, "multiplier": multiplier,
    }
    return master, orders


def build_2025(wb, src, year):
    sup_rows = sheet_rows(wb["Suppliers"], 70)
    suppliers, skus, cached = parse_suppliers_2025(sup_rows)
    print(f"  {len(suppliers)} suppliers, {len(skus)} SKU lines")

    names, catalog, seasons = {}, {}, {}
    for r in wb["Buying Report"].iter_rows(min_row=2, values_only=True):
        r = list(r) + [None] * 6
        if isinstance(r[0], str) and r[0].strip():
            code = r[0].strip()
            names.setdefault(code, str(r[1]) if r[1] is not None else "")
            catalog.setdefault(code, r[3])
            seasons.setdefault(code, str(r[5]).strip() if r[5] is not None else "")

    pallets = {}
    if "TIxHI" in wb.sheetnames:
        for r in wb["TIxHI"].iter_rows(values_only=True):
            r = list(r) + [None] * 20
            if isinstance(r[0], str) and r[0].strip() and r[0].strip() != "Product":
                pallets.setdefault(r[0].strip(), {"fpq": num(r[7]), "type": r[9] if isinstance(r[9], str) else None})

    origins = {}
    if "Value Splits" in wb.sheetnames:
        for r in sheet_rows(wb["Value Splits"], 59)[3:460]:
            if isinstance(r[0], str) and r[0].strip() and r[2] is None and isinstance(r[1], str):
                origins[r[0].strip()] = r[1].strip()

    cur_week = WEEKS   # completed/historical year → all weeks show actuals
    for sid, sku in skus.items():
        sku["name"] = names.get(sku["code"]) or sku.pop("name_x", "")
        sku.pop("name_x", None)
        sku["status"] = norm_status(catalog.get(sku["code"]), sku.get("status"))
        sku["season"] = seasons.get(sku["code"], "") or "No Defined Season"
        sku["base_forecast"] = sku.pop("sales_forecast")
        sku["running_stock"] = sku.pop("running_stock_x")
        p = pallets.get(sku["code"], {})
        sku["fpq"] = p.get("fpq", 0)
        sku["pallet_type"] = p.get("type")
        sku["image"] = None
    for s in suppliers:
        s["origin"] = origins.get(s["name"])

    orders = {sid: sku.pop("orders") for sid, sku in skus.items()}
    capacities = {"racking_websa": 3960, "racking_express": 2848, "racking_refit": 2184,
                  "stillage_websa": 912, "stillage_lough": 243, "stillage_express": 0, "stillage_free": 0}
    master = {
        "generated": datetime.now().isoformat(timespec="seconds"),
        "source": str(src), "year": year,
        "weeks": WEEKS, "week1_start": "2024-12-30", "data_week": cur_week,
        "suppliers": suppliers, "skus": list(skus.values()),
        "history": {"quarters": {str(year - 3): [0, 0, 0, 0], str(year - 2): [0, 0, 0, 0]}},
        "capacities": capacities, "container_cbm": 68, "multiplier": 1.0,
    }
    return master, orders


def main():
    argv = sys.argv[1:]
    reset_orders = "--reset-orders" in argv
    year, positionals, i = 2026, [], 0
    while i < len(argv):
        a = argv[i]
        if a == "--year":
            year = int(argv[i + 1]); i += 2; continue
        if a == "--reset-orders":
            i += 1; continue
        positionals.append(a); i += 1
    src = Path(positionals[0]) if positionals else DEFAULT_EXCEL_BY_YEAR.get(year, DEFAULT_EXCEL)
    out = OUT_DIR / str(year)
    out.mkdir(parents=True, exist_ok=True)
    print(f"Reading {src} ... (year {year})")
    wb = load_workbook(src, read_only=True, data_only=True)

    master, orders = (build_2025 if year <= 2025 else build_2026)(wb, src, year)

    (out / "master.json").write_text(json.dumps(master), encoding="utf-8")
    (out / "orders_imported.json").write_text(json.dumps(orders), encoding="utf-8")
    if reset_orders or not (out / "orders.json").exists():
        (out / "orders.json").write_text(json.dumps(orders), encoding="utf-8")
        print(f"  {out / 'orders.json'} written from Excel order quantities")
    else:
        print(f"  {out / 'orders.json'} kept (your edits preserved); use --reset-orders to overwrite")
    if not (OUT_DIR / "settings.json").exists():
        (OUT_DIR / "settings.json").write_text(json.dumps(
            {"multiplier": master.get("multiplier", 1), "container_cbm": 68,
             "capacities": master["capacities"]}), encoding="utf-8")
    years = sorted(p.name for p in OUT_DIR.iterdir() if p.is_dir() and p.name.isdigit())
    (OUT_DIR / "years.json").write_text(json.dumps(
        {"years": years, "default": max(years) if years else str(year)}), encoding="utf-8")
    print(f"Done. Year {year} written to {out}. Years available: {years}")


if __name__ == "__main__":
    main()
