"""Local server for the 2026 Tradeplan app. No external dependencies.

Run:  python server.py        (or double-click run.bat)
Then open http://localhost:8765 (opens automatically).
"""
import json
import math
import re
import threading
import time
import urllib.request
import webbrowser
from datetime import date, datetime
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse, parse_qs, unquote

import supplier_form

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
PORT = 8765
WEEK1_START = date(2025, 12, 29)
_weather_cache = {}   # (lat,lon) -> (timestamp, payload)


def safe_filename(name):
    return re.sub(r'[\\/:*?"<>|]+', " ", str(name)).strip() or "supplier"


def years_index():
    idx = read_json(DATA / "years.json", None)
    if idx and idx.get("years"):
        return idx["years"], str(idx.get("default") or idx["years"][-1])
    yrs = sorted(p.name for p in DATA.iterdir() if p.is_dir() and p.name.isdigit())
    return yrs, (yrs[-1] if yrs else None)


def year_dir(qs):
    """Resolve the data folder for the requested ?year=, falling back to default."""
    yrs, default = years_index()
    y = (qs.get("year") or [default])[0] if qs else default
    if y not in yrs:
        y = default
    return (DATA / str(y)) if y else DATA, y


def planning_week(d):
    return (d - WEEK1_START).days // 7 + 1


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _isodate(v):
    """Normalise a spreadsheet date cell to an ISO 'YYYY-MM-DD' string, or None."""
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, str):
        s = v.strip()
        if s in ("", "-", "N/A", "n/a", "TBC", "tbc"):
            return None
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y"):
            try:
                return datetime.strptime(s, fmt).date().isoformat()
            except ValueError:
                pass
    return None


# Container 'Order' cells are messy: several PO numbers per cell, inconsistent spacing,
# typos (POI→PO, P0→PO, missing 'O'), concatenated 'PO..PO..', and non-PO charter refs
# (CHN-/CHA/LOT). Returns (sorted distinct PO numbers, has_etc) — has_etc flags an
# incomplete list ('etc'/'more'/'+'/'…') so a consolidated container can be reported as
# possibly referencing more POs than listed.
def po_tokens(order_str):
    s = str(order_str).upper()
    has_etc = bool(re.search(r"\bETC\b|\bMORE\b|\.\.\.|…|\+", s))
    s = s.replace("POI", "PO")                       # POI057193 → PO057193
    found = set(re.findall(r"PO\d{4,7}", s))         # also splits concatenated PO..PO..
    # missing-O variant e.g. 'P059604' (a P followed by 5-6 digits, not part of a PO token)
    for m in re.findall(r"(?<![A-Z0-9])P(\d{5,6})(?![\d])", s):
        found.add("PO" + m)
    return sorted(found), has_etc


def seasonal_normal(doy):
    # crude UK daily-mean climatology: ~4°C late Jan, ~17°C mid-July
    return 10.5 - 6.5 * math.cos(2 * math.pi * (doy - 15) / 365)


def fetch_weather(lat, lon):
    key = (round(lat, 2), round(lon, 2))
    hit = _weather_cache.get(key)
    if hit and (time.time() - hit[0]) < 3 * 3600:
        return hit[1]
    url = (f"https://api.open-meteo.com/v1/forecast?latitude={lat}&longitude={lon}"
           "&daily=temperature_2m_max,temperature_2m_min&forecast_days=16&timezone=auto")
    try:
        with urllib.request.urlopen(url, timeout=8) as resp:
            raw = json.loads(resp.read().decode("utf-8"))
        daily = raw["daily"]
        agg = {}   # planning week -> [temps], [normals]
        for ds, tmax, tmin in zip(daily["time"], daily["temperature_2m_max"], daily["temperature_2m_min"]):
            if tmax is None or tmin is None:
                continue
            d = datetime.strptime(ds, "%Y-%m-%d").date()
            wk = planning_week(d)
            mean = (tmax + tmin) / 2.0
            agg.setdefault(wk, [[], []])
            agg[wk][0].append(mean)
            agg[wk][1].append(seasonal_normal(d.timetuple().tm_yday))
        weeks = {}
        for wk, (temps, normals) in agg.items():
            t = sum(temps) / len(temps)
            n = sum(normals) / len(normals)
            weeks[wk] = {"temp": round(t, 1), "normal": round(n, 1), "anomaly": round(t - n, 1)}
        payload = {"ok": True, "weeks": weeks, "lat": lat, "lon": lon,
                   "generated": datetime.now().isoformat(timespec="minutes")}
    except Exception as exc:
        payload = {"ok": False, "error": str(exc)}
    _weather_cache[key] = (time.time(), payload)
    return payload


def read_json(path, default):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return default


def supplier_form_cfg():
    """Lead-time / payment assumptions the supplier order-planning form embeds,
    set in the app under Settings → Supplier forms (global, all years)."""
    return read_json(DATA / "settings.json", {}).get("supplier_form") or {}


def with_imported_at(path):
    """Read a JSON dict and guarantee an 'importedAt' field — backfilled from the file's
    modification time for uploads made before we started stamping them."""
    data = read_json(path, None)
    if isinstance(data, dict) and not data.get("importedAt"):
        try:
            data["importedAt"] = datetime.fromtimestamp(path.stat().st_mtime).isoformat(timespec="seconds")
        except OSError:
            pass
    return data


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT / "static"), **kwargs)

    def log_message(self, fmt, *args):
        pass  # keep the console quiet

    def end_headers(self):
        # never cache app files, so a refresh always shows the latest version
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def send_json(self, obj, status=200):
        body = json.dumps(obj).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        parsed = urlparse(self.path)
        qs = parse_qs(parsed.query)
        if parsed.path == "/api/years":
            yrs, default = years_index()
            self.send_json({"years": yrs, "default": default})
        elif parsed.path == "/api/data":
            ydir, year = year_dir(qs)
            master = read_json(ydir / "master.json", None)
            if master is None:
                self.send_json({"error": f"data for year {year} missing - run import_data.py --year {year}"}, 500)
                return
            self.send_json({
                "year": year,
                "master": master,
                "orders": read_json(ydir / "orders.json", {}),
                "settings": read_json(DATA / "settings.json", {}),
                # saved per-year proposed-rebuy layer (null if this year was never
                # built yet → client auto-builds; {} means the user cleared it)
                "proposed": read_json(ydir / "proposed.json", None),
                # prior-year-end stock this (forecast) year was last chained to
                "stockbase": read_json(ydir / "stockbase.json", None),
                # global PO ↔ container linking data (shared across years; null until uploaded)
                "poWebsa": with_imported_at(DATA / "po_websa.json"),
                "poContainers": with_imported_at(DATA / "po_containers.json"),
            })
        elif parsed.path == "/api/imported-orders":
            ydir, _ = year_dir(qs)
            self.send_json(read_json(ydir / "orders_imported.json", {}))
        elif parsed.path == "/api/export-supplier":
            self.export_supplier(parse_qs(parsed.query))
        elif parsed.path == "/api/configs":
            self.list_configs()
        elif parsed.path == "/api/weather":
            q = parse_qs(parsed.query)
            try:
                lat = float((q.get("lat") or ["52.77"])[0])
                lon = float((q.get("lon") or ["-1.21"])[0])
            except ValueError:
                self.send_json({"ok": False, "error": "bad coordinates"}, 400)
                return
            self.send_json(fetch_weather(lat, lon))
        else:
            super().do_GET()

    def export_supplier(self, qs):
        name = unquote((qs.get("name") or [""])[0])
        if not name:
            self.send_json({"error": "missing supplier name"}, 400)
            return
        ydir, year = year_dir(qs)
        try:
            bio = supplier_form.export_supplier(name, data_dir=ydir, year=year, form_cfg=supplier_form_cfg())
            data = bio.getvalue()
        except Exception as exc:
            self.send_json({"error": f"export failed: {exc}"}, 500)
            return
        fname = f"{safe_filename(name)} - {year} Order Planning.xlsx"
        self.send_response(200)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", f'attachment; filename="{fname}"')
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def export_suppliers_zip(self, qs):
        """Build one supplier form per requested name and return them as a single
        .zip (a 'folder download'), so multiple suppliers don't trigger many
        separate file downloads. An empty/missing name list means every supplier."""
        import io
        import zipfile
        ydir, year = year_dir(qs)
        try:
            length = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(length)) if length else {}
        except (ValueError, json.JSONDecodeError):
            body = {}
        names = body.get("names") or []
        master = read_json(ydir / "master.json", {"skus": []})
        if not names:  # default: all suppliers, in workbook order
            names = list(dict.fromkeys(s["supplier"] for s in master["skus"]))
        buf = io.BytesIO()
        used = set()
        fcfg = supplier_form_cfg()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for name in names:
                try:
                    bio = supplier_form.export_supplier(name, data_dir=ydir, year=year, form_cfg=fcfg)
                except Exception:
                    continue  # skip a supplier that fails rather than break the whole zip
                fname = f"{safe_filename(name)} - {year} Order Planning.xlsx"
                n, base = 2, fname
                while fname in used:  # avoid collisions after filename sanitising
                    fname = base[:-5] + f" ({n})" + ".xlsx"
                    n += 1
                used.add(fname)
                zf.writestr(fname, bio.getvalue())
        data = buf.getvalue()
        zipname = f"{year} Order Planning forms.zip"
        self.send_response(200)
        self.send_header("Content-Type", "application/zip")
        self.send_header("Content-Disposition", f'attachment; filename="{zipname}"')
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def export_forecast(self, qs):
        """Return an .xlsx of the client-computed weekly sales-unit forecast
        (SKU in col A, W1..W53 across, Total at the end)."""
        ydir, year = year_dir(qs)
        try:
            length = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(length)) if length else {}
        except (ValueError, json.JSONDecodeError):
            body = {}
        try:
            bio = supplier_form.export_forecast(body.get("rows") or [], year=year, week1=body.get("week1"))
            data = bio.getvalue()
        except Exception as exc:
            self.send_json({"error": f"export failed: {exc}"}, 500)
            return
        fname = f"{year} Sales Unit Forecast.xlsx"
        self.send_response(200)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", f'attachment; filename="{fname}"')
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def export_arrivals(self):
        """Return the Arrivals page as an .xlsx — the client sends its already-joined
        upcoming-container rows (respecting the on-screen filter)."""
        try:
            length = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(length)) if length else {}
        except (ValueError, json.JSONDecodeError):
            body = {}
        try:
            bio = supplier_form.export_arrivals(body)
            data = bio.getvalue()
        except Exception as exc:
            self.send_json({"error": f"export failed: {exc}"}, 500)
            return
        self.send_response(200)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", 'attachment; filename="Upcoming Containers.xlsx"')
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/save":
            try:
                ydir, _ = year_dir(parse_qs(parsed.query))
                length = int(self.headers.get("Content-Length", 0))
                body = json.loads(self.rfile.read(length))
                if "orders" in body:
                    (ydir / "orders.json").write_text(json.dumps(body["orders"]), encoding="utf-8")
                if "settings" in body:   # settings are global across years
                    (DATA / "settings.json").write_text(json.dumps(body["settings"]), encoding="utf-8")
                if "proposed" in body:   # editable rebuy suggestions (per year), for export merge
                    (ydir / "proposed.json").write_text(json.dumps(body["proposed"]), encoding="utf-8")
                if "stockbase" in body:  # prior-year-end stock a forecast year was last chained to
                    (ydir / "stockbase.json").write_text(json.dumps(body["stockbase"]), encoding="utf-8")
                self.send_json({"ok": True})
            except Exception as exc:  # report save failures to the UI
                self.send_json({"ok": False, "error": str(exc)}, 500)
        elif parsed.path == "/api/export-suppliers":
            self.export_suppliers_zip(parse_qs(parsed.query))
        elif parsed.path == "/api/export-forecast":
            self.export_forecast(parse_qs(parsed.query))
        elif parsed.path == "/api/export-arrivals":
            self.export_arrivals()
        elif parsed.path == "/api/import-supplier":
            self.import_supplier(parse_qs(parsed.query))
        elif parsed.path == "/api/parse-asp":
            self.parse_asp()
        elif parsed.path == "/api/apply-asp":
            self.apply_asp(parse_qs(parsed.query))
        elif parsed.path == "/api/parse-landed":
            self.parse_landed()
        elif parsed.path == "/api/apply-landed":
            self.apply_landed(parse_qs(parsed.query))
        elif parsed.path == "/api/parse-buying":
            self.parse_buying()
        elif parsed.path == "/api/apply-buying":
            self.apply_buying(parse_qs(parsed.query))
        elif parsed.path == "/api/parse-wksales":
            self.parse_wksales()
        elif parsed.path == "/api/apply-wksales":
            self.apply_wksales(parse_qs(parsed.query))
        elif parsed.path == "/api/parse-duty":
            self.parse_duty()
        elif parsed.path == "/api/apply-duty":
            self.apply_duty(parse_qs(parsed.query))
        elif parsed.path == "/api/parse-po":
            self.parse_po()
        elif parsed.path == "/api/parse-containers":
            self.parse_containers()
        elif parsed.path == "/api/save-config":
            self.save_config(parse_qs(parsed.query))
        elif parsed.path == "/api/load-config":
            self.load_config(parse_qs(parsed.query))
        elif parsed.path == "/api/revert-original":
            self.revert_original(parse_qs(parsed.query))
        elif parsed.path == "/api/save-year":
            self.save_year(parse_qs(parsed.query))
        else:
            self.send_json({"error": "not found"}, 404)

    def list_configs(self):
        """List the named saved configurations (newest first), with the years each covers."""
        cdir = DATA / "configs"
        out = []
        if cdir.is_dir():
            for f in cdir.glob("*.json"):
                d = read_json(f, None)
                if d:
                    out.append({"name": d.get("name", f.stem), "file": f.stem,
                                "saved_at": d.get("saved_at"),
                                "years": sorted((d.get("years") or {}).keys())})
        out.sort(key=lambda x: x.get("saved_at") or "", reverse=True)
        self.send_json({"configs": out})

    def save_config(self, qs):
        """Save a configuration under a name into data/configs/. scope=all snapshots
        every year + the global forecast settings; scope=<year> snapshots just that
        year's orders + proposed rebuys (plus settings, for context)."""
        name = unquote((qs.get("name") or [""])[0]).strip()
        scope = (qs.get("scope") or ["all"])[0]
        if not name:
            self.send_json({"ok": False, "error": "missing name"}, 400)
            return
        try:
            cdir = DATA / "configs"
            cdir.mkdir(parents=True, exist_ok=True)
            years = [p.name for p in DATA.iterdir() if p.is_dir() and p.name.isdigit()]
            if scope != "all":
                years = [scope] if scope in years else []
            snapshot = {
                "name": name,
                "saved_at": datetime.now().isoformat(timespec="seconds"),
                "scope": scope,
                "settings": read_json(DATA / "settings.json", {}),
                "years": {},
            }
            for y in years:
                p = DATA / y
                snapshot["years"][y] = {
                    "orders": read_json(p / "orders.json", {}),
                    "proposed": read_json(p / "proposed.json", {}),
                }
            stem = safe_filename(name)
            (cdir / (stem + ".json")).write_text(json.dumps(snapshot), encoding="utf-8")
            self.send_json({"ok": True, "name": name, "file": stem, "saved_at": snapshot["saved_at"]})
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, 500)

    def load_config(self, qs):
        """Restore a named configuration to the working files (client reloads after).
        scope=all restores every year in the config + the global settings; scope=<year>
        restores just that year's orders + proposed (leaving global settings untouched)."""
        stem = safe_filename(unquote((qs.get("file") or qs.get("name") or [""])[0]).strip())
        scope = (qs.get("scope") or ["all"])[0]
        d = read_json(DATA / "configs" / (stem + ".json"), None) if stem else None
        if not d:
            self.send_json({"ok": False, "error": "configuration not found"}, 404)
            return
        try:
            yrs = d.get("years") or {}
            targets = list(yrs.keys()) if scope == "all" else ([scope] if scope in yrs else [])
            if scope == "all" and isinstance(d.get("settings"), dict):
                (DATA / "settings.json").write_text(json.dumps(d["settings"]), encoding="utf-8")
            for y in targets:
                ydir = DATA / str(y)
                if not ydir.is_dir():
                    continue
                yd = yrs[y]
                (ydir / "orders.json").write_text(json.dumps(yd.get("orders", {})), encoding="utf-8")
                (ydir / "proposed.json").write_text(json.dumps(yd.get("proposed", {})), encoding="utf-8")
                if (ydir / "stockbase.json").exists():   # let forecast years re-chain cleanly
                    (ydir / "stockbase.json").unlink()
            self.send_json({"ok": True, "name": d.get("name", stem), "restored": targets})
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, 500)

    def revert_original(self, qs):
        """Reset orders to the originally-imported figures and clear rebuys. scope=all
        does every year; scope=<year> does just that year."""
        ydir, year = year_dir(qs)
        scope = (qs.get("scope") or [year])[0]
        years = [p.name for p in DATA.iterdir() if p.is_dir() and p.name.isdigit()]
        targets = years if scope == "all" else [scope if scope in years else year]
        try:
            for y in targets:
                yd = DATA / str(y)
                imp = read_json(yd / "orders_imported.json", {})
                (yd / "orders.json").write_text(json.dumps(imp), encoding="utf-8")
                (yd / "proposed.json").write_text("{}", encoding="utf-8")
                if (yd / "stockbase.json").exists():
                    (yd / "stockbase.json").unlink()
            self.send_json({"ok": True, "restored": targets})
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, 500)

    def save_year(self, qs):
        """Persist a client-generated forecast year as its own data folder."""
        year = (qs.get("year") or [""])[0]
        if not year.isdigit():
            self.send_json({"ok": False, "error": "bad year"}, 400)
            return
        try:
            length = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(length))
            ydir = DATA / year
            ydir.mkdir(parents=True, exist_ok=True)
            (ydir / "master.json").write_text(json.dumps(body["master"]), encoding="utf-8")
            (ydir / "orders.json").write_text(json.dumps(body.get("orders", {})), encoding="utf-8")
            (ydir / "orders_imported.json").write_text(json.dumps(body.get("orders", {})), encoding="utf-8")
            years = sorted(p.name for p in DATA.iterdir() if p.is_dir() and p.name.isdigit())
            (DATA / "years.json").write_text(json.dumps({"years": years, "default": years[-1] if years else year}), encoding="utf-8")
            self.send_json({"ok": True, "year": year, "years": years})
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, 500)

    def import_supplier(self, qs):
        import io
        name = unquote((qs.get("name") or [""])[0])
        ydir, _ = year_dir(qs)
        try:
            length = int(self.headers.get("Content-Length", 0))
            raw = self.rfile.read(length)
            parsed = supplier_form.parse_supplier_form(io.BytesIO(raw))
        except Exception as exc:
            self.send_json({"ok": False, "error": f"could not read form: {exc}"}, 400)
            return
        master = read_json(ydir / "master.json", {"skus": []})
        orders = read_json(ydir / "orders.json", {})
        # map SKU code -> id, preferring the named supplier when given
        by_code = {}
        for s in master["skus"]:
            if name and s["supplier"] != name:
                continue
            by_code.setdefault(s["code"], s["id"])
        if not name:  # no supplier hint: map across everything
            for s in master["skus"]:
                by_code.setdefault(s["code"], s["id"])

        updated, unmatched, weeks = [], [], set()
        for code, wkmap in parsed["orders"].items():
            sid = by_code.get(code)
            if not sid:
                unmatched.append(code)
                continue
            # mirror the form exactly: clear the SKU's whole schedule, then apply
            # the form's quantities. Any week not in the form (or 0/blank in it)
            # ends up blank.
            vec = [0] * supplier_form.WEEKS
            for wk, qty in wkmap.items():
                if 1 <= wk <= supplier_form.WEEKS:
                    vec[wk - 1] = qty
                    weeks.add(wk)
            orders[sid] = vec
            updated.append(code)

        (ydir / "orders.json").write_text(json.dumps(orders), encoding="utf-8")
        self.send_json({"ok": True, "updated": len(updated), "unmatched": unmatched,
                        "weeks": sorted(weeks), "supplier": name,
                        "total_units": sum(sum(v.values()) for v in parsed["orders"].values())})

    @staticmethod
    def aggregate_asp(raw):
        """Read a 'Product Sales by Account' export and return ({code->basic ASP},
        {code->top-channel ASP}, file_sku_count). basic = total value / total units across
        every channel (volume-weighted). top = the per-unit price at the single channel with
        the highest sales VALUE for that product — so a product that sells mostly on a
        higher-priced marketplace isn't dragged down by low-priced DSV channels (or vice
        versa). Only positive results kept."""
        import io
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
        ws = wb.active
        rows = ws.iter_rows(values_only=True)
        header = next(rows)
        def col(*names):
            want = {n.lower() for n in names}
            for i, h in enumerate(header):
                if isinstance(h, str) and h.strip().lower() in want:
                    return i
            return None
        ci, vi, qi = col("product sku", "sku"), col("sales value ty", "sales value"), col("sales qty ty", "sales qty")
        if ci is None or vi is None or qi is None:
            raise ValueError("Expected columns 'Product SKU', 'Sales Value TY' and 'Sales Qty TY' in the file.")
        tot = {}          # code -> [sumValue, sumQty]   (basic)
        top = {}          # code -> [bestValue, bestQty]  (the single highest-value channel)
        for row in rows:
            code = row[ci]
            if not code:
                continue
            code = str(code).strip()
            v = row[vi] if isinstance(row[vi], (int, float)) else 0
            q = row[qi] if isinstance(row[qi], (int, float)) else 0
            if q <= 0 or v <= 0:
                continue
            a = tot.setdefault(code, [0.0, 0.0]); a[0] += v; a[1] += q
            b = top.get(code)
            if b is None or v > b[0]:
                top[code] = [v, q]
        basic = {c: round(v / q, 2) for c, (v, q) in tot.items() if q > 0}
        topm = {c: round(v / q, 2) for c, (v, q) in top.items() if q > 0}
        return basic, topm, len(tot)

    def parse_asp(self):
        """Parse an uploaded sales file and return the per-SKU average selling price
        (no data is changed yet — the client previews then applies)."""
        try:
            length = int(self.headers.get("Content-Length", 0))
            raw = self.rfile.read(length)
            basic, top, file_skus = self.aggregate_asp(raw)
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, 400)
            return
        self.send_json({"ok": True, "basic": basic, "top": top,
                        "fileSkus": file_skus, "withSales": len(basic)})

    def apply_asp(self, qs):
        """Write new ASPs into master.json for an explicit list of years, matched by SKU
        code. Each updated product is tagged with how its price was set — src='upload'
        (from a sales file) or src='manual' (hand-edited) — so the UI can distinguish
        auto-updated, manually-set and not-yet-updated products."""
        try:
            length = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(length)) if length else {}
            asp = body.get("asp") or {}
            src = body.get("src") or "upload"
            yrs, _ = years_index()
            targets = [str(y) for y in (body.get("years") or []) if str(y) in yrs]
            applied = {}
            for y in targets:
                mpath = DATA / y / "master.json"
                master = read_json(mpath, None)
                if not master:
                    continue
                n = 0
                for s in master["skus"]:
                    a = asp.get(s.get("code"))
                    if a is not None and a > 0:
                        s["asp_prev"] = s.get("asp")   # keep the prior price for variance reporting
                        s["asp"] = a
                        s["asp_src"] = src
                        n += 1
                mpath.write_text(json.dumps(master), encoding="utf-8")
                applied[y] = n
            self.send_json({"ok": True, "applied": applied, "years": targets, "src": src})
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, 500)

    # ---- FOB / landed-cost update (from a "Landed Costs" export) ----
    @staticmethod
    def aggregate_landed(raw):
        """Read a Landed Cost export → {code: {fob, landed, fobLast, landedLast}}.
        Columns: A=code, C=average landed (GBP), D=last landed (GBP),
        E=last-receipted FOB (USD), H=current-outstanding FOB (USD)."""
        import io
        import warnings
        import openpyxl
        def num(x):
            return float(x) if isinstance(x, (int, float)) else None
        out, n = {}, 0
        # The source file's date columns (e.g. F) can hold Excel error/overflow serials that
        # openpyxl warns about on read. We never use those columns, so silence the noise —
        # otherwise every bad date cell floods the console.
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
            ws = wb.active
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                if i == 0 or not row or not row[0]:
                    continue                                   # header / blank
                code = str(row[0]).strip()
                if not code:
                    continue
                landed = num(row[2]) if len(row) > 2 else None       # C  average landed (GBP)
                landed_last = num(row[3]) if len(row) > 3 else None   # D  last landed (GBP)
                fob_last = num(row[4]) if len(row) > 4 else None      # E  last-receipted FOB (USD)
                fob = num(row[7]) if len(row) > 7 else None           # H  outstanding FOB (USD)
                rec = {}
                if fob and fob > 0: rec["fob"] = round(fob, 2)
                if landed and landed > 0: rec["landed"] = round(landed, 2)
                if fob_last and fob_last > 0: rec["fobLast"] = round(fob_last, 2)
                if landed_last and landed_last > 0: rec["landedLast"] = round(landed_last, 2)
                if rec:
                    out[code] = rec
                    n += 1
        return out, n

    def parse_landed(self):
        """Parse an uploaded Landed Cost file and return per-SKU FOB/landed costs
        (no data is changed yet — the client previews, allows overrides, then applies)."""
        try:
            length = int(self.headers.get("Content-Length", 0))
            landed, count = self.aggregate_landed(self.rfile.read(length))
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, 400)
            return
        self.send_json({"ok": True, "landed": landed, "fileSkus": count})

    def apply_landed(self, qs):
        """Write new FOB (USD) + landed (GBP) costs into master.json for the given years,
        matched by code. Each value is tagged src='upload' (from the file) or 'manual'
        (overridden in the preview) so the UI can colour it on the plan."""
        try:
            length = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(length)) if length else {}
            fob = body.get("fob") or {}
            landed = body.get("landed") or {}
            manual = set(body.get("manual") or [])
            yrs, _ = years_index()
            targets = [str(y) for y in (body.get("years") or []) if str(y) in yrs]
            applied = {}
            for y in targets:
                mpath = DATA / y / "master.json"
                master = read_json(mpath, None)
                if not master:
                    continue
                n = 0
                for s in master["skus"]:
                    c = s.get("code")
                    touched = False
                    if c in fob and fob[c] and fob[c] > 0:
                        s["fob_prev"] = s.get("fob")
                        s["fob"] = fob[c]
                        s["fob_src"] = "manual" if c in manual else "upload"
                        touched = True
                    if c in landed and landed[c] and landed[c] > 0:
                        s["landed_prev"] = s.get("landed")
                        s["landed"] = landed[c]
                        s["landed_src"] = "manual" if c in manual else "upload"
                        touched = True
                    if touched:
                        n += 1
                mpath.write_text(json.dumps(master), encoding="utf-8")
                applied[y] = n
            self.send_json({"ok": True, "applied": applied, "years": targets})
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, 500)

    # ---- catalogue status + outstanding purchases (from the "Buying Report") ----
    @staticmethod
    def aggregate_buying(raw):
        """Read a Buying Report → {code: {status, osPurchases}} for WEBSA rows only.
        Columns: B=product code, D=location (keep only 'WEBSA'), E=catalog status
        (LIVE/NOT LIVE → Live/Not Live), M=OS Purchases."""
        import io
        import warnings
        import openpyxl
        def norm_status(v):
            s = str(v or "").strip().upper()
            if s == "LIVE":
                return "Live"
            if s in ("NOT LIVE", "NOTLIVE", "NOT-LIVE"):
                return "Not Live"
            return None
        out, n = {}, 0
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
            ws = wb.active
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                if i == 0 or not row or len(row) < 5:
                    continue                                   # header / too-short
                loc = str(row[3] or "").strip().upper() if len(row) > 3 else ""
                if loc != "WEBSA":
                    continue                                   # only the WEBSA location
                code = str(row[1] or "").strip()               # B  product code
                if not code:
                    continue
                rec = {}
                st = norm_status(row[4])                        # E  catalog status
                if st:
                    rec["status"] = st
                osp = row[12] if len(row) > 12 else None        # M  OS Purchases
                if isinstance(osp, (int, float)):
                    rec["osPurchases"] = round(float(osp))
                if rec:
                    out[code] = rec
                    n += 1
        return out, n

    def parse_buying(self):
        """Parse an uploaded Buying Report and return per-SKU catalogue status + outstanding
        purchases (WEBSA rows only). Nothing changes yet — the client previews then applies."""
        try:
            length = int(self.headers.get("Content-Length", 0))
            buying, count = self.aggregate_buying(self.rfile.read(length))
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, 400)
            return
        self.send_json({"ok": True, "buying": buying, "fileSkus": count})

    def apply_buying(self, qs):
        """Write catalogue status + outstanding-purchase quantities into master.json for the
        given years, matched by code (WEBSA rows only, from the client)."""
        try:
            length = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(length)) if length else {}
            buying = body.get("buying") or {}
            yrs, _ = years_index()
            targets = [str(y) for y in (body.get("years") or []) if str(y) in yrs]
            applied = {}
            for y in targets:
                mpath = DATA / y / "master.json"
                master = read_json(mpath, None)
                if not master:
                    continue
                n = 0
                for s in master["skus"]:
                    rec = buying.get(s.get("code"))
                    if not rec:
                        continue
                    touched = False
                    if rec.get("status"):
                        s["status"] = rec["status"]
                        touched = True
                    if rec.get("osPurchases") is not None:
                        s["os_purchases"] = rec["osPurchases"]
                        touched = True
                    if touched:
                        n += 1
                mpath.write_text(json.dumps(master), encoding="utf-8")
                applied[y] = n
            self.send_json({"ok": True, "applied": applied, "years": targets})
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, 500)

    # ---- per-product duty rates (from the Tradeplan "Landed Costs" sheet, col "Duty Rate") ----
    @staticmethod
    def aggregate_duty(raw):
        """Read the Tradeplan 'Landed Costs' sheet → {code: duty_pct}. Finds the header row
        containing 'Duty Rate', plus the product-code column ('cmp_product' or 'Product').
        Stored fractions (0.06) are returned as percentages (6.0); values already >=1 are kept."""
        import io
        import warnings
        import openpyxl
        out, n = {}, 0
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            hdr = next((i for i, r in enumerate(rows)
                        if r and any(isinstance(c, str) and c.strip().lower() == "duty rate" for c in r)), None)
            if hdr is None:
                raise ValueError("Couldn't find a 'Duty Rate' column in the sheet.")
            header = rows[hdr]
            def find(*names):
                want = {x.lower() for x in names}
                return next((j for j, c in enumerate(header)
                             if isinstance(c, str) and c.strip().lower() in want), None)
            code_col, duty_col = find("cmp_product", "product"), find("duty rate")
            if code_col is None or duty_col is None:
                raise ValueError("Expected 'cmp_product' (or 'Product') and 'Duty Rate' columns.")
            for r in rows[hdr + 1:]:
                if not r or code_col >= len(r) or not r[code_col]:
                    continue
                code = str(r[code_col]).strip()
                d = r[duty_col] if duty_col < len(r) else None
                if code and isinstance(d, (int, float)):
                    out[code] = round(d * 100, 2) if abs(d) < 1 else round(d, 2)   # fraction → %
                    n += 1
        return out, n

    def parse_duty(self):
        """Parse an uploaded Landed Costs sheet and return per-SKU duty % (nothing changes yet)."""
        try:
            length = int(self.headers.get("Content-Length", 0))
            duty, count = self.aggregate_duty(self.rfile.read(length))
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, 400)
            return
        self.send_json({"ok": True, "duty": duty, "fileSkus": count})

    def apply_duty(self, qs):
        """Write per-product duty % into master.json for the given years, matched by code."""
        try:
            length = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(length)) if length else {}
            duty = body.get("duty") or {}
            yrs, _ = years_index()
            targets = [str(y) for y in (body.get("years") or []) if str(y) in yrs]
            applied = {}
            for y in targets:
                mpath = DATA / y / "master.json"
                master = read_json(mpath, None)
                if not master:
                    continue
                n = 0
                for s in master["skus"]:
                    r = duty.get(s.get("code"))
                    if r is not None:
                        s["duty_rate"] = r
                        n += 1
                mpath.write_text(json.dumps(master), encoding="utf-8")
                applied[y] = n
            self.send_json({"ok": True, "applied": applied, "years": targets})
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, 500)

    # ---- weekly actual sales (the "WKnn Sales" export: Product SKU + Sales TY £) ----
    @staticmethod
    def aggregate_wksales(raw):
        """Read a weekly sales export → {code: sales £ this week}. Finds the header row
        containing 'Product SKU' and 'Sales TY'; duplicate codes are summed."""
        import io
        import warnings
        import openpyxl
        out, n = {}, 0
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
            ws = wb.active
            code_c = sales_c = hdr = None
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                if hdr is None:
                    lower = [str(c or "").strip().lower() for c in row]
                    if "product sku" in lower and "sales ty" in lower:
                        hdr, code_c, sales_c = i, lower.index("product sku"), lower.index("sales ty")
                    continue
                code = str(row[code_c] or "").strip() if len(row) > code_c else ""
                val = row[sales_c] if len(row) > sales_c else None
                if not code or not isinstance(val, (int, float)):
                    continue
                out[code] = out.get(code, 0.0) + float(val)
                n += 1
            if hdr is None:
                raise ValueError("Couldn't find a header row with 'Product SKU' and 'Sales TY'.")
        return out, n

    def parse_wksales(self):
        """Parse an uploaded weekly sales export and return per-SKU sales £. Nothing
        changes yet — the client previews (converting £ → units by ASP) then applies."""
        try:
            length = int(self.headers.get("Content-Length", 0))
            sales, count = self.aggregate_wksales(self.rfile.read(length))
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, 400)
            return
        self.send_json({"ok": True, "sales": sales, "fileRows": count})

    def apply_wksales(self, qs):
        """Write one week's actual sales units into the year's master.json (matched by
        code; units already converted client-side via each SKU's ASP) and advance the
        actuals/forecast boundary: data_week = max(data_week, week+1)."""
        try:
            ydir, year = year_dir(qs)
            length = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(length)) if length else {}
            week = int(body.get("week") or 0)
            units = body.get("units") or {}
            if not (1 <= week <= 53):
                self.send_json({"ok": False, "error": f"week {week} out of range"}, 400)
                return
            mpath = ydir / "master.json"
            master = read_json(mpath, None)
            if not master:
                self.send_json({"ok": False, "error": f"no master.json for {year}"}, 500)
                return
            n = 0
            for s in master["skus"]:
                u = units.get(s.get("code"))
                if u is None:
                    continue
                act = s.get("actual") or [0] * 53
                while len(act) < 53:
                    act.append(0)
                act[week - 1] = u
                s["actual"] = act
                n += 1
            new_dw = max(int(master.get("data_week") or 1), week + 1)
            master["data_week"] = new_dw
            mpath.write_text(json.dumps(master), encoding="utf-8")
            self.send_json({"ok": True, "applied": n, "week": week, "dataWeek": new_dw, "year": year})
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, 500)

    # ---- purchase-order / container parsing (PO ↔ container-date linking) ----
    @staticmethod
    def parse_po_websa(raw):
        """Parse a 'WEBSA Open PO Export'. One row per PO-line (PO × product). Returns
        { pos: {PO#: {supplier, lines:[{code, ordered, delivered, outstanding, due}]}},
          codes: [distinct product codes], rows }. Quantities outstanding = ordered-delivered."""
        import io
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
        ws = wb.active
        rows = ws.iter_rows(values_only=True)
        header = [("" if c is None else str(c)).strip().lower() for c in next(rows)]

        def col(*names):
            for want in names:
                for i, h in enumerate(header):
                    if h == want.lower():
                        return i
            # loose contains-match fallback
            for want in names:
                for i, h in enumerate(header):
                    if want.lower() in h:
                        return i
            return None

        ci_po = col("poit_costcode: websa", "po", "po number")
        ci_sup = col("plca applymap name", "supplier name", "supplier")
        ci_code = col("product code", "code")
        ci_ord = col("quantity ordered", "qty ordered", "ordered")
        ci_del = col("quantity delivered", "qty delivered", "delivered")
        ci_due = col("due date", "estimated uk arrival", "arrival date")
        if ci_po is None or ci_code is None:
            raise ValueError("Expected 'poit_costcode: WEBSA' and 'Product Code' columns.")
        pos, codes = {}, set()
        n = 0
        for row in rows:
            po = row[ci_po]
            code = row[ci_code]
            if not po or not code:
                continue
            n += 1
            po = str(po).strip().upper()
            code = str(code).strip()
            codes.add(code)
            ordered = _num(row[ci_ord]) if ci_ord is not None else 0
            delivered = _num(row[ci_del]) if ci_del is not None else 0
            entry = pos.setdefault(po, {"supplier": (str(row[ci_sup]).strip() if ci_sup is not None and row[ci_sup] else ""), "lines": []})
            entry["lines"].append({
                "code": code,
                "ordered": ordered,
                "delivered": delivered,
                "outstanding": max(0, ordered - delivered),
                "due": _isodate(row[ci_due]) if ci_due is not None else None,
            })
        wb.close()
        return {"pos": pos, "codes": sorted(codes), "rows": n}

    @staticmethod
    def parse_qlik(raw):
        """Parse a 'Qlik Container Export'. One row per container line; the 'Order' cell may
        list several PO numbers (consolidated container) and is often messy. Returns
        { dates: {PO#: [{deliveryCB, etaPort, etd, status, container, shipment, supplier}]},
          unparsed: [order cells with no PO], truncated: [cells ending 'etc'], rows }."""
        import io
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
        ws = wb.active
        rows = ws.iter_rows(values_only=True)
        header = [("" if c is None else str(c)).strip().lower() for c in next(rows)]

        def col(*names):
            for want in names:
                for i, h in enumerate(header):
                    if h == want.lower():
                        return i
            for want in names:
                for i, h in enumerate(header):
                    if want.lower() in h:
                        return i
            return None

        ci_ord = col("order")
        ci_sup = col("[supplier name]", "supplier name")
        ci_status = col("status")
        ci_etd = col("estimated date of departure", "etd")
        ci_eta = col("eta (arrival at uk port)", "eta")
        ci_cb = col("delivery date to cb", "delivery date")
        ci_cont = col("container no.", "container no", "container")
        ci_ship = col("shipment no", "shipment")
        if ci_ord is None:
            raise ValueError("Expected an 'Order' column in the container export.")
        dates, unparsed, truncated = {}, [], []
        n = 0
        for row in rows:
            order = row[ci_ord]
            if order is None or str(order).strip() == "":
                continue
            n += 1
            s = str(order)
            pos, has_etc = po_tokens(s)
            if has_etc:
                truncated.append(s.strip())
            if not pos:
                unparsed.append(s.strip())
                continue
            rec = {
                "deliveryCB": _isodate(row[ci_cb]) if ci_cb is not None else None,
                "etaPort": _isodate(row[ci_eta]) if ci_eta is not None else None,
                "etd": _isodate(row[ci_etd]) if ci_etd is not None else None,
                "status": (str(row[ci_status]).strip() if ci_status is not None and row[ci_status] is not None else ""),
                "container": (str(row[ci_cont]).strip() if ci_cont is not None and row[ci_cont] else ""),
                "shipment": (str(row[ci_ship]).strip() if ci_ship is not None and row[ci_ship] else ""),
                "supplier": (str(row[ci_sup]).strip() if ci_sup is not None and row[ci_sup] else ""),
            }
            for po in pos:
                dates.setdefault(po, []).append(rec)
        wb.close()
        # de-dupe diagnostic lists, keep them bounded
        unparsed = sorted(set(unparsed))
        truncated = sorted(set(truncated))
        return {"dates": dates, "unparsed": unparsed[:300], "unparsedCount": len(unparsed),
                "truncated": truncated[:300], "truncatedCount": len(truncated), "rows": n}

    def parse_po(self):
        """Parse + persist the WEBSA Open PO export (global, shared across years)."""
        try:
            length = int(self.headers.get("Content-Length", 0))
            data = self.parse_po_websa(self.rfile.read(length))
            data["importedAt"] = datetime.now().isoformat(timespec="seconds")
            (DATA / "po_websa.json").write_text(json.dumps(data), encoding="utf-8")
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, 400)
            return
        self.send_json({"ok": True, "poCount": len(data["pos"]), "codes": len(data["codes"]), "rows": data["rows"]})

    def parse_containers(self):
        """Parse + persist the Qlik Container export (global, shared across years)."""
        try:
            length = int(self.headers.get("Content-Length", 0))
            data = self.parse_qlik(self.rfile.read(length))
            data["importedAt"] = datetime.now().isoformat(timespec="seconds")
            (DATA / "po_containers.json").write_text(json.dumps(data), encoding="utf-8")
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, 400)
            return
        self.send_json({"ok": True, "poCount": len(data["dates"]), "rows": data["rows"],
                        "unparsed": data["unparsedCount"], "truncated": data["truncatedCount"]})


def main():
    server = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    url = f"http://localhost:{PORT}"
    print(f"2026 Tradeplan running at {url}  (Ctrl+C to stop)")
    threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")


if __name__ == "__main__":
    main()
