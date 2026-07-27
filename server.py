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


# ---- change log + revert -------------------------------------------------
# A timestamped history of file uploads and data edits. Each entry may carry a
# "before" snapshot of the data files it touched, enabling a one-click revert.
# The full history (metadata) is kept in changelog.json; to bound disk use we
# keep the actual revert snapshots only for the most recent few changes.
CHANGELOG_PATH = DATA / "changelog.json"
SNAP_DIR = DATA / "change_snapshots"
CHANGELOG_MAX = 300            # metadata entries kept in the full log
CHANGELOG_KEEP_SNAPSHOTS = 5   # most-recent revertable changes that keep a snapshot
COALESCE_WINDOW_S = 600        # merge same-label edits within 10 minutes into one entry


def _read_changelog():
    obj = read_json(CHANGELOG_PATH, None)
    return obj if isinstance(obj, dict) and isinstance(obj.get("entries"), list) else {"entries": []}


def _write_changelog(obj):
    CHANGELOG_PATH.write_text(json.dumps(obj), encoding="utf-8")


def _delete_snapshot(cid):
    if not cid:
        return
    try:
        (SNAP_DIR / (cid + ".json")).unlink()
    except OSError:
        pass


def _prune_changelog(obj):
    """Keep snapshots for only the most-recent CHANGELOG_KEEP_SNAPSHOTS revertable
    entries (older ones stay in the log but lose their revert ability), and cap the
    total number of metadata entries."""
    kept = 0
    for e in obj["entries"]:                       # newest first
        if e.get("revertable") and e.get("snap"):
            kept += 1
            if kept > CHANGELOG_KEEP_SNAPSHOTS:
                _delete_snapshot(e.get("id"))
                e["revertable"] = False
                e["snap"] = False
    if len(obj["entries"]) > CHANGELOG_MAX:
        for e in obj["entries"][CHANGELOG_MAX:]:
            if e.get("snap"):
                _delete_snapshot(e.get("id"))
        obj["entries"] = obj["entries"][:CHANGELOG_MAX]


def record_change(kind, label, detail, rel_files, coalesce_key=None):
    """Log a change (kind 'upload'|'edit'), snapshotting the CURRENT contents of
    rel_files (paths relative to DATA) BEFORE the caller mutates them. Consecutive
    same-key changes within COALESCE_WINDOW_S merge into the first entry (keeping its
    original snapshot) so a burst of manual edits is one revert point, not dozens."""
    now = datetime.now()
    obj = _read_changelog()
    entries = obj["entries"]
    if coalesce_key and entries:
        top = entries[0]
        try:
            age = (now - datetime.fromisoformat(top.get("ts"))).total_seconds()
        except (ValueError, TypeError):
            age = 1e9
        if (top.get("coalesceKey") == coalesce_key and top.get("revertable")
                and not top.get("reverted") and age <= COALESCE_WINDOW_S):
            top["ts"] = now.isoformat(timespec="seconds")
            top["detail"] = detail
            top["count"] = int(top.get("count", 1)) + 1
            _write_changelog(obj)
            return
    cid = now.strftime("%Y%m%d%H%M%S%f")
    snap = {}
    for rel in rel_files:
        p = DATA / rel
        try:
            snap[rel] = p.read_text(encoding="utf-8") if p.exists() else None
        except OSError:
            snap[rel] = None
    SNAP_DIR.mkdir(exist_ok=True)
    (SNAP_DIR / (cid + ".json")).write_text(json.dumps({"files": snap}), encoding="utf-8")
    entries.insert(0, {
        "id": cid, "ts": now.isoformat(timespec="seconds"), "kind": kind,
        "label": label, "detail": detail, "files": list(rel_files),
        "coalesceKey": coalesce_key, "revertable": True, "reverted": False,
        "snap": True, "count": 1,
    })
    _prune_changelog(obj)
    _write_changelog(obj)


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
                "channelIndex": with_imported_at(DATA / "channel_index.json"),
            })
        elif parsed.path == "/api/imported-orders":
            ydir, _ = year_dir(qs)
            self.send_json(read_json(ydir / "orders_imported.json", {}))
        elif parsed.path == "/api/export-supplier":
            self.export_supplier(parse_qs(parsed.query))
        elif parsed.path == "/api/configs":
            self.list_configs()
        elif parsed.path == "/api/changelog":
            self.send_json({"ok": True, "entries": _read_changelog()["entries"]})
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
                # log data edits (only real saves that carry the order layer — the
                # silent proposed/stockbase housekeeping saves are not logged)
                if "orders" in body:
                    yr = ydir.name
                    lbl = str(body.get("changeLabel") or "Manual edits")[:80]
                    snapf = [f"{yr}/orders.json"] + ([f"{yr}/proposed.json"] if "proposed" in body else [])
                    record_change("edit", lbl, f"{yr} plan", snapf, coalesce_key=f"{lbl}|{yr}")
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
        elif parsed.path == "/api/apply-cbm":
            self.apply_cbm(parse_qs(parsed.query))
        elif parsed.path == "/api/add-product":
            self.add_product(parse_qs(parsed.query))
        elif parsed.path == "/api/parse-buying":
            self.parse_buying()
        elif parsed.path == "/api/apply-buying":
            self.apply_buying(parse_qs(parsed.query))
        elif parsed.path == "/api/parse-channelindex":
            self.parse_channelindex()
        elif parsed.path == "/api/parse-wksales":
            self.parse_wksales()
        elif parsed.path == "/api/apply-wksales":
            self.apply_wksales(parse_qs(parsed.query))
        elif parsed.path == "/api/parse-duty":
            self.parse_duty()
        elif parsed.path == "/api/apply-duty":
            self.apply_duty(parse_qs(parsed.query))
        elif parsed.path == "/api/revert-change":
            self.revert_change(parse_qs(parsed.query))
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
            if targets:
                record_change("upload", "Selling prices (ASP)", f"{', '.join(targets)} · {src}",
                              [f"{y}/master.json" for y in targets])
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
            if targets:
                record_change("upload", "FOB & landed costs", ", ".join(targets),
                              [f"{y}/master.json" for y in targets])
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

    def apply_cbm(self, qs):
        """Write hand-edited CBM (m³/unit) into master.json for the given years, matched
        by code. Each edited product is tagged cbm_src='manual' (vs the workbook-import
        default) and keeps cbm_prev, so the Data editor can show import-vs-manual."""
        try:
            length = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(length)) if length else {}
            cbm = body.get("cbm") or {}
            yrs, _ = years_index()
            targets = [str(y) for y in (body.get("years") or []) if str(y) in yrs]
            if targets:
                record_change("upload", "Product CBM", f"{', '.join(targets)} · manual",
                              [f"{y}/master.json" for y in targets])
            applied = {}
            for y in targets:
                mpath = DATA / y / "master.json"
                master = read_json(mpath, None)
                if not master:
                    continue
                n = 0
                for s in master["skus"]:
                    v = cbm.get(s.get("code"))
                    if v is not None and v > 0:
                        s["cbm_prev"] = s.get("cbm")
                        s["cbm"] = v
                        s["cbm_src"] = "manual"
                        n += 1
                mpath.write_text(json.dumps(master), encoding="utf-8")
                applied[y] = n
            self.send_json({"ok": True, "applied": applied, "years": targets})
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, 500)

    def add_product(self, qs):
        """Create a brand-new product (and its supplier, if new) in master.json for the
        given years. No sales history: ly/actual seed to zero, stock history to the
        entered current stock, and base_forecast to the annual figure spread evenly.
        Cost/CBM/ASP are tagged 'manual'. Logged as a revertable changelog entry."""
        try:
            length = int(self.headers.get("Content-Length", 0))
            b = json.loads(self.rfile.read(length)) if length else {}
            code = str(b.get("code") or "").strip()
            name = str(b.get("name") or "").strip()
            supplier = str(b.get("supplier") or "").strip()
            if not code or not name or not supplier:
                self.send_json({"ok": False, "error": "Product code, name and supplier are all required."}, 400)
                return
            yrs, _ = years_index()
            targets = [str(y) for y in (b.get("years") or []) if str(y) in yrs]
            if not targets:
                self.send_json({"ok": False, "error": "No valid target years."}, 400)
                return
            masters, clash = {}, []
            for y in targets:
                m = read_json(DATA / y / "master.json", None)
                if not m:
                    continue
                masters[y] = m
                if any(str(s.get("code", "")).strip() == code for s in m.get("skus", [])):
                    clash.append(y)
            if clash:
                self.send_json({"ok": False, "error": f"A product with code '{code}' already exists in {', '.join(clash)}."}, 400)
                return

            def num(v, d=None):
                try:
                    return float(v)
                except (TypeError, ValueError):
                    return d

            WEEKS = 53
            annual = num(b.get("annualForecast"), 0) or 0
            per = round(annual / WEEKS, 4) if annual > 0 else 0.0
            stock_now = num(b.get("stock_now"), 0) or 0
            common = {
                "code": code, "name": name, "supplier": supplier,
                "season": str(b.get("season") or "No Defined Season"),
                "category": (str(b.get("category")).strip() or None) if b.get("category") else None,
                "status": str(b.get("status") or "Live"),
                "fob": num(b.get("fob")), "fob_src": "manual",
                "landed": num(b.get("landed")), "landed_src": "manual",
                "asp": num(b.get("asp")), "asp_src": "manual",
                "cbm": num(b.get("cbm")), "cbm_src": "manual",
                "stock_now": stock_now,
                "fpq": num(b.get("fpq")),
                "pallet_type": (str(b.get("palletType")).strip() or None) if b.get("palletType") else None,
                "image": (str(b.get("image")).strip() or None) if b.get("image") else None,
                "duty_rate": None,
            }
            ns = b.get("newSupplier") or {}
            supplier_rec = {"name": supplier, "number": ns.get("number"),
                            "contact": ns.get("contact") or None, "port": ns.get("port") or None,
                            "email": ns.get("email") or None, "origin": ns.get("origin") or None}
            record_change("edit", f"Add product {code}", f"{name} · {', '.join(targets)}",
                          [f"{y}/master.json" for y in targets])
            applied = {}
            for y in targets:
                m = masters.get(y)
                if not m:
                    continue
                ids = {s.get("id") for s in m.get("skus", [])}
                sid, n = code, 2
                while sid in ids:
                    sid = f"{code}#{n}"; n += 1
                sku = dict(common)
                sku["id"] = sid
                sku["ly"] = [0.0] * WEEKS
                sku["actual"] = [0.0] * WEEKS
                sku["base_forecast"] = [per] * WEEKS
                sku["running_stock"] = [float(stock_now)] * WEEKS
                m.setdefault("skus", []).append(sku)
                if not any(str(s.get("name", "")).strip() == supplier for s in m.get("suppliers", [])):
                    m.setdefault("suppliers", []).append(dict(supplier_rec))
                (DATA / y / "master.json").write_text(json.dumps(m), encoding="utf-8")
                applied[y] = sid
            self.send_json({"ok": True, "code": code, "supplier": supplier, "years": targets, "applied": applied})
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, 500)

    # ---- catalogue status + outstanding purchases (from the "Buying Report") ----
    @staticmethod
    def aggregate_buying(raw):
        """Read a Buying Report → {code: {status, osPurchases, stock}} for WEBSA rows only.
        Columns: B=product code, D=location (keep only 'WEBSA'), E=catalog status
        (LIVE/NOT LIVE → Live/Not Live), H=Stock (live warehouse stock), M=OS Purchases."""
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
                stk = row[7] if len(row) > 7 else None          # H  live warehouse stock
                if isinstance(stk, (int, float)):
                    rec["stock"] = float(stk)
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
            if targets:
                record_change("upload", "Buying report", "catalogue status, stock & OS purchases · " + ", ".join(targets),
                              [f"{y}/master.json" for y in targets])
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
                    if rec.get("stock") is not None:
                        s["stock_now"] = rec["stock"]
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
            if targets:
                record_change("upload", "Duty rates", ", ".join(targets),
                              [f"{y}/master.json" for y in targets])
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

    def live_year_data_week(self):
        """`data_week` of the year that contains today (matched by each master's
        week1_start), or None if no held year covers today. Used to gate the
        channel-index weight basis: data_week < 2 means this year has no completed
        week yet (a fresh January), so we lean on last year's channel mix."""
        today = datetime.now().date()
        for yd in DATA.iterdir():
            if not yd.is_dir():
                continue
            m = read_json(yd / "master.json", None)
            if not m:
                continue
            try:
                y, mo, d = (int(x) for x in str(m.get("week1_start") or "").split("-"))
                off = (today - datetime(y, mo, d).date()).days
            except (ValueError, TypeError):
                continue
            if 0 <= off < 53 * 7:
                return int(m.get("data_week") or 1)
        return None

    # ---- channel index (per-SKU per-customer unit share + selling price) ----
    @staticmethod
    def aggregate_channelindex(raw, ty_ready=True):
        """Build the per-SKU per-customer channel index. Returns
        ( { skus: {code: [{c, r, p}]}, customers: {code: name} }, kept_rows, basis )
        — the client renormalises `r` to shares summing to 1, so `r` is a raw weight.

        Weight basis (NEW 'Channel Sales' format): normally each customer's
        **this-year units** (Sales Qty TY) — the current channel mix. When
        `ty_ready` is False (start of a new year, before a full week of TY data),
        weight by **last-year units** (Sales Qty LY) instead. Per SKU the basis
        falls back to the other year when the preferred one has no units for that
        product, so a line that only sold in one of the two years is still split
        rather than dropped. Price = this-year Average Selling Price, else LY Av
        Selling Price, else null (→ SKU ASP). Columns matched by header name.

        OLD forecasting-workbook format (a 'CustomerIndex' sheet + 'Add-ons pc'
        names, pre-computed ratios) stays supported as a fallback; basis 'ratio'."""
        import io
        import warnings
        import openpyxl

        def num(v):
            return float(v) if isinstance(v, (int, float)) else None

        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)

            # --- OLD format: explicit CustomerIndex sheet with pre-computed ratios ---
            if "CustomerIndex" in wb.sheetnames:
                skus, rows = {}, 0
                it = wb["CustomerIndex"].iter_rows(values_only=True)
                next(it, None)                                 # header
                for row in it:
                    code = str(row[0] or "").strip()
                    cust = str(row[1] or "").strip()
                    ratio = row[2] if len(row) > 2 else None
                    if not code or not cust or not isinstance(ratio, (int, float)) or ratio <= 0:
                        continue
                    price = row[3] if len(row) > 3 else None
                    skus.setdefault(code, []).append({
                        "c": cust, "r": round(float(ratio), 6),
                        "p": round(float(price), 4) if isinstance(price, (int, float)) else None})
                    rows += 1
                customers = {}
                if "Add-ons pc" in wb.sheetnames:
                    it = wb["Add-ons pc"].iter_rows(values_only=True)
                    next(it, None)
                    for row in it:
                        no = str(row[0] or "").strip()
                        nm = str(row[1] or "").strip()
                        if no and nm:
                            customers[no] = nm
                return {"skus": skus, "customers": customers}, rows, "ratio"

            # --- flat exports, columns matched by header name. Two layouts supported:
            #   'flat' : Product SKU / Customer Code / Sales Qty TY / Sales Qty LY / ASP
            #            (both-years export — weight TY units, LY fallback)
            #   'soh'  : cmp_product / soh_cusndcode / Quantity / Average Channel Sale
            #            Price £ — a single-period, single-channel (online) export with
            #            no TY/LY split, so weight by Quantity and price = Sales / Qty.
            ws = wb.active
            it = ws.iter_rows(values_only=True)
            col = {}
            fmt = None
            for hdr in it:                                     # find + map the header row
                lower = [str(c or "").strip().lower() for c in hdr]
                if "product sku" in lower and "customer code" in lower:
                    fmt = "flat"
                    col = {name: i for i, name in enumerate(lower)}
                    break
                if "cmp_product" in lower and "soh_cusndcode" in lower:
                    fmt = "soh"
                    col = {name: i for i, name in enumerate(lower)}
                    break
            if not fmt:
                raise ValueError("Couldn't find a Channel Sales header row — expected either "
                                 "'Product SKU' + 'Customer Code', or 'cmp_product' + 'soh_cusndcode'. "
                                 "Is this the Channel Sales export?")

            def cell(row, name):
                i = col.get(name)
                return row[i] if (i is not None and i < len(row)) else None

            # --- 'soh' layout: one period, one channel. Weight = Quantity (summed per
            # SKU+customer, dropping returns/credits); price = Sales / Qty. ---
            if fmt == "soh":
                agg, customers = {}, {}
                for row in it:
                    code = str(cell(row, "cmp_product") or "").strip()
                    cust = str(cell(row, "soh_cusndcode") or "").strip()
                    if not code or not cust:
                        continue
                    nm = str(cell(row, "[soh source acc name]") or "").strip()
                    if nm:
                        customers[cust] = nm
                    qty = num(cell(row, "quantity")) or 0.0
                    if qty <= 0:                               # skip returns/credits + zero-qty rows
                        continue
                    sales = num(cell(row, "sales")) or 0.0
                    e = agg.setdefault((code, cust), {"qty": 0.0, "sales": 0.0})
                    e["qty"] += qty
                    e["sales"] += sales
                skus, rows = {}, 0
                for (code, cust), e in agg.items():
                    if e["qty"] <= 0:
                        continue
                    price = e["sales"] / e["qty"] if e["qty"] > 0 else None
                    skus.setdefault(code, []).append({
                        "c": cust, "r": round(e["qty"], 4),
                        "p": round(price, 4) if (price and price > 0) else None})
                    rows += 1
                return {"skus": skus, "customers": customers}, rows, "online"

            # gather every customer row per SKU with both years' units + a price
            raw_by_code, customers = {}, {}
            for row in it:
                code = str(cell(row, "product sku") or "").strip()
                cust = str(cell(row, "customer code") or "").strip()
                if not code or not cust:
                    continue
                name = str(cell(row, "customer name") or "").strip()
                if name:
                    customers[cust] = name
                ty = num(cell(row, "sales qty ty")) or 0.0
                ly = num(cell(row, "sales qty ly")) or 0.0
                if ty <= 0 and ly <= 0:                        # never bought this SKU → skip
                    continue
                price = num(cell(row, "average selling price"))
                if not (price and price > 0):
                    price = num(cell(row, "ly av selling price"))
                raw_by_code.setdefault(code, []).append({"c": cust, "ty": ty, "ly": ly, "p": price})

            primary, fallback = ("ty", "ly") if ty_ready else ("ly", "ty")
            skus, rows = {}, 0
            for code, lst in raw_by_code.items():
                # per SKU use the preferred year's units; fall back to the other year
                # only when the preferred year has no units for this product at all
                use = primary if sum(r[primary] for r in lst) > 0 else fallback
                for r in lst:
                    w = r[use]
                    if w <= 0:
                        continue
                    skus.setdefault(code, []).append({
                        "c": r["c"], "r": round(w, 4),
                        "p": round(r["p"], 4) if (r["p"] and r["p"] > 0) else None})
                    rows += 1
        return {"skus": skus, "customers": customers}, rows, primary

    def parse_channelindex(self):
        """Parse + persist the channel index (GLOBAL, like the PO uploads). The weight
        basis is this-year units, unless the live year has no completed week yet
        (data_week < 2) — a fresh January — in which case last-year units are used."""
        try:
            dw = self.live_year_data_week()
            ty_ready = (dw is None) or (dw >= 2)   # ≥1 full week banked → trust this year
            length = int(self.headers.get("Content-Length", 0))
            data, rows, basis = self.aggregate_channelindex(self.rfile.read(length), ty_ready=ty_ready)
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, 400)
            return
        data["importedAt"] = datetime.now().isoformat(timespec="seconds")
        data["basis"] = basis
        record_change("upload", "Channel index", f"{len(data['skus'])} SKUs · {rows} shares · basis {basis}", ["channel_index.json"])
        (DATA / "channel_index.json").write_text(json.dumps(data), encoding="utf-8")
        self.send_json({"ok": True, "skus": len(data["skus"]), "rows": rows,
                        "customers": len(data["customers"]), "basis": basis})

    # ---- weekly actual sales (the "WKnn Sales" export: Product SKU / Sales TY £ / Qty TY units) ----
    @staticmethod
    def aggregate_wksales(raw):
        """Read a weekly sales export → {code: {"val": £, "qty": units}}. Finds the header
        row containing 'Product SKU', 'Sales TY' and 'Qty TY'; duplicate codes are summed."""
        import io
        import warnings
        import openpyxl
        out, n = {}, 0
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
            ws = wb.active
            code_c = sales_c = qty_c = hdr = None
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                if hdr is None:
                    lower = [str(c or "").strip().lower() for c in row]
                    if "product sku" in lower and "sales ty" in lower and "qty ty" in lower:
                        hdr = i
                        code_c, sales_c, qty_c = lower.index("product sku"), lower.index("sales ty"), lower.index("qty ty")
                    continue
                code = str(row[code_c] or "").strip() if len(row) > code_c else ""
                val = row[sales_c] if len(row) > sales_c else None
                qty = row[qty_c] if len(row) > qty_c else None
                if not code or not isinstance(qty, (int, float)):
                    continue
                rec = out.setdefault(code, {"val": 0.0, "qty": 0.0})
                rec["qty"] += float(qty)
                if isinstance(val, (int, float)):
                    rec["val"] += float(val)
                n += 1
            if hdr is None:
                raise ValueError("Couldn't find a header row with 'Product SKU', 'Sales TY' and 'Qty TY'.")
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
        """Write one week's actual sales units (file's 'Qty TY') into the year's
        master.json, stamp each seller's weekly ASP (asp_wk = Sales TY / Qty TY),
        and advance the actuals/forecast boundary: data_week = max(data_week, week+1).
        Closing stock is CHAINED for every SKU (sold or not) from the uploaded week
        through the last actualised week: running_stock[w] = previous week's closing
        + committed arrivals − actual units sold. Chaining from the prior week's
        closing (not the live stock_now snapshot) keeps the figure consistent with
        the grid's own history and safe to re-derive — re-applying a past week
        simply re-chains it and every later actualised week."""
        try:
            ydir, year = year_dir(qs)
            length = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(length)) if length else {}
            week = int(body.get("week") or 0)
            units = body.get("units") or {}
            wkasp = body.get("wkasp") or {}
            if not (1 <= week <= 53):
                self.send_json({"ok": False, "error": f"week {week} out of range"}, 400)
                return
            mpath = ydir / "master.json"
            master = read_json(mpath, None)
            if not master:
                self.send_json({"ok": False, "error": f"no master.json for {year}"}, 500)
                return
            orders = read_json(ydir / "orders.json", {}) or {}
            old_dw = int(master.get("data_week") or 1)
            closed_through = max(week, old_dw - 1)   # re-chain any later actualised weeks too
            n = 0
            for s in master["skus"]:
                code = s.get("code")
                u = units.get(code)
                act = s.get("actual") or [0] * 53
                while len(act) < 53:
                    act.append(0)
                if u is not None:
                    act[week - 1] = u
                    n += 1
                s["actual"] = act
                if wkasp.get(code) is not None:
                    s["asp_wk"] = wkasp[code]
                    s["asp_wk_week"] = week
                rs = s.get("running_stock") or [0] * 53
                while len(rs) < 53:
                    rs.append(0)
                ordv = orders.get(s.get("id")) or [0] * 53
                for w in range(week, closed_through + 1):
                    prev = rs[w - 2] if w >= 2 else float(s.get("stock_now") or 0)
                    rs[w - 1] = max(0.0, float(prev or 0) + float(ordv[w - 1] or 0) - float(act[w - 1] or 0))
                s["running_stock"] = rs
            new_dw = max(old_dw, week + 1)
            master["data_week"] = new_dw
            record_change("upload", f"Weekly sales · week {week}", f"{n} products · {year}", [f"{year}/master.json"])
            mpath.write_text(json.dumps(master), encoding="utf-8")
            self.send_json({"ok": True, "applied": n, "week": week, "dataWeek": new_dw,
                            "year": year, "closedThrough": closed_through})
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

    def revert_change(self, qs):
        """Restore the data files a logged change touched, from its before-snapshot,
        then log the revert itself (not revertable). Returns the affected years so the
        client can reload."""
        try:
            cid = (qs.get("id") or [""])[0]
            obj = _read_changelog()
            entry = next((e for e in obj["entries"] if e.get("id") == cid), None)
            if not entry:
                self.send_json({"ok": False, "error": "change not found"}, 404)
                return
            if not entry.get("revertable") or entry.get("reverted"):
                self.send_json({"ok": False, "error": "this change can no longer be reverted"}, 400)
                return
            snap = read_json(SNAP_DIR / (cid + ".json"), None)
            if not snap or "files" not in snap:
                self.send_json({"ok": False, "error": "snapshot no longer available"}, 400)
                return
            restored = []
            for rel, content in snap["files"].items():
                p = DATA / rel
                if content is None:
                    if p.exists():
                        p.unlink()
                else:
                    p.parent.mkdir(parents=True, exist_ok=True)
                    p.write_text(content, encoding="utf-8")
                restored.append(rel)
            entry["reverted"] = True
            entry["revertable"] = False
            entry["snap"] = False
            _delete_snapshot(cid)
            now = datetime.now()
            when = (entry.get("ts") or "")[:16].replace("T", " ")
            obj["entries"].insert(0, {
                "id": now.strftime("%Y%m%d%H%M%S%f"), "ts": now.isoformat(timespec="seconds"),
                "kind": "revert", "label": "Reverted: " + (entry.get("label") or ""),
                "detail": "Undid change from " + when, "files": [], "coalesceKey": None,
                "revertable": False, "reverted": False, "snap": False, "count": 1,
            })
            _write_changelog(obj)
            years = sorted({rel.split("/")[0] for rel in restored if "/" in rel})
            self.send_json({"ok": True, "restored": restored, "years": years})
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, 500)

    def parse_po(self):
        """Parse + persist the WEBSA Open PO export (global, shared across years)."""
        try:
            length = int(self.headers.get("Content-Length", 0))
            data = self.parse_po_websa(self.rfile.read(length))
            data["importedAt"] = datetime.now().isoformat(timespec="seconds")
            record_change("upload", "WEBSA Open PO", f"{len(data['pos'])} POs · {data['rows']} lines", ["po_websa.json"])
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
            record_change("upload", "Qlik containers", f"{len(data['dates'])} POs · {data['rows']} rows", ["po_containers.json"])
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
