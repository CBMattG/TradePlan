"""Per-supplier Excel order-planning form: export and parse.

Mirrors the factory "Initial Order Planning / Scheduling" workbook style:
a fixed left reference block (SKU, FOB, CBM, totals) plus one 4-column shipment
group per week (New Order / Value / CBM / Remaining in factory), with the same
colour scheme, merged headers and frozen panes.

export_supplier(...)  -> writes the .xlsx for one supplier
parse_supplier_form(...) -> reads New Order quantities back out, keyed by SKU + week
"""
import io
import json
import re
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

DATA = Path(__file__).resolve().parent / "data"
WEEKS = 53
WEEK1_START = date(2025, 12, 29)
CONTAINER_CBM = 68

# Lead-time / payment assumptions shown in each form's delivery footer + summary.
# The app overrides these from Settings → Supplier forms (data/settings.json →
# "supplier_form"); these are the fallbacks used when a value isn't set. Percentages
# are stored as fractions of order value (0.01 = 1%).
FORM_DEFAULTS = {
    "sailing_days": 50,     # port-to-port sea days
    "grace_days": 7,        # production grace period
    "inland_days": 7,       # UK inland transport days
    "marketing_pct": 0.01,  # marketing contribution
    "deposit_pct": 0.15,    # deposit
}

# ---- palette (solid ARGB, approximating the sample's themed fills) ----
C_REF_HDR = "FFFFECAF"   # amber  – reference headers
C_INPUT_HDR = "FFFFC000"  # orange – the key input column (proposed order qty)
C_ARRIVAL = "FF00B0F0"   # blue   – estimated arrival week band
C_SHIP = "FF92D050"      # green  – shipment / week band
C_SUBHDR = "FFD9E1F2"    # light blue-grey – sub headers
C_STATUS = "FFFFFF00"    # yellow – status + 2026 FOB (editable)
C_INPUT = "FFFFF2CC"     # light yellow – New Order input cells
C_SKU = "FFF2F2F2"       # very light grey – SKU column
C_TOTAL = "FFFCE4D6"     # light orange – totals row

FONT = "Calibri"
thin = Side(style="thin", color="FFBFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
THICK = Side(style="thick", color="FF000000")   # heavy black – the A–K frozen block
MED = Side(style="medium", color="FF000000")     # black – separates each week group

FMT_USD = '[$$-409]#,##0.00'
FMT_PCT = '0.00%'
FMT_CBM = '0.000'
FMT_QTY = '0'
FMT_QTY_T = '#,##0'
DATE_FMT = 'mm-dd-yy'   # footer dates — matches the hand-built form (Excel's default date style)

REF_COLS = [  # (header, width, number_format, fill)
    ("SKU", 22, None, C_REF_HDR),
    ("Status", 9.5, None, C_REF_HDR),
    ("2025 FOB", 11, FMT_USD, C_REF_HDR),
    ("2026 FOB", 10, FMT_USD, C_REF_HDR),
    ("% Variance", 9, FMT_PCT, C_REF_HDR),
    ("CBM", 8, FMT_CBM, C_REF_HDR),
    ("2025 Volume", 11, FMT_QTY_T, C_REF_HDR),
    ("Total FOB Value\n(2025 FOB Cost)", 11, FMT_USD, C_REF_HDR),
    ("2026 Proposed Total Order QTY", 13, FMT_QTY_T, C_INPUT_HDR),
    ("Total FOB Value\n(2026 FOB Cost)", 12, FMT_USD, C_REF_HDR),
    ("Total CBM", 8.5, '#,##0.00', C_REF_HDR),
]
GROUP_SUBHDRS = ["New Order", "Value", "CBM", "Remaining QTY in FTY after shipment"]
GROUP_WIDTHS = [9.5, 13, 7.5, 17]
HDR_ROW1, HDR_SHIP, HDR_SUB, DATA0 = 1, 2, 3, 4   # compact: no blank spacer rows


def box(ws, r1, c1, r2, c2, side):
    """Draw a heavy outer border around a rectangle, keeping inner thin lines."""
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            b = cell.border
            cell.border = Border(
                left=side if c == c1 else b.left,
                right=side if c == c2 else b.right,
                top=side if r == r1 else b.top,
                bottom=side if r == r2 else b.bottom,
            )


def week_start(w):
    return WEEK1_START + timedelta(days=(w - 1) * 7)


def _fill(c):
    return PatternFill("solid", fgColor=c)


def _year_totals_for_codes(ydir, codes, container_cbm=CONTAINER_CBM):
    """Order Quantity / Value / Containers for a set of SKU codes in one year's
    data folder. Years are matched by SKU code (supplier names differ between
    years), each year using its own FOB / CBM. Uses each year's full order plan
    (committed orders + any proposed rebuys), so a forecast year whose orders are
    all proposals (e.g. a fresh 2027) still shows its planned volume."""
    ydir = Path(ydir)
    try:
        master = json.loads((ydir / "master.json").read_text(encoding="utf-8"))
        orders = json.loads((ydir / "orders.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return 0.0, 0.0, 0.0
    prop = {}
    pp = ydir / "proposed.json"
    if pp.exists():
        try:
            prop = json.loads(pp.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            prop = {}
    qty = val = cbm = 0.0
    for s in master["skus"]:
        if s["code"] not in codes:
            continue
        tot = sum(orders.get(s["id"]) or [0] * WEEKS) + sum(prop.get(s["id"]) or [])
        qty += tot
        val += tot * s.get("fob", 0)
        cbm += tot * s.get("cbm", 0)
    return qty, val, (cbm / container_cbm if container_cbm else 0.0)


def _ref_cols(year):
    """Reference-block headers, with the year labels matched to the export year:
    the previous year for the historical reference columns, the export year for
    the proposed-order columns."""
    cy = str(year) if year is not None else "This Year"
    py = str(int(year) - 1) if (year is not None and str(year).isdigit()) else "Last Year"
    return [
        ("SKU", 22, None, C_REF_HDR),
        ("Status", 9.5, None, C_REF_HDR),
        (f"{py} FOB", 11, FMT_USD, C_REF_HDR),
        (f"{cy} FOB", 10, FMT_USD, C_REF_HDR),
        ("% Variance", 9, FMT_PCT, C_REF_HDR),
        ("CBM", 8, FMT_CBM, C_REF_HDR),
        (f"{py} Volume", 11, FMT_QTY_T, C_REF_HDR),
        (f"Total FOB Value\n({py} FOB Cost)", 11, FMT_USD, C_REF_HDR),
        (f"{cy} Proposed Total Order QTY", 13, FMT_QTY_T, C_INPUT_HDR),
        (f"Total FOB Value\n({cy} FOB Cost)", 12, FMT_USD, C_REF_HDR),
        ("Total CBM", 8.5, '#,##0.00', C_REF_HDR),
    ]


def _prev_year_by_code(ddir, year, codes):
    """{code: {'fob':, 'units':}} for the previous year, matched by SKU code,
    using that year's committed order units (what was actually ordered, not sales).
    Empty if the previous year isn't held."""
    if year is None or not str(year).isdigit():
        return {}
    pdir = Path(ddir).parent / str(int(year) - 1)
    try:
        master = json.loads((pdir / "master.json").read_text(encoding="utf-8"))
        orders = json.loads((pdir / "orders.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    out = {}
    for s in master["skus"]:
        c = s["code"]
        if c not in codes:
            continue
        units = sum(orders.get(s["id"]) or [0] * WEEKS)
        if c in out:
            out[c]["units"] += units
        else:
            out[c] = {"fob": s.get("fob", 0), "units": units}
    return out


def export_supplier(supplier, master=None, orders=None, current_week=None, data_dir=None, year=None, form_cfg=None):
    """Build the workbook for one supplier; returns BytesIO of the .xlsx."""
    ddir = Path(data_dir) if data_dir else DATA
    if master is None:
        master = json.loads((ddir / "master.json").read_text(encoding="utf-8"))
    cfg = {**FORM_DEFAULTS, **(form_cfg or {})}
    # Week-1 anchor for THIS year's calendar: a 2027 form must date its weeks from
    # 2026-12-28, not the module default (2025-12-29 = 2026's week 1). Each year's
    # master carries its own week1_start; fall back to the module constant.
    try:
        _y, _m, _d = (int(x) for x in str(master.get("week1_start")).split("-"))
        wk1 = date(_y, _m, _d)
    except (ValueError, AttributeError, TypeError):
        wk1 = WEEK1_START

    def wstart(w):
        return wk1 + timedelta(days=(w - 1) * 7)
    if orders is None:
        orders = json.loads((ddir / "orders.json").read_text(encoding="utf-8"))
        # merge proposed rebuy suggestions so the exported form = committed + proposed
        prop_path = ddir / "proposed.json"
        if prop_path.exists():
            try:
                prop = json.loads(prop_path.read_text(encoding="utf-8"))
                for sid, arr in prop.items():
                    base = orders.get(sid) or [0] * WEEKS
                    orders[sid] = [(base[w] if w < len(base) else 0) + (arr[w] if w < len(arr) else 0) for w in range(WEEKS)]
            except (OSError, ValueError):
                pass
    if current_week is None:
        current_week = int(master.get("data_week", 1))

    skus = [s for s in master["skus"] if s["supplier"] == supplier]
    skus.sort(key=lambda s: s["code"])

    ref_cols = _ref_cols(year)                                   # year labels match the export year
    prevmap = _prev_year_by_code(ddir, year, {s["code"] for s in skus})  # prev-year FOB + order units by code

    def ordv(s):
        return orders.get(s["id"], [0] * WEEKS)

    # only show weeks that actually carry a planned order for this supplier
    has_order = [any(ordv(s)[w] for s in skus) for w in range(WEEKS)]
    weeks = sorted({w + 1 for w in range(WEEKS) if has_order[w]})

    wb = Workbook()
    ws = wb.active
    ws.title = "Scheduling"
    ws.sheet_view.showGridLines = False

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    centerw = Alignment(horizontal="center", vertical="center", wrap_text=False)
    left = Alignment(horizontal="left", vertical="center")

    def put(r, c, val, fmt=None, fill=None, bold=False, align=center, size=9, color=None):
        cell = ws.cell(row=r, column=c, value=val)
        cell.font = Font(name=FONT, size=size, bold=bold, color=color)
        cell.alignment = align
        cell.border = BORDER
        if fmt:
            cell.number_format = fmt
        if fill:
            cell.fill = _fill(fill)
        return cell

    # ---- reference headers (merged vertically rows 1..5) ----
    for i, (hdr, width, _fmt, fill) in enumerate(ref_cols, start=1):
        put(HDR_ROW1, i, hdr, fill=fill, bold=True)
        ws.merge_cells(start_row=HDR_ROW1, start_column=i, end_row=DATA0 - 1, end_column=i)
        ws.column_dimensions[get_column_letter(i)].width = width
        for rr in range(HDR_ROW1, DATA0):
            ws.cell(row=rr, column=i).border = BORDER

    # ---- shipment group headers ----
    base = len(ref_cols) + 1  # first shipment column (L = 12)
    col_week = {}             # New Order column -> week
    for gi, wk in enumerate(weeks):
        b = base + gi * 4
        cl = get_column_letter(b)
        # Blue arrival header is a FORMULA: it reads the week number out of the
        # green "Shipment – Week N" cell (row 3) and rebuilds the label + w/c date,
        # so editing row 3 auto-updates this header. (+4 after "Week" tolerates a
        # space or none.)
        wk_expr = f'VALUE(TRIM(MID({cl}{HDR_SHIP},FIND("Week",{cl}{HDR_SHIP})+4,20)))'
        date_expr = f'DATE({wk1.year},{wk1.month},{wk1.day})+({wk_expr}-1)*7'
        arrival = ('=IFERROR("Estimated Arrival"&CHAR(10)&"Week "&' + wk_expr +
                   '&"  (w/c "&TEXT(' + date_expr + ',"dd/mm/yy")&")","Estimated Arrival")')
        ws.merge_cells(start_row=HDR_ROW1, start_column=b, end_row=HDR_ROW1, end_column=b + 3)
        put(HDR_ROW1, b, arrival, fill=C_ARRIVAL, bold=True, size=11)
        ws.merge_cells(start_row=HDR_SHIP, start_column=b, end_row=HDR_SHIP, end_column=b + 3)
        put(HDR_SHIP, b, f"Shipment – Week {wk}", fill=C_SHIP, bold=True)
        for j, sub in enumerate(GROUP_SUBHDRS):
            put(HDR_SUB, b + j, sub, fill=C_SUBHDR, bold=False)
            ws.column_dimensions[get_column_letter(b + j)].width = GROUP_WIDTHS[j]
        col_week[b] = wk

    # ---- data rows ----
    r = DATA0
    for s in skus:
        ov = ordv(s)
        total_order = sum(ov)
        prev = prevmap.get(s["code"], {})
        prev_units = int(round(prev.get("units", 0)))            # previous year's ORDER units (not sales)
        prev_fob = round(prev.get("fob", s["fob"]), 2)           # previous year's FOB cost (fallback to current)
        status = {"Live": "LIVE", "Not Live": "NOT LIVE"}.get(s.get("status"), "")
        put(r, 1, s["code"], align=left, fill=C_SKU)
        put(r, 2, status, fill=C_STATUS if status else None)
        put(r, 3, prev_fob, fmt=FMT_USD)                              # previous-year FOB
        put(r, 4, round(s["fob"], 2), fmt=FMT_USD, fill=C_STATUS)     # current-year FOB (editable)
        put(r, 5, f"=IFERROR((D{r}/C{r})-1,0)", fmt=FMT_PCT)
        put(r, 6, round(s["cbm"], 6), fmt=FMT_CBM)
        put(r, 7, prev_units, fmt=FMT_QTY_T)                          # previous-year order units
        put(r, 8, f"=G{r}*C{r}", fmt=FMT_USD)
        put(r, 9, total_order, fmt=FMT_QTY_T, fill=C_INPUT, bold=True)  # current-year proposed total
        put(r, 10, f"=I{r}*D{r}", fmt=FMT_USD)
        put(r, 11, f"=I{r}*F{r}", fmt='#,##0.00')
        prev_rem = None
        for gi, wk in enumerate(weeks):
            b = base + gi * 4
            no_col = get_column_letter(b)
            qty = ov[wk - 1]
            put(r, b, qty if qty else None, fmt=FMT_QTY, fill=C_INPUT)
            put(r, b + 1, f"={no_col}{r}*$D{r}", fmt=FMT_USD)   # value at CURRENT-year FOB (col D)
            put(r, b + 2, f"={no_col}{r}*$F{r}", fmt=FMT_CBM)
            if prev_rem is None:
                put(r, b + 3, f"=$I{r}-{no_col}{r}", fmt=FMT_QTY)
            else:
                put(r, b + 3, f"={prev_rem}{r}-{no_col}{r}", fmt=FMT_QTY)
            prev_rem = get_column_letter(b + 3)
        r += 1

    last = r - 1
    # ---- totals row ----
    tr = r
    put(tr, 1, "Totals", bold=True, align=left, fill=C_TOTAL)
    for c in range(2, len(ref_cols) + 1):
        put(tr, c, None, fill=C_TOTAL)
    for c in (7, 8, 9, 10, 11):
        L = get_column_letter(c)
        fmt = FMT_QTY_T if c in (7, 9) else ('#,##0.00' if c == 11 else FMT_USD)
        put(tr, c, f"=SUM({L}{DATA0}:{L}{last})", fmt=fmt, bold=True, fill=C_TOTAL)
    for gi in range(len(weeks)):
        b = base + gi * 4
        for off, fmt in ((0, FMT_QTY_T), (1, FMT_USD), (2, '#,##0.00')):
            L = get_column_letter(b + off)
            put(tr, b + off, f"=SUM({L}{DATA0}:{L}{last})", fmt=fmt, bold=True, fill=C_TOTAL)
        put(tr, b + 3, None, fill=C_TOTAL)

    # ---- geometry for everything below the totals row (offsets from `tr`) ----
    foot0 = tr + 1          # first delivery-footer row (under each week block)
    sr = tr + 2             # 'Estimated # of Containers' line (cols I–K)
    b1 = tr + 6             # 'Order plan by year' table header (sits below the footer)
    ov_row = b1 + 2         # its Order Value row — drives the payment summary
    lt_hdr = tr + 11        # lead-time assumptions header (holds the master cells)
    sail_row, grace_row, inland_row = lt_hdr + 1, lt_hdr + 2, lt_hdr + 3
    pay_hdr = tr + 16       # payment terms header
    mkt_row, dep_row, fin_row = pay_hdr + 1, pay_hdr + 2, pay_hdr + 3

    # ---- delivery footer under each week block ----------------------------
    # Explains the Goods Ready date the factory must hit for us to receipt at the
    # requested arrival week. Sailing days / grace / inland days are FORMULAS that
    # each read a single master cell (in the summary below), so editing one master
    # value re-flows every block's Goods Ready date. Goods Ready = arrival week
    # (this block's w/c date) − inland − grace − sailing.
    C_GOODS = "FFE2EFDA"   # light green — the highlighted 'Goods Ready date' row
    for gi, wk in enumerate(weeks):
        b = base + gi * 4
        vc = get_column_letter(b + 3)          # value column for this block
        w = wstart(wk)
        footer = [
            ("Estimated Sailing Days:",   f"=$B${sail_row}",   None),
            ("Production grace period:",  f"=$B${grace_row}",  None),
            ("UK Inland transport days:", f"=$B${inland_row}", None),
            ("UK Arrival week est:",      f"=DATE({w.year},{w.month},{w.day})", DATE_FMT),
            ("Goods Ready date:",
             f"={vc}{foot0 + 3}-{vc}{foot0 + 2}-{vc}{foot0 + 1}-{vc}{foot0}", DATE_FMT),
        ]
        for j, (label, value, fmt) in enumerate(footer):
            row = foot0 + j
            last = (j == len(footer) - 1)       # 'Goods Ready date' row: bold + light-green fill
            fill = C_GOODS if last else None
            # label merged across the first three block columns, centred (no wrap)
            put(row, b, label, align=centerw, size=11, bold=last, fill=fill)
            ws.merge_cells(start_row=row, start_column=b, end_row=row, end_column=b + 2)
            if fill:
                for cc in (b + 1, b + 2):
                    ws.cell(row=row, column=cc).fill = _fill(fill)
            put(row, b + 3, value, fmt=fmt, align=centerw, size=11, bold=last, fill=fill)
        # borders: a medium box around the footer with a thin label|value divider,
        # no interior horizontal lines (matches the hand-built form exactly).
        for j in range(len(footer)):
            r = foot0 + j
            top = MED if (j == 0 or j == len(footer) - 1) else None
            bot = MED if j == len(footer) - 1 else None
            for k in range(4):
                left = MED if k == 0 else None
                right = MED if k == 3 else (thin if k in (0, 2) else None)
                ws.cell(row=r, column=b + k).border = Border(left=left, right=right, top=top, bottom=bot)

    # ---- summary block ----
    kcol = get_column_letter(11)
    put(sr, 9, f'=CONCATENATE("Estimated # of Containers: ",ROUND({kcol}{tr}/{CONTAINER_CBM},2))',
        bold=True, align=left, fill=None)
    ws.merge_cells(start_row=sr, start_column=9, end_row=sr, end_column=11)

    # ---- cross-year order summary (Quantity / Value / Containers per year) ----
    # One column per plan year held in the app, matched to THIS supplier's SKU
    # codes across years. The exported year's column is a LIVE formula tied to
    # the form totals (so it updates as the factory edits the New Order cells);
    # the other years are static committed actuals computed at export time. The
    # table grows automatically as new years are added.
    root_data = ddir.parent
    all_years = sorted(p.name for p in root_data.iterdir() if p.is_dir() and p.name.isdigit())
    exp_year = str(year) if year else (ddir.name if ddir.name.isdigit() else None)
    scodes = {s["code"] for s in skus}
    static = {y: _year_totals_for_codes(root_data / y, scodes)
              for y in all_years if y != exp_year}

    ncol = {}
    put(b1, 1, "Order plan by year", align=left, bold=True, fill=C_REF_HDR)
    for i, y in enumerate(all_years):
        ncol[y] = 2 + i
        put(b1, ncol[y], int(y), bold=True, fill=C_REF_HDR)
    var_n = 2 + len(all_years)
    var_p = var_n + 1
    put(b1, var_n, "Variance #", bold=True, fill=C_REF_HDR)
    put(b1, var_p, "Variance %", bold=True, fill=C_REF_HDR)

    # the two most recent years drive the variance columns
    prev_y = all_years[-2] if len(all_years) >= 2 else None
    last_y = all_years[-1] if all_years else None

    rows = [
        ("Order Quantity",       FMT_QTY_T,   f"=I{tr}"),
        ("Order Value",          FMT_USD,     f"=J{tr}"),
        ("Number of Containers", "#,##0.00",  f"=K{tr}/{CONTAINER_CBM}"),
    ]
    for idx, (label, fmt, live_expr) in enumerate(rows):
        rr = b1 + 1 + idx
        put(rr, 1, label, align=left, bold=True)
        for y in all_years:
            c = ncol[y]
            if y == exp_year:
                put(rr, c, live_expr, fmt=fmt, bold=True, fill=C_INPUT)
            else:
                v = static[y][idx]
                put(rr, c, round(v) if idx == 0 else round(v, 2), fmt=fmt)
        if prev_y and last_y and prev_y != last_y:
            pc = get_column_letter(ncol[prev_y])
            lc = get_column_letter(ncol[last_y])
            put(rr, var_n, f"={lc}{rr}-{pc}{rr}", fmt=fmt, bold=True)
            # variance % = (new - old) / old  → negative when the latest year is lower
            put(rr, var_p, f"=IFERROR(({lc}{rr}-{pc}{rr})/{pc}{rr},0)", fmt="0%", bold=True)
        else:
            put(rr, var_n, None, fmt=fmt)
            put(rr, var_p, None)

    # the exported year's Order Value cell (falls back to the live form total)
    oval = f"{get_column_letter(ncol[exp_year])}{ov_row}" if exp_year in ncol else f"J{tr}"

    def hdr_merge(row, c1, c2, text, fill=C_REF_HDR):
        put(row, c1, text, align=left, bold=True, fill=fill)
        ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
        for cc in range(c1 + 1, c2 + 1):           # style (not value) the covered cells
            cell = ws.cell(row=row, column=cc)
            cell.fill = _fill(fill); cell.border = BORDER

    # ---- lead-time assumptions: the master cells the delivery footer references ----
    # Editing one of these B-column cells re-flows every block's Goods Ready date.
    hdr_merge(lt_hdr, 1, 3, "Lead-time assumptions (drive every block's Goods Ready date)")
    for lbl, mrow, key in (("Estimated Sailing Days:", sail_row, "sailing_days"),
                           ("Production grace period:", grace_row, "grace_days"),
                           ("UK Inland transport days:", inland_row, "inland_days")):
        put(mrow, 1, lbl, align=left)
        put(mrow, 2, int(cfg[key]), fmt=FMT_QTY, fill=C_INPUT, bold=True)
        put(mrow, 3, "days", align=left)

    # ---- payment terms & contributions (percentages of the order value) ----
    hdr_merge(pay_hdr, 1, 3, "Payment terms & contributions (of order value)")
    put(mkt_row, 1, "Marketing Contribution:", align=left)
    put(mkt_row, 2, float(cfg["marketing_pct"]), fmt=FMT_PCT, fill=C_INPUT, bold=True)
    put(mkt_row, 3, f"=B{mkt_row}*{oval}", fmt=FMT_USD)
    put(dep_row, 1, "Deposit Terms:", align=left)
    put(dep_row, 2, float(cfg["deposit_pct"]), fmt=FMT_PCT, fill=C_INPUT, bold=True)
    put(dep_row, 3, f"=B{dep_row}*{oval}", fmt=FMT_USD)
    put(fin_row, 1, "Final Balance:", align=left, bold=True)
    put(fin_row, 3, f"={oval}-C{mkt_row}-C{dep_row}", fmt=FMT_USD, bold=True)

    # ---- heavy borders: the A–K frozen block, and each week group ----
    box(ws, HDR_ROW1, 1, tr, len(ref_cols), THICK)
    for gi in range(len(weeks)):
        b = base + gi * 4
        box(ws, HDR_ROW1, b, tr, b + 3, MED)   # the week block (its footer draws its own box below)
        # the arrival/shipment headers are merged across the whole group; openpyxl
        # draws a merged range from its anchor cell, so force left+right medium there
        ws.cell(row=HDR_ROW1, column=b).border = Border(left=MED, right=MED, top=MED, bottom=thin)
        ws.cell(row=HDR_SHIP, column=b).border = Border(left=MED, right=MED, top=thin, bottom=thin)

    # ---- sheet chrome ----
    ws.row_dimensions[HDR_ROW1].height = 34
    ws.row_dimensions[HDR_SHIP].height = 15
    ws.row_dimensions[HDR_SUB].height = 44
    ws.freeze_panes = ws.cell(row=DATA0, column=base)
    ws.sheet_properties.pageSetUpPr.fitToPage = False

    # banner with supplier name + a note (row above everything via a title in A is awkward; use sheet title + header comment)
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


def export_forecast(rows, year=None, week1=None):
    """Sales-unit forecast workbook: SKU code in column A, catalogue status (Live /
    Not Live) in column B, then one column per week (W1..W53) of forecast sales
    units, with a Total at the far right. `rows` = [{code, status?, forecast:
    [..53..]}]; `week1` (ISO date) adds the week-commencing date under each header."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales Unit Forecast"
    hdr = _fill(C_REF_HDR)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center")
    wk1 = None
    try:
        y, m, d = (int(x) for x in str(week1).split("-"))
        wk1 = date(y, m, d)
    except (ValueError, AttributeError, TypeError):
        wk1 = None

    def put(r, c, val, bold=False, fill=None, align=center, fmt=None):
        cell = ws.cell(row=r, column=c, value=val)
        cell.font = Font(name=FONT, size=9, bold=bold)
        cell.alignment = align
        cell.border = BORDER
        if fill:
            cell.fill = fill
        if fmt:
            cell.number_format = fmt
        return cell

    W0 = 3   # first week column (C) — A=SKU, B=Status
    put(1, 1, "SKU", bold=True, fill=hdr, align=left)
    put(1, 2, "Status", bold=True, fill=hdr)
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 10
    for w in range(WEEKS):
        lbl = f"W{w + 1}"
        if wk1:
            lbl += "\n" + (wk1 + timedelta(days=7 * w)).strftime("%d/%m/%y")
        put(1, W0 + w, lbl, bold=True, fill=hdr)
        ws.column_dimensions[get_column_letter(W0 + w)].width = 8.5
    tcol = W0 + WEEKS
    put(1, tcol, "Total", bold=True, fill=hdr)
    ws.column_dimensions[get_column_letter(tcol)].width = 11

    r = 2
    for row in rows:
        put(r, 1, row.get("code", ""), align=left, fill=_fill(C_SKU))
        status = row.get("status") or ""
        put(r, 2, status, align=left, fill=_fill(C_STATUS) if status else None)
        fc = row.get("forecast") or []
        for w in range(WEEKS):
            v = fc[w] if w < len(fc) else 0
            put(r, W0 + w, int(round(v or 0)), fmt=FMT_QTY)
        L, R = get_column_letter(W0), get_column_letter(W0 + WEEKS - 1)
        put(r, tcol, f"=SUM({L}{r}:{R}{r})", bold=True, fmt=FMT_QTY_T)
        r += 1

    ws.freeze_panes = "C2"
    ws.row_dimensions[1].height = 30
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


def export_search(payload):
    """Flat workbook of the products behind the current sidebar search — one row per
    product with supplier, category and the figures a promo/marketing plan needs
    (stock on hand and its retail value, weeks cover, remaining forecast, YTD vs plan,
    price and margin). `payload` is the client's already-computed rows, so the file
    matches the screen exactly. Sheet is autofiltered with a totals row on top."""
    GBP = '£#,##0.00'
    cols = [
        ("Product", 18, "code", None), ("Description", 40, "name", None),
        ("Supplier", 26, "supplier", None), ("Category", 20, "category", None),
        ("Season", 14, "season", None), ("Status", 9, "status", None),
        ("NPD", 6, "npd", None), ("Tags", 34, "tags", None),
        ("Stock Now", 10, "stock", FMT_QTY_T), ("Stock Value (RRP)", 14, "stockValue", GBP),
        ("Stock at Cost", 12, "stockCost", GBP),
        ("Weeks Cover", 11, "cover", "#,##0.0"), ("Out of Stock Wk", 12, "outWeek", "#,##0"),
        ("Forecast Left", 12, "fcRest", FMT_QTY_T), ("Forecast Year", 12, "fcYear", FMT_QTY_T),
        ("Sold YTD", 12, "soldValue", GBP), ("YTD vs Plan %", 12, "ytdPct", "+0.0;-0.0;0.0"),
        ("Committed Left", 12, "committed", FMT_QTY_T), ("Proposed", 10, "proposed", FMT_QTY_T),
        ("Selling Price", 11, "asp", GBP), ("Landed Cost", 11, "landed", GBP),
        ("FOB", 10, "fob", GBP), ("Margin %", 9, "margin", "0.0"),
        ("CBM", 8, "cbm", FMT_CBM), ("Units/Pallet", 11, "fpq", "#,##0"),
        ("Pallet Type", 11, "palletType", None),
    ]
    rows = payload.get("rows") or []
    wb = Workbook()
    ws = wb.active
    ws.title = "Products"
    left = Alignment(horizontal="left", vertical="center")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    what = payload.get("label") or payload.get("term") or "search"
    kind = "tagged" if payload.get("isTag") else "matching"
    ws.cell(row=1, column=1, value=f"{payload.get('year', '')} products {kind} “{what}”").font = \
        Font(name=FONT, size=13, bold=True)
    bits = [f"{len(rows)} product{'' if len(rows) == 1 else 's'}",
            f"as at week {payload.get('week', '')}",
            f"generated {payload.get('generated', '')}"]
    if payload.get("statusFilter"):
        bits.insert(1, payload["statusFilter"])
    ws.cell(row=2, column=1, value=" · ".join(bits) + (
        " · stock figures are the current week's; forecast/committed are what remains from"
        " this week on · sorted by stock value, highest first")
    ).font = Font(name=FONT, size=9, italic=True, color="FF808080")

    # Totals sit ABOVE the header, not below the data — anything inside the autofilter
    # range would be treated as a data row and get sorted/hidden with it.
    tr, hr = 4, 5
    SUMMED = {"stock", "stockValue", "stockCost", "fcRest", "fcYear", "soldValue", "committed", "proposed"}
    for c, (_l, _w, k, fmt) in enumerate(cols, start=1):
        v = "TOTAL" if k == "code" else (sum(r.get(k) or 0 for r in rows) if k in SUMMED else None)
        cell = ws.cell(row=tr, column=c, value=v)
        cell.font = Font(name=FONT, size=9, bold=True)
        cell.fill = _fill(C_REF_HDR)
        cell.border = BORDER
        cell.alignment = left if isinstance(v, str) else center
        if fmt and isinstance(v, (int, float)):
            cell.number_format = fmt

    for c, (label, width, _k, _f) in enumerate(cols, start=1):
        cell = ws.cell(row=hr, column=c, value=label)
        cell.font = Font(name=FONT, size=9, bold=True)
        cell.fill = _fill(C_SUBHDR)
        cell.border = BORDER
        cell.alignment = center
        ws.column_dimensions[get_column_letter(c)].width = width
    ws.row_dimensions[hr].height = 26

    r = hr + 1
    for row in rows:
        for c, (_l, _w, k, fmt) in enumerate(cols, start=1):
            v = row.get(k)
            if v == "":
                v = None
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = Font(name=FONT, size=9)
            cell.alignment = left if isinstance(v, str) else center
            cell.border = BORDER
            if fmt and isinstance(v, (int, float)):
                cell.number_format = fmt
            # the two figures a buyer scans for: red when selling behind plan, and a
            # bold NPD flag
            if k == "ytdPct" and isinstance(v, (int, float)):
                cell.font = Font(name=FONT, size=9, bold=abs(v) >= 15,
                                 color="FFB23B2C" if v <= -15 else "FF1A7E34" if v >= 15 else "FF000000")
            elif k == "npd" and v:
                cell.font = Font(name=FONT, size=9, bold=True, color="FF5B3FA8")
        r += 1

    ws.freeze_panes = ws.cell(row=hr + 1, column=3)   # keep code + description in view
    ws.auto_filter.ref = f"A{hr}:{get_column_letter(len(cols))}{max(r - 1, hr)}"

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


def export_arrivals(payload):
    """"Upcoming Containers" workbook shared from the Arrivals page. One row per
    outstanding product line, flat and autofiltered so the team can sort/pivot:
    sheet 1 = booked Qlik containers by arrival date, sheet 2 = outstanding POs
    awaiting a booking. `payload` is the client's already-joined view (its
    buildArrivalEvents(), filtered as shown on screen), so the export always
    matches the page exactly."""
    wb = Workbook()
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")

    def _date(iso):
        try:
            y, m, d = (int(x) for x in str(iso)[:10].split("-"))
            return date(y, m, d)
        except (ValueError, AttributeError, TypeError):
            return None

    C_WEEK_BAND = "FFE9F0F9"   # soft blue – shades every other calendar week's rows
    week_rule = Side(style="medium", color="FFA9BFDB")   # stronger rule where a new week starts

    def _week_key(row):   # default banding: Mon–Sun calendar week of the row's date
        dv = row.get("date")
        return dv.isocalendar()[:2] if isinstance(dv, date) else None

    def sheet(ws, title_txt, cols, rows, note, months=None, months_title="Monthly totals", band_key=None):
        band_key = band_key or _week_key
        ws.cell(row=1, column=1, value=title_txt).font = Font(name=FONT, size=13, bold=True)
        ws.cell(row=2, column=1, value=note).font = Font(name=FONT, size=9, italic=True, color="FF808080")
        hr = 4  # header row
        if months:
            # prominent per-month totals — one row per month actually present in the
            # data, so the block tracks whatever window the file covers. Each entry is
            # (label, col4_text, col5_text) so booked ("N containers"/"units") and
            # awaiting ("N to book"/"units") can share the same layout.
            ws.cell(row=4, column=1, value=months_title).font = Font(name=FONT, size=10, bold=True)
            for i, (label, c4_text, c5_text) in enumerate(months):
                rr = 5 + i
                lc = ws.cell(row=rr, column=1, value=label)
                lc.font = Font(name=FONT, size=10, bold=True)
                lc.alignment = left
                for col in range(1, 4):   # fill every cell of the merge-to-be
                    ws.cell(row=rr, column=col).fill = _fill(C_REF_HDR)
                    ws.cell(row=rr, column=col).border = BORDER
                cc = ws.cell(row=rr, column=4, value=c4_text)
                cc.font = Font(name=FONT, size=10, bold=True)
                cc.border = BORDER
                uc = ws.cell(row=rr, column=5, value=c5_text)
                uc.font = Font(name=FONT, size=10)
                uc.alignment = left
                for col in range(5, 8):
                    ws.cell(row=rr, column=col).border = BORDER
                ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=3)
                ws.merge_cells(start_row=rr, start_column=5, end_row=rr, end_column=7)
            hr = 5 + len(months) + 1
        for c, (label, width, _) in enumerate(cols, start=1):
            cell = ws.cell(row=hr, column=c, value=label)
            cell.font = Font(name=FONT, size=9, bold=True)
            cell.fill = _fill(C_SUBHDR)
            cell.border = BORDER
            cell.alignment = center
            ws.column_dimensions[get_column_letter(c)].width = width
        r = hr + 1  # first data row
        prev_group = object()
        prev_week = object()
        band = False
        band_fill = _fill(C_WEEK_BAND)
        for row in rows:
            new_group = row.get("_group") != prev_group
            prev_group = row.get("_group")
            # band_key drives the alternating shading + strong separating rule: the
            # booked sheet bands by calendar week, the awaiting sheet by booking month
            bkey = band_key(row)
            new_week = bkey != prev_week
            if new_week:
                band = not band
                prev_week = bkey
            for c, (_, _, k) in enumerate(cols, start=1):
                v = row.get(k)
                cell = ws.cell(row=r, column=c, value=v)
                cell.font = Font(name=FONT, size=9)
                cell.alignment = left if isinstance(v, str) else center
                if isinstance(v, date):
                    cell.number_format = "dd-mm-yyyy"
                elif isinstance(v, (int, float)):
                    cell.number_format = "#,##0.0" if k in ("cbm", "containers") else "#,##0"
                # payment / overdue status text colouring (matches the page's badges)
                if k == "status" and isinstance(v, str) and v:
                    up = v.upper()
                    if "NOT" in up:
                        cell.font = Font(name=FONT, size=9, bold=True, color="FFC55A11")  # orange – NOT PAID
                    elif "PAID" in up:
                        cell.font = Font(name=FONT, size=9, bold=True, color="FF1A7E34")  # green – PAID
                elif k == "overdue" and v == "OVERDUE":
                    cell.font = Font(name=FONT, size=9, bold=True, color="FFB23B2C")      # red – overdue PO
                if band:
                    cell.fill = band_fill
                if new_week:      # strong rule where a new arrival week begins
                    cell.border = Border(top=week_rule)
                elif new_group:   # thin rule per container/PO group within the week
                    cell.border = Border(top=thin)
            r += 1
        ws.freeze_panes = ws.cell(row=hr + 1, column=1)
        ws.auto_filter.ref = f"A{hr}:{get_column_letter(len(cols))}{max(r - 1, hr)}"

    booked_cols = [
        ("Arrival", 9, "date"), ("Wk", 5, "week"), ("PO", 11, "po"),
        ("Supplier", 30, "supplier"), ("Container", 14, "container"), ("Status", 12, "status"),
        ("ETD", 9, "etd"), ("ETA UK Port", 11, "etaPort"), ("Delivery to CB", 13, "deliveryCB"),
        ("Product", 18, "code"), ("Description", 36, "name"), ("Season", 16, "season"),
        ("Arrival Units", 12, "qty"), ("Current Stock", 12, "stock"),
    ]
    rows = []
    for ev in payload.get("booked") or []:
        for ln in ev.get("lines") or [{}]:
            rows.append({
                "_group": (ev.get("po") or "") + (ev.get("container") or ""),
                "date": _date(ev.get("date")), "week": ev.get("week"), "po": ev.get("po", ""),
                "supplier": ev.get("supplier", ""), "container": ev.get("container", ""),
                "status": (ev.get("status") or "") + (" · SPLIT" if ev.get("split") else ""),
                "etd": _date(ev.get("etd")), "etaPort": _date(ev.get("etaPort")),
                "deliveryCB": _date(ev.get("deliveryCB")),
                "code": ln.get("code", ""), "name": ln.get("name", ""), "season": ln.get("season", ""),
                "qty": ln.get("qty"), "stock": ln.get("stock"),
            })
    # monthly totals derived from the rows themselves, so the block always follows
    # the arrival dates actually present in the uploaded Qlik file (months with no
    # arrivals simply don't appear); containers deduped by container no (PO fallback)
    MONTH_NAMES = ["January", "February", "March", "April", "May", "June",
                   "July", "August", "September", "October", "November", "December"]
    mon = {}
    for row in rows:
        dv = row.get("date")
        if not isinstance(dv, date):
            continue
        g = mon.setdefault((dv.year, dv.month), {"cont": set(), "units": 0.0})
        g["cont"].add(row.get("container") or row.get("po"))
        g["units"] += row.get("qty") or 0
    month_totals = [(f"{MONTH_NAMES[m - 1]} {y}",
                     f"{len(g['cont'])} container{'s' if len(g['cont']) != 1 else ''}",
                     f"{int(round(g['units'])):,} arrival units")
                    for (y, m), g in sorted(mon.items())]
    note = (f"Generated {payload.get('generated', '')} · arrival units = ordered − delivered (WEBSA Open PO)"
            " · arrival = delivery-to-CB, else UK-port ETA (Qlik)")
    ws = wb.active
    ws.title = "Upcoming Containers"
    sheet(ws, "Upcoming container arrivals", booked_cols, rows, note, months=month_totals)

    await_cols = [
        ("WEBSA Due", 10, "date"), ("Est. Booking", 11, "bookDate"), ("PO", 11, "po"),
        ("Supplier", 28, "supplier"), ("Overdue", 9, "overdue"),
        ("Product", 18, "code"), ("Description", 36, "name"), ("Season", 16, "season"),
        ("Arrival Units", 12, "qty"), ("Current Stock", 12, "stock"),
    ]
    arows = []
    for ev in payload.get("awaiting") or []:
        for ln in ev.get("lines") or [{}]:
            arows.append({
                "_group": ev.get("po", ""),
                "date": _date(ev.get("date")), "bookDate": _date(ev.get("bookDate")),
                "po": ev.get("po", ""), "supplier": ev.get("supplier", ""),
                "overdue": "OVERDUE" if ev.get("overdue") else "",
                "code": ln.get("code", ""), "name": ln.get("name", ""), "season": ln.get("season", ""),
                "qty": ln.get("qty"), "stock": ln.get("stock"),
            })
    # "to book" totals per estimated-booking month (due − 64d): how many containers
    # must be booked each month, deduped by PO
    amon = {}
    for row in arows:
        dv = row.get("bookDate")
        key = (dv.year, dv.month) if isinstance(dv, date) else None
        g = amon.setdefault(key, {"pos": set(), "units": 0.0})
        g["pos"].add(row.get("po"))
        g["units"] += row.get("qty") or 0
    await_month_totals = [
        ("No est. booking date" if k is None else f"{MONTH_NAMES[k[1] - 1]} {k[0]}",
         f"{len(g['pos'])} to book", f"{int(round(g['units'])):,} units")
        for k, g in sorted(amon.items(), key=lambda kv: (kv[0] is None, kv[0] or (0, 0)))]

    def _await_band(row):   # band the awaiting sheet by estimated booking month
        dv = row.get("bookDate")
        return (dv.year, dv.month) if isinstance(dv, date) else None

    lead = payload.get("lead") or {}
    ls = int(lead.get("sailing", 50)); lg = int(lead.get("grace", 7)); li = int(lead.get("inland", 7))
    lead_note = (f"est. booking = WEBSA due − {ls + lg + li} days "
                 f"({ls} sailing + {lg} factory grace + {li} UK inland)")
    sheet(wb.create_sheet("Awaiting Booking"), "Outstanding POs awaiting a container booking", await_cols, arows,
          f"POs with outstanding balance but no dated container in the Qlik export · {lead_note}, grouped by booking month.",
          months=await_month_totals, months_title="Containers to book per month", band_key=_await_band)

    # ---- optional plan-side sheets (the "Include plan orders" toggle) ----
    # Committed stock with no PO raised + proposed rebuys, straight off the plan, so the
    # booking outlook can be seen before anything is officially raised. Containers are an
    # estimate from CBM — nothing here is booked, so there is no real container count.
    plan = payload.get("plan") or []
    if plan:
        plan_cols = [
            ("Arrival w/c", 11, "date"), ("Wk", 5, "week"), ("Est. Booking", 11, "bookDate"),
            ("Type", 20, "kind"), ("Supplier", 28, "supplier"),
            ("Product", 18, "code"), ("Description", 36, "name"), ("Season", 16, "season"),
            ("Arrival Units", 12, "qty"), ("CBM", 9, "cbm"), ("Current Stock", 12, "stock"),
        ]
        prows = []
        for ev in plan:
            kind = "Proposed rebuy" if ev.get("kind") == "proposed" else "Committed - no PO"
            for ln in ev.get("lines") or [{}]:
                prows.append({
                    "_group": f"{ev.get('supplier','')}|{ev.get('week')}|{ev.get('kind')}",
                    "date": _date(ev.get("date")), "week": ev.get("week"),
                    "bookDate": _date(ev.get("bookDate")), "kind": kind,
                    "supplier": ev.get("supplier", ""),
                    "code": ln.get("code", ""), "name": ln.get("name", ""), "season": ln.get("season", ""),
                    "qty": ln.get("qty"), "cbm": ln.get("cbm"), "stock": ln.get("stock"),
                })
        ccbm = float(payload.get("containerCbm") or 68) or 68
        pmon = {}
        for row in prows:
            dv = row.get("bookDate")
            key = (dv.year, dv.month) if isinstance(dv, date) else None
            g = pmon.setdefault(key, {"cbm": 0.0, "units": 0.0})
            g["cbm"] += row.get("cbm") or 0
            g["units"] += row.get("qty") or 0
        plan_month_totals = [
            ("No est. booking date" if k is None else f"{MONTH_NAMES[k[1] - 1]} {k[0]}",
             f"~{g['cbm'] / ccbm:.1f} containers", f"{int(round(g['units'])):,} units ({g['cbm']:,.1f} cbm)")
            for k, g in sorted(pmon.items(), key=lambda kv: (kv[0] is None, kv[0] or (0, 0)))]

        def _plan_band(row):   # band by estimated booking month, like the awaiting sheet
            dv = row.get("bookDate")
            return (dv.year, dv.month) if isinstance(dv, date) else None

        sheet(wb.create_sheet("Plan Orders (No PO)"),
              f"{payload.get('year', '')} plan orders with no PO raised".strip(),
              plan_cols, prows,
              "Committed order units in a week with no matching PO/container, plus all proposed rebuys, "
              f"from the current week on · est. booking = arrival week − {ls + lg + li} days · "
              f"containers estimated at {ccbm:g} cbm each — nothing here is booked.",
              months=plan_month_totals, months_title="Plan orders to book per month", band_key=_plan_band)

        # ---- Full Outlook: one row per month, booked arrivals alongside everything
        # still to be booked (real POs + plan orders), so the whole forward view reads
        # off a single table.
        ws = wb.create_sheet("Full Outlook", 0)
        ws.cell(row=1, column=1, value="Full container outlook by month").font = Font(name=FONT, size=13, bold=True)
        ws.cell(row=2, column=1, value=(
            f"Generated {payload.get('generated', '')} · booked = dated Qlik containers, by arrival month · "
            "to book = outstanding POs and plan orders, by ESTIMATED BOOKING month (arrival − lead) · "
            f"plan containers estimated from CBM at {ccbm:g} cbm each")
        ).font = Font(name=FONT, size=9, italic=True, color="FF808080")
        oc = [("Month", 18), ("Booked containers arriving", 15), ("Booked units", 13),
              ("POs to book", 11), ("PO units", 12),
              ("Plan orders to book (est. containers)", 16), ("Plan units", 12),
              ("Total units to book", 15)]
        hr = 4
        for c, (label, width) in enumerate(oc, start=1):
            cell = ws.cell(row=hr, column=c, value=label)
            cell.font = Font(name=FONT, size=9, bold=True)
            cell.fill = _fill(C_SUBHDR)
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(c)].width = width
        ws.row_dimensions[hr].height = 30

        def _mkey(label):   # "January 2027" -> (2027, 1); the "no date" buckets sort last
            try:
                nm, yr = str(label).rsplit(" ", 1)
                return (int(yr), MONTH_NAMES.index(nm) + 1)
            except (ValueError, IndexError):
                return None

        cols_by_month = {}
        for label, cont, units in month_totals:          # booked arrivals
            cols_by_month.setdefault(_mkey(label), {})["booked"] = (cont, units)
        for label, pos, units in await_month_totals:     # POs awaiting a booking
            cols_by_month.setdefault(_mkey(label), {})["po"] = (pos, units)
        for label, cont, units in plan_month_totals:     # plan orders
            cols_by_month.setdefault(_mkey(label), {})["plan"] = (cont, units)

        def _n(txt):        # first number out of "12 containers" / "3,450 units (12.3 cbm)"
            m = re.match(r"~?([\d,]+(?:\.\d+)?)", str(txt or ""))
            return float(m.group(1).replace(",", "")) if m else 0.0

        r = hr + 1
        tot = [0.0] * 6
        for k in sorted(cols_by_month, key=lambda x: (x is None, x or (0, 0))):
            g = cols_by_month[k]
            label = "No date" if k is None else f"{MONTH_NAMES[k[1] - 1]} {k[0]}"
            vals = [_n(g.get("booked", ("", ""))[0]), _n(g.get("booked", ("", ""))[1]),
                    _n(g.get("po", ("", ""))[0]), _n(g.get("po", ("", ""))[1]),
                    _n(g.get("plan", ("", ""))[0]), _n(g.get("plan", ("", ""))[1])]
            tot = [a + b for a, b in zip(tot, vals)]
            for c, v in enumerate([label] + vals + [vals[3] + vals[5]], start=1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.font = Font(name=FONT, size=10)
                cell.border = BORDER
                if isinstance(v, float):
                    cell.number_format = "#,##0.0" if c == 6 else "#,##0"
            r += 1
        for c, v in enumerate(["Total"] + tot + [tot[3] + tot[5]], start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = Font(name=FONT, size=10, bold=True)
            cell.fill = _fill(C_REF_HDR)
            cell.border = BORDER
            if isinstance(v, float):
                cell.number_format = "#,##0.0" if c == 6 else "#,##0"
        ws.freeze_panes = ws.cell(row=hr + 1, column=1)

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


# ---------------------------------------------------------------------------
def parse_supplier_form(path_or_bytes):
    """Read a scheduling form and return {sku_code: {week: qty}} from the
    New Order columns, mapping each column to its arrival week."""
    wb = load_workbook(path_or_bytes, data_only=True)
    ws = wb["Scheduling"] if "Scheduling" in wb.sheetnames else wb.worksheets[0]

    # locate header row (col A == 'SKU') and the sub-header row (has 'New Order')
    hdr_row = None
    for r in range(1, 12):
        if str(ws.cell(row=r, column=1).value).strip().lower() == "sku":
            hdr_row = r
            break
    if hdr_row is None:
        hdr_row = 1
    sub_row = None
    for r in range(hdr_row, hdr_row + 8):
        for c in range(1, ws.max_column + 1):
            if str(ws.cell(row=r, column=c).value).strip().lower() == "new order":
                sub_row = r
                break
        if sub_row:
            break
    if sub_row is None:
        raise ValueError("Could not find a 'New Order' header row – is this a scheduling form?")

    # map columns -> week using any 'Week N' text in the header rows (incl. merged spans)
    col_week = {}
    wk_re = re.compile(r"week\s*#?\s*(\d+)", re.I)
    spans = list(ws.merged_cells.ranges)
    for r in range(1, sub_row):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if not isinstance(v, str):
                continue
            m = wk_re.search(v)
            if not m:
                continue
            wk = int(m.group(1))
            # find the merged span this cell belongs to (else just this column)
            lo, hi = c, c
            for sp in spans:
                if sp.min_row <= r <= sp.max_row and sp.min_col <= c <= sp.max_col:
                    lo, hi = sp.min_col, sp.max_col
                    break
            for cc in range(lo, hi + 1):
                col_week[cc] = wk

    # New Order columns
    no_cols = [c for c in range(1, ws.max_column + 1)
               if str(ws.cell(row=sub_row, column=c).value).strip().lower() == "new order"]

    updates = {}
    weeks_seen = set()
    r = sub_row + 1
    while r <= ws.max_row:
        code = ws.cell(row=r, column=1).value
        if isinstance(code, str):
            cl = code.strip()
            if cl.lower() in ("totals", "total", ""):
                break
            if cl:
                row_upd = {}
                for c in no_cols:
                    wk = col_week.get(c)
                    if not wk:
                        continue
                    val = ws.cell(row=r, column=c).value
                    qty = int(round(val)) if isinstance(val, (int, float)) else 0
                    row_upd[wk] = row_upd.get(wk, 0) + qty
                    weeks_seen.add(wk)
                if row_upd:
                    updates[cl] = row_upd
        r += 1
    return {"orders": updates, "weeks": sorted(weeks_seen)}


if __name__ == "__main__":
    import sys
    sup = sys.argv[1] if len(sys.argv) > 1 else "Midan Global Ltd"
    bio = export_supplier(sup)
    out = Path(f"_test_{sup.split()[0]}.xlsx")
    out.write_bytes(bio.getvalue())
    print("wrote", out, out.stat().st_size, "bytes")
    res = parse_supplier_form(str(out))
    print("parsed", len(res["orders"]), "SKUs, weeks", res["weeks"][:8], "...")
    # round-trip check vs orders.json
    master = json.loads((DATA / "master.json").read_text())
    orders = json.loads((DATA / "orders.json").read_text())
    skus = {s["code"]: s for s in master["skus"] if s["supplier"] == sup}
    mism = 0
    for code, wkmap in res["orders"].items():
        s = skus.get(code)
        if not s:
            continue
        ov = orders.get(s["id"], [0] * WEEKS)
        for wk, qty in wkmap.items():
            if ov[wk - 1] != qty:
                mism += 1
                if mism <= 5:
                    print(f"  mismatch {code} wk{wk}: form={qty} app={ov[wk-1]}")
    print("round-trip mismatches:", mism)
