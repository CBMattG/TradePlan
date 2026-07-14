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
