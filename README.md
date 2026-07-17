# 2026 Tradeplan (standalone app)

A standalone replacement for `2026 Tradeplan - V2 - Forecast Unlinked.xlsx`.
Edit weekly **Committed Orders** quantities per SKU and instantly see the effect on
stock closing, weeks cover, sales value, FOB spend, CBM/containers, stock holding
value, supplier value splits and warehouse capacity.

No installs needed beyond Python (already on this machine). No internet required.

## Run it

Double-click **run.bat** (or `python server.py`). The browser opens at
http://localhost:8765 automatically.

## Multiple years

The app is the master for every plan year. The **year selector** (top-left) switches
all data — sales, forecast, orders, stock, suppliers — between years held under
`data/<year>/`. Each year keeps its own orders, current week and **forecast mode**
(Original / Seasonality / Target — set independently per year, see *Forecast mode* below).
The multiplier, capacities, cover bands and shared model tuning (weather, own-shape weight)
apply across years. Supplier export/import and CSV export all act on the selected year.

Import a year from its Excel workbook:

    python import_data.py --year 2026        # 2026 layout (default path)
    python import_data.py --year 2025         # 2025 layout (auto-detected)
    python import_data.py --year 2025 "path\to\2025 file.xlsx"

The importer auto-handles the slightly different 2025 sheet layout. Historical years
(e.g. 2025) load with their current week set to 53 so all weeks show actuals.

**Generate a forecast year** — the build button (next to the year selector) targets
the year **after the one you're viewing**: from 2026 it reads **↻ 2027** (rebuild)
or, if that year doesn't exist yet, **+ 2027** (build new). It pulls in each
product's **forecasted weekly sales** and **projected stock** starting from the viewed
year's projected year-end stock. The rolled-over sales forecast is **always rebuilt
from sales patterns** — each product's recency-weighted demand (this year's
sell-through + stock-out-corrected last year + the original planner forecast as a
stabiliser) shaped by its seasonal curve — so fast movers and lines that sold out rise
and slow movers decline, rather than the prior year being copied verbatim. (This base
rebuild is independent of each year's **Forecast mode**, which re-models a year's
forecast on top of this rolled-over base.) The
new year is a real, browsable, editable year — view the sales model per supplier, run
its own rebuy plan, and export supplier forms for it. **Rebuilding** lets a corrected
base year flow forward into an existing forecast year (it regenerates that year's
forecast, stock and rebuys from scratch); it will **never overwrite an imported
(real) year** — only generated forecast years can be rebuilt.

## Using the app

- **Rebuy suggestions (in the Plan grid)** — the auto-rebuy schedule is built into the
  Plan page. Every product shows a teal **"Proposed Rebuy"** row (editable) with the
  algorithm's suggested **additional** orders on top of your committed orders, and the
  Stock Closing / Weeks Cover / panel reflect committed + proposed **live** so you see
  the impact as you edit (and as you switch year). It simulates each Live product's
  stock from current free stock + the (seasonality-modelled) forecast, rolling into
  next year, to hold the target **weeks cover**, timed by **lead time** (China ~11 wk /
  non-China ~3, by origin). The rebuy bar toggles **Full 68 CBM only / Allow partial**
  (full ships ~68 CBM containers; partial builds up to a configurable ~28 CBM half-load)
  and shows the whole-plan proposed total. On a **forecast year** the algorithm can
  also place orders arriving in the **first weeks** of the year (weeks 1–11) — those
  are placed in the prior year's tail (an 11-week China order placed in late December
  lands in early January), so an early-year stock-out with forecast demand now gets a
  rebuy instead of being left uncovered.
  - It keeps each line continuously stocked through its **active season**: when a
    line is about to run out within its lead time, the container is shipped so it
    **arrives by the stock-out** rather than after it (no out-of-stock gap before the
    next container lands).
  - In **Full 68 CBM** mode, when a shipment is forced out early to beat a stock-out,
    the container is **topped up toward a full 68 CBM** by pulling the supplier's
    soonest upcoming demand forward — so a container you're paying for anyway goes out
    close to full instead of half-empty. It only adds units genuinely needed later
    (never invents demand) and respects the seasonal off-season guard, so suppliers
    with many small-CBM lines (e.g. mirrors/bistro sets) consolidate into fewer full
    containers instead of lots of part-full weekly shipments. Partial mode is unchanged.
  - For **seasonal lines** it stops restocking into the off-season: once an arrival
    would land when demand has fallen below a set fraction of the line's peak week, it
    leaves the line to **run down** instead of carrying off-season stock (e.g. summer
    parasols arriving in winter) — unless that period's demand is genuinely strong.
  - **Rebuy parameters** live in **Settings → Rebuy**: the default **weeks cover
    target**; a **weeks-cover-by-season** override (e.g. raise Summer so seasonal
    lines build more buffer as their demand ramps, capturing sharp in-season spikes);
    an **extra cover when in-season** for seasonal products — added across the whole
    in-season period (scaled by how far above the year-round average demand runs) and
    timed by lead so the buffer is built up *as the season approaches*, not after the
    peak; a **"stop seasonal rebuys below X% of peak demand"** threshold (default 40%
    — roughly "no summer arrivals past August"; raise it to stop earlier, lower it to
    keep buying further into the shoulder); a **partial-container target** (default
    ~28 CBM) so that in "Allow partial" mode shipments build up to a worthwhile
    half-load before going, rather than lots of tiny amounts; the partial-container
    **max wait** (ship after this many weeks even if the target isn't reached, topped
    up toward it); and **one product per container** — pick the suppliers whose MOQ is
    a full container of a single product, and the algorithm orders each of their
    products alone in whole-container (~68 CBM) lots instead of mixing products
    (removing the manual edits those suppliers used to need).
  - **⟳ Run rebuy** re-runs the algorithm now and refills the Proposed Rebuy row from
    the **current committed orders** — so you can step through suppliers: commit a few,
    then re-run to see the plan adjust (a supplier whose committed orders now meet
    demand stops being proposed). **Clear suggestions** empties the Proposed Rebuy row.
  - **Apply to: This supplier / Whole year** — a scope toggle in the rebuy bar governs
    all three actions (Run rebuy, Clear suggestions, Commit). "This supplier" (the
    default) acts only on the supplier you're viewing, leaving every other supplier's
    proposals untouched; "Whole year" acts across all suppliers. Your choice is
    remembered. (This replaces the separate per-supplier commit button.)
  - Each year **remembers its own proposed-rebuy state**: if you clear or edit the
    suggestions in one year, switch to another and come back, your changes are still
    there (they're auto-built only on a year's first visit, then saved per year). Use
    **⟳ Run rebuy** whenever you want to regenerate them from the latest forecast.
  - **Commit** (scoped by the toggle above) folds the proposed units into the Committed
    Orders (they become the blue committed numbers) and flow into supplier exports.
  - **Export / Import** include the proposed rebuys: Export to Excel gives the factory
    form with committed + proposed; edit it and Import to confirm those as committed
    orders. Proposals then rebuild on top.
- **Supplier KPIs** — each supplier in the list shows its **forecast sales value**
  (demand × price, uncapped — meaningful even with no committed orders) and a
  colour-coded **stocked-in %** (committed supply vs forecast demand: red &lt;100%
  needs orders, green ~100%, blue &gt;100% overstock).
- **Save config / Load config** — the **💾 Save config** button (top bar) asks for a
  **name** and a **scope** (all years, or just the current year) and saves that
  configuration under the name (in `data/configs/`): the chosen years' order quantities
  and rebuy plans, plus the global sales-forecast settings. The **Load config** dropdown
  (top bar) lists your saved configurations; picking one lets you choose to restore
  **all years in it or a single year**. Your working changes still save automatically
  and reload when you reopen the app; named saves are checkpoints you can return to.
- **Revert orders** — the **↺ Revert orders** button (Plan view, supplier header) lets
  you choose a **scope** (this year, or all years) and what to revert to:
  - **Most recent save** — restores from your latest named **Save config** (shown with
    its name and time; disabled until you've saved one).
  - **Original Excel file** — resets order quantities to the figures first imported from
    the spreadsheet and clears all rebuy suggestions.
- **Current-week highlight** — the green "this week" column tracks the **real current
  week** of whichever year actually contains today (e.g. week 26 / w/c 22 Jun in the
  2026 sheet, advancing each Monday). Past and future years show no week highlight.
- **Forecast-year stock & LY chaining** — a generated forecast year (e.g. 2027) starts each
  product from the **prior year's current projected year-end stock** (committed orders
  **and** proposed rebuys — the same Stock Closing the grid shows), and its **"LY Sales"
  row is re-derived live** from the prior year's **realised actuals where they exist,
  then its forecast for the rest of the year** (actuals run to the prior year's imported
  data week). So if you change 2026 — edit the Committed Orders or Proposed Rebuy row, or
  **import an updated week of actual sales** — switching to 2027 re-derives its starting
  stock *and* its LY reference from the new 2026 position and **re-runs the rebuys** to
  show the knock-on effect, instead of keeping a stale snapshot. (The 2027 **sales
  forecast** itself — the plan — refreshes when you **rebuild** the year with ↻.) In the
  **LY Sales row of a forecast year**, the weeks that are still the prior year's
  **forecast** (rather than banked actuals) are **tinted** (a subtle blue-grey italic),
  so you can see exactly where the real sales reporting stops and the forecast takes over.
  The **"Sales Forecast" row** keeps showing the **original weekly forecast** in every
  cell, but its **Total** is the full-year **outturn** — actual sales for the weeks that
  have them + forecast for the rest — so it reads as where the year will land and equals
  the **next** year's LY Sales total. A small **"plan N"** sub-value under that total
  keeps the pure full-year forecast visible for tracking how each line is tracking vs
  target. (Underneath, the pure forecast still drives stock, cover, stock-out and rebuy.)
- **Whole-plan totals** — the band across the top of the page always shows the
  overall sheet numbers (forecast sales, FOB, units, volume, containers, peak
  stock/racking/stillage) no matter which supplier you're looking at, and updates
  live as you edit. Forecast sales, FOB, units, volume and containers are shown
  as **committed / proposed / total** — the big figure is the **total** (committed
  orders **plus** the proposed rebuys), with a sub-line splitting out the committed
  base and the proposed-rebuy addition (in teal). So **before committing** anything
  you can see what the whole-sheet sales and FOB spend would become if the
  suggested rebuys were placed. (Forecast sales is stock-aware — adding proposed
  orders lifts sales only where extra stock lets more forecast demand be met;
  FOB/units/volume/containers are the direct order additions.) Committing folds the
  proposed figures into the committed ones.
- **Rebuy status (Live / Not Live)** — each product carries its catalog status
  from the Buying Report: a green **Live** badge (plan for rebuy) or a grey
  **Not Live** badge (discontinued — run the remaining stock down, don't rebuy).
  Not Live rows are dimmed. The supplier header shows the Live/Not Live split, the
  whole-plan band shows the totals, and the **Rebuy status** buttons filter the
  grid to All / Live / Not Live so you can focus on what needs reordering.
  (SKUs not found in the Buying Report show an amber "Unknown" badge.)
- **Suppliers list** (left) — the £ beside each supplier is its **total forecast
  sales for 2026**. Sort the list by that value (high→low or low→high), A–Z, or
  back to the original workbook order using the buttons at the top. Your chosen
  **sort order**, the **visible rows** (the row-toggle chips), the **Rebuy status
  filter** and the **selected supplier** are all remembered — they stay put when
  you switch year and when you reopen the app. (If the selected supplier doesn't
  exist in the year you switch to, it falls back to the first supplier.)
- **Export / Import supplier form (Excel)** — the buttons in each supplier's
  header download a factory-style order-planning workbook (the "Scheduling" form:
  SKU, FOB, CBM, proposed order qty, plus a New Order / Value / CBM /
  Remaining-in-factory block per arrival week). The form shows **only the weeks
  that currently have a planned order** and includes the committed orders **plus
  any proposed rebuys**. The **year labels in the reference columns match the year
  you're exporting** (e.g. a 2027 form reads "2026 FOB / 2027 FOB", "2026 Volume",
  "2027 Proposed Total Order QTY") rather than being fixed. The previous-year
  **Volume** column shows that year's actual **order units** (what was ordered),
  matched by SKU code from the prior year's data — not last year's sales. At the bottom is a **cross-year summary** — Order
  Quantity, Order Value and Number of Containers for **every year held in the app**
  (matched to this supplier's products by SKU code, since names differ between
  years), one column per year and growing automatically as new years are added.
  The exported year's column is **live** (tied to the form totals, so it updates
  as the factory edits the New Order cells); the other years are that year's order
  plan. The **Variance #** / **Variance %** columns compare the two most recent
  years (e.g. 2026 → 2027); the % is `(new − old) / old`, so a drop shows a
  negative percentage. **Export to Excel** downloads the active supplier's form;
  **Export multiple…** opens a picker to export **all suppliers or any selection**
  at once — one form per supplier, bundled into a single **.zip** so you get one
  download instead of a prompt per supplier. Suppliers with **no orders planned**
  (nothing committed or proposed for the year) are **skipped**, and a dialog lists
  exactly which ones were left out. Send the form(s) to the factory, let
  them adjust quantities, then **Import from Excel** makes the app mirror the file
  exactly: each product's order forecast is set to match the form, and any week
  that is 0/blank in the form (or no longer shown) is set back to blank. Other
  suppliers are untouched. The importer also accepts the original container-style
  forms (it maps each container's order to its "Estimated Arrival: Week N" header).
- **Supplier summary panel** — above the grid, five colour-coded cards show the
  supplier's full-year Sales, Order FOB, Order Units, CBM (with container estimate)
  and peak Stock Holding, each with a weekly sparkline (current week marked) so you
  can see the timing/seasonality at a glance. Updates live as you edit orders.
- **Plan** tab — pick a supplier on the left. The yellow **Committed Orders** row is
  editable; type a quantity and everything recalculates immediately. Use
  arrow keys / Enter to move between cells, and you can paste a row of numbers
  straight from Excel. Blue bold numbers mean the value differs from what was
  originally imported from the spreadsheet. As you scroll right across the weeks,
  each product's header (code, name and cost details) stays pinned on the left.
- A pinned **Order CBM footer** runs along the bottom of the grid: the supplier's total
  order volume (committed + proposed) for **each week**, colour-coded by how close it is
  to filling a container — amber while building up, green at roughly a full container,
  blue once it's over one — so you can see at a glance which weeks are light, full or
  over (hover a week for the exact CBM, container fraction and committed/proposed split).
- Red cells = projected stock-out (zero stock with forecast demand).
  Amber weeks-cover = under 2 weeks. Green column = current week.
- **Summary** tab — weekly business totals: sales, FOB, CBM, containers, stock
  holding, pallet/stillage/racking spaces and warehouse utilisation, plus
  quarterly sales vs 2023/2024.
- **Value Splits** tab — every supplier's share of forecast sales, FOB, CBM and
  containers. Click a supplier to jump to its plan.
- **Forecast mode — per year** (Settings → Forecast) — each year builds its Sales
  Forecast **independently**. The Forecast settings start with a row of **year tabs**
  (each showing that year's current mode); pick a year, then choose one of three modes
  for it. So you can leave **2026 on Original** but run **Seasonality with a growth %**
  on **2027** — there's no single global toggle any more. **Reset this year to Original**
  drops the selected year back to the imported forecast. Past weeks are never changed.
  - **Original** — use the imported / rolled-over forecast as-is.
  - **Seasonality model** — re-models the future forecast from each product's **season**
    (Summer/Winter/Continuity, from the Buying Report) and **category**. The level is
    anchored on this year's actual sell-through and last year's sales, with **last-year
    peak-season stock-outs detected** (0 sales in a peak week with sales either side) and
    excluded so a strong line isn't under-bought, and **seasonal peaks calibrated from the
    aggregated sales history** (e.g. Outdoor Heating peaks in winter). The **Own-shape
    weight** (a shared setting, default 60%) blends each product's **own** last-year shape
    into its profile, so a sharp single-product spike keeps its shape instead of being
    smeared flat — and **Continuity lines pick up their own within-year uplifts/drops**
    rather than being flat-lined. The model **down-weights this year's run-rate until
    enough annual demand has been observed** (so a Christmas range seen only in its quiet
    off-season isn't under-forecast), and on a **generated forecast year leans on realised
    sales rather than the previous model's rollover**, so lines that outsold their old plan
    aren't dragged down. A **Model strength** slider blends model vs original; a **live
    weather forecast** (Open-Meteo, no key) nudges the nearest ~2 weeks. **Expected
    year-on-year growth %** scales that year's whole forecast up/down for expected
    business growth — because the model otherwise runs roughly **flat vs last year's
    actual sales**, this is how you carry a growth assumption in the *main* forecast (e.g.
    +7% to lift 2027 above its flat ~£12.6m baseline) **without** using Target mode.
  - **Target sales** — set an overall **£ sales target** and the forecast is rebuilt to
    hit it: distributed across products by **recency-weighted historical sales**, shaped
    by each product's seasonality/category curve, normalised so the whole-sheet forecast
    **value** lands on the target. Use it to **scope a business target** (is £15m
    realistic?) and see, per product, the uplift it implies plus the knock-on rebuys (the
    Proposed Rebuy suggestions rebuild against the raised forecast). It's for scoping, not
    the day-to-day forecast — the growth % above is how you keep growth in the main
    forecast. *Note:* the whole-plan band's total sales is **stock-capped sell-through**
    (`min(forecast, stock)`), so it sits a little below the demand target — that's what
    you'd actually sell given stock timing; compare the target against the uncapped
    forecast (the "plan N" sub-value on the Sales Forecast total).
  - A **banner** on the plan shows the active mode for the year, with **adjust** / **turn
    off** links. **Shared model settings** (own-shape weight, weather on/off + sensitivity
    + location, the **"Show calibrated curves"** diagnostic) apply to every year that uses
    Seasonality or Target. The **ⓘ button on each product** opens a per-SKU explainer
    (demand signals, blended baseline, applied curve, original→modelled result).
- **Forecast xlsx export** — the **Forecast xlsx** button (top bar) downloads each
  product's **weekly sales-unit forecast**: SKU code in column A, catalogue status
  (Live / Not Live) in column B, then one column per week (W1–W53, with the
  week-commencing date) and a Total. It reflects whatever the forecast currently is —
  multiplier, and that year's forecast mode.
- **Update selling prices (ASP)** — **Settings → Data**. Two ways to set each product's
  **Average Selling Price** (which feeds every sales-value figure and target tracking):
  - **Upload a sales file** — a *"Product Sales by Account"* export (last 12 months),
    matched by SKU code. The upload is scored **three ways at once** and shown in a
    **comparison table** for the loaded year, so you can see the impact before choosing:
    **Original** (your current prices), **Basic average** (total value ÷ total units
    across every channel — volume-weighted) and **Highest-selling customer** (the price at
    the single channel with the most sales for that product — stops a line that sells
    mostly on a higher-priced marketplace being dragged down by cheaper DSV channels).
    Each row shows the **average £/unit**, the resulting **stock-capped sales** and
    **forecast plan**, and the change vs your current prices. **Click a row to choose that
    basis**, then you see the **largest price moves** and **tick exactly which years to
    apply to** (current and later forecast years pre-ticked; historical years left alone
    unless ticked) and apply.
  - **Manual prices** — a filterable list of every product with an editable price, for new
    lines not in the sales file. Each row shows the **Previous price** and the **Var %**
    (how the current price differs from it — e.g. how much the last upload moved it), the
    source tag, and the editable price (variance updates live as you type). A "show only
    not-updated" toggle surfaces what still needs a price. You can also edit a single
    product's price straight from the supplier screen by clicking its **ASP chip**.
  - Every product shows where its price came from: **not updated** (still the original
    imported price — amber, and its row is highlighted in the supplier screen),
    **upload** (from a sales file — green) or **manual** (hand-edited — blue), so the
    ones needing attention stand out. Manual edits apply to the current year and later
    forecast years.
- **Weeks Cover colours** — the Weeks Cover row is shaded by value. The default
  treats ~4 weeks as the green "sweet spot": below that shades toward red (stock-out
  risk), above it shades through yellow/orange/red to purple (overstock). Every
  threshold and colour is editable under **Settings → Weeks Cover colour bands**
  (add/remove bands, change the numbers and colours, or reset to default).
- **Settings** — sales forecast multiplier (the old Value Splits P494 lever),
  current week, container CBM, warehouse capacities, the Weeks Cover colour bands,
  and a button to restore the order quantities originally imported from Excel.
- Changes **save automatically** (data/orders.json + data/settings.json).
  **Export CSV** downloads all order quantities.

## How the maths works (mirrors the Excel exactly)

For each SKU and week:

- **Sales Forecast** = imported base forecast × multiplier (rounded)
- **Stock Closing** = weeks before current week: actuals from Running Stock;
  current week: this week's orders + live stock; future weeks:
  max(0, previous stock + orders − forecast)
- **Sales Value** = past weeks: actual units × ASP; future: min(forecast, stock) × ASP
  (you can't sell stock you don't have)
- **Weeks Cover** = stock ÷ average forecast of the next 4 weeks
- **Order FOB / CBM / Stock Holding £** = orders × FOB, orders × unit CBM, stock × landed cost
- Containers = weekly CBM ÷ container size; warehouse spaces = ceil(stock ÷ full
  pallet qty) summed by pallet type (from the TIxHI tab)

The import was verified cell-by-cell against Excel's own calculated results:
100% match on forecast and sales value, 99.94% on stock (the only differences are
cells where a number had been typed over the formula in the workbook itself).

## Refreshing data from the workbook

When the weekly imports in the Excel file are updated (Sales Trending, Running
Stock, Buying Report, Qlik ASP):

    python import_data.py

Your edited order quantities are **kept**. To throw them away and re-import
Excel's order quantities too:

    python import_data.py --reset-orders

## Files

| File | Purpose |
|---|---|
| `import_data.py` | One-off / weekly import from the Excel workbook → `data/*.json` |
| `supplier_form.py` | Builds & parses the per-supplier factory order-planning Excel form |
| `server.py` | Local web server (Python stdlib + openpyxl for the supplier forms) |
| `static/` | The app UI (HTML/CSS/JS, calculation engine in `app.js`) |
| `data/master.json` | Reference + weekly data imported from Excel |
| `data/orders.json` | **Your** order quantities (auto-saved) |
| `data/orders_imported.json` | Order quantities as imported from Excel (backup) |
| `data/settings.json` | Multiplier, current week, capacities |

## Not carried over (yet)

- Carriage cost by carrier (hidden 'Carriage Costs' tabs)
- Lost Sales Calculator / Monthly Variance tabs
- Product image gallery (the app shows the primary PIM image per SKU)

## Restarting the server
Windows lets a second server start while an old one still owns the port, so a "restart"
can leave a STALE process answering requests (symptoms: new features visible in the page
but their endpoints failing / new data missing). If a restart doesn't seem to take
effect: close the old console window first — if in doubt, Task Manager -> end all
`python` processes, then run run.bat once (same applies to the PromoPlan app on 8766).
