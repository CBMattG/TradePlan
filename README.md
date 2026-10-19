# Tradeplan (standalone app)

A standalone replacement for the Excel Tradeplan workbooks (originally
`2026 Tradeplan - V2 - Forecast Unlinked.xlsx`); the app is now the master for every
plan year. Edit weekly **Committed Orders** quantities per SKU and instantly see the
effect on stock closing, weeks cover, sales value, FOB spend, CBM/containers, stock
holding value, supplier splits and warehouse capacity.

No installs needed beyond Python + openpyxl (already on this machine). No internet
required for planning; the optional live-weather nudge (Open-Meteo) and product images
need a connection.

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

Seed a year from its Excel workbook (2025 and 2026 only — see *Refreshing data*
below before re-running this on a year that's in use):

    python import_data.py --year 2026        # 2026 layout (default path)
    python import_data.py --year 2025         # 2025 layout
    python import_data.py --year 2025 "path\to\2025 file.xlsx"

The importer uses the 2025 sheet layout for years up to 2025 and the 2026 layout
otherwise. Only 2025 and 2026 have default workbook paths. Historical years (e.g. 2025)
load with their current week set to 53 so all weeks show actuals.

**Add a forecast year** — the **+ 2029**-style button next to the year selector always
adds the year **after the latest one held** (whichever year you're viewing). It is
built from that latest year: every product, its rolled-over weekly sales forecast, and
its projected year-end stock (committed + proposed) as the new year's starting stock.
The rolled-over sales forecast is **always rebuilt from sales patterns** — each
product's recency-weighted demand (this year's sell-through + stock-out-corrected last
year + the original planner forecast as a stabiliser) shaped by its seasonal curve — so
fast movers and lines that sold out rise and slow movers decline, rather than the prior
year being copied verbatim. (This base rebuild is independent of each year's **Forecast
mode**, which re-models a year's forecast on top of this rolled-over base.) Product
attributes (cartons, duty, pack size, cost sources…) carry over. The new year starts
with **no orders** — it's a real, browsable, editable year: run its own rebuy plan and
export supplier forms for it. **Existing years are never rebuilt or overwritten** (the
server refuses); they don't need it — a forecast year's starting stock and LY Sales
re-chain to the prior year every time it's opened, and its forecast can be re-modelled
in Settings → Forecast.

## Using the app

- **Rebuy suggestions (in the Plan grid)** — the auto-rebuy schedule is built into the
  Plan page. Every product shows a teal **"Proposed Rebuy"** row (editable) with the
  algorithm's suggested **additional** orders on top of your committed orders, and the
  Stock Closing / Weeks Cover / panel reflect committed + proposed **live** so you see
  the impact as you edit (and as you switch year). It simulates each Live product's
  stock from current free stock + the (seasonality-modelled) forecast, rolling into
  next year, to hold the target **weeks cover**, timed by **lead time** (China ~11 wk /
  non-China ~3, by origin). The **⟳ Rebuy** dropdown in the Plan toolbar toggles **Full 68 CBM only / Allow partial**
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
    demand stops being proposed). **Clear** empties the Proposed Rebuy row.
  - **Apply to: This supplier / Whole year** — a scope toggle in the **⟳ Rebuy** dropdown
    governs all three actions (Run rebuy, Clear, Commit). "This supplier" (the
    default) acts only on the supplier you're viewing, leaving every other supplier's
    proposals untouched; "Whole year" acts across all suppliers. Your choice is
    remembered. (This replaces the separate per-supplier commit button.)
  - Each year **remembers its own proposed-rebuy state** exactly as you left it.
    Suggestions are only ever built when you click **⟳ Run rebuy** — a year you've never
    run has none. If a forecast year's starting stock moves because the prior year
    changed, the toolbar shows **⚠ suggestions may be out of date**; nothing is rebuilt
    for you, so re-run when you want them refreshed from the latest forecast.
  - **Not Live lines never carry proposed rebuys** — any left over from when a line was
    Live are cleared when the year loads, when the Buying Report marks it Not Live, or
    when you set it Not Live in Details. Committed orders are never touched.
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
- **Revert orders** — **↺ Revert orders** (in the Plan toolbar's **⋯** menu) lets
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
  stock *and* its LY reference from the new 2026 position, instead of keeping a stale
  snapshot, and flags its rebuy suggestions as possibly out of date — click **⟳ Run
  rebuy** to see the knock-on effect. (The 2027 **sales forecast** itself — the plan —
  is re-modelled in **Settings → Forecast**.) In the
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
  as **committed + proposed = total** — the big figure is the **total** (committed
  orders **plus** the proposed rebuys), with a small teal **+x** beside it for the
  proposed-rebuy addition when there is one; hover a figure for the full committed +
  proposed split. So **before committing** anything
  you can see what the whole-sheet sales and FOB spend would become if the
  suggested rebuys were placed. (Forecast sales is stock-aware — adding proposed
  orders lifts sales only where extra stock lets more forecast demand be met;
  FOB/units/volume/containers are the direct order additions.) **Selling units** — the
  units behind the sales £ (actuals, then forecast capped by stock) — are shown the same
  way. Committing folds the proposed figures into the committed ones.
- **Rebuy status (Live / Not Live)** — each product carries its catalog status
  from the Buying Report: a green **Live** badge (plan for rebuy) or a grey
  **Not Live** badge (discontinued — run the remaining stock down, don't rebuy).
  Not Live rows are dimmed. The whole-plan band shows the Live/Not Live totals, and
  the **Status** dropdown in the Plan toolbar filters the grid to All / Live / Not
  Live so you can focus on what needs reordering.
  (SKUs not found in the Buying Report show an amber "Unknown" badge.)
- **Suppliers list** (left) — the £ beside each supplier is its **total forecast
  sales for the selected year**. Sort the list by that value (high→low or low→high), A–Z, or
  back to the original workbook order using the buttons at the top. Your chosen
  **sort order**, the **visible rows** (the toolbar's **Rows** dropdown), the **Status
  filter** and the **selected supplier** are all remembered — they stay put when
  you switch year and when you reopen the app. (If the selected supplier doesn't
  exist in the year you switch to, it falls back to the first supplier.)
- **Export / Import supplier form (Excel)** — the Plan toolbar's **⋯** menu downloads
  a factory-style order-planning workbook (the "Scheduling" form:
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
  The form's lead-time cells (sailing, production grace, UK inland — giving each
  container block's **Goods Ready date**) and its **Marketing Contribution / Deposit /
  Final Balance** percentages come from **Settings → Supplier forms**; in the workbook
  each is a single master cell the factory can flex.
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
  Each product collapses to a summary strip (stock now, on purchase, cover now, YTD vs
  forecast); click its image, code or name to expand the weekly grid, or use
  **⌄ Expand all** / **⌃ Collapse all**. Tick **<prior year> trend** to overlay a faint
  dashed last-year line on each supplier card.
- A pinned **Order CBM footer** runs along the bottom of the grid: the supplier's total
  order volume (committed + proposed) for **each week**, colour-coded by how close it is
  to filling a container — amber while building up, green at roughly a full container,
  blue once it's over one — so you can see at a glance which weeks are light, full or
  over (hover a week for the exact CBM, container fraction and committed/proposed split).
- Red cells = projected stock-out (zero stock with forecast demand).
  Amber weeks-cover = under 2 weeks. Green column = current week.
- **Summary** tab — the whole plan (or a filtered subset) for the year. **Filters** for
  code, name, status, supplier, category and **NPD only** narrow every figure, so a
  filtered Summary is the real plan for that subset. It shows cards (forecast sales,
  **selling units**, FOB, order units, volume, containers, peak stock £, peak
  racking/stillage); the whole-year sales-shape chart with containers arriving; the
  **sales-channel split**; weekly totals (sales £, selling units, forecast units, order
  units/FOB, CBM, containers, stock holding, pallet/stillage/racking spaces and
  utilisation); quarterly sales vs the prior-year history (hidden while filtered); and a
  **Supplier breakdown** (each supplier's share, YTD vs plan) you can rank, search,
  filter by origin group or ahead/behind plan, and cap to the top 10/20/50.
- **Arrivals** tab — upcoming container arrivals, joining the **WEBSA Open PO** balances
  (ordered − delivered) to the **Qlik Container** bookings (arrival = delivery-to-CB,
  else UK-port ETA), grouped in shaded week blocks. Outstanding POs with no dated
  booking are grouped by **estimated booking month** (due date minus an adjustable
  **booking lead time**: sailing + factory grace + UK inland), with overdue ones
  flagged. Filter by PO, product, supplier or container; click a PO for full detail.
  **Include plan orders (no PO)** adds the plan's committed-but-unraised stock and
  proposed rebuys for the full booking outlook (remembered). Needs both PO uploads.
  **Export xlsx** downloads "Upcoming Containers DD-MM-YYYY.xlsx" matching the filter:
  booked containers, Awaiting Booking, and (with plan orders on) Plan Orders + a Full
  Outlook — each week-banded with Monthly totals, a **Loading Port** column (the
  supplier's port in the plan, else its origin), current stock, 4/12-week average sales
  and weeks of cover after the delivery — plus a **Bookings by Port** sheet counting
  containers per month by loading port.
- **Warehouse Capacity** tab — pallet+racking and stillage use as a % of capacity
  through the year, with a 100% line and the current week marked. Hover the graph to
  follow a week or click to hold it; two leaderboards list the top 15 space-using
  products that week (Pallets & racking / Stillages), each with its own search (code,
  product or supplier) that also drives the graph, cards and weekly table. **NPD only**
  narrows everything to new products. Below is the weekly capacity totals table.
- **➕ Product** (top bar) — adds a brand-new product (and its supplier, if new — origin
  sets the lead time). Enter code, name, supplier, season, category, status, FOB,
  landed, ASP, current stock and an **annual forecast**, which is spread over the weeks
  by the season's curve (flat for continuity); current stock counts from the current
  week. Optional carton sizes + pack size auto-fill CBM, pallet loading and
  pallet/stillage type. Tick **NPD** to tag it as a new product. It applies to the
  current and later years, and is logged and revertable.
- **Details** (button on each product header) — every stored value for the product,
  editable in one place: identity and classification (incl. moving it to another
  supplier, status, image, NPD year), costs (FOB, landed, ASP, duty), space and packing,
  stock, and the **53-week planner forecast** with helpers (spread by season, keep shape
  and scale to annual, flat, copy last year, undo). Each value shows where it came from
  (import, manual, from cartons, original import, or sell-through plan). The **product
  code** can be renamed behind a confirmation (it must be unused in each target year;
  orders and rebuys follow the product). Only changed fields are saved; the save is
  revertable.
- **CBM chip / carton editor** — the CBM pill on each product (import, manual or "from
  cartons") opens a carton editor: enter one or more carton sizes and weights plus pack
  size, and it calculates item CBM and the units per pallet (or per stillage, if too big
  for a pallet) from the pallet/stillage dimensions in **Settings → Capacity & weeks**,
  showing the calculation next to the figure currently in use.
- **Condition tags** — live labels worked out from the plan (nothing stored):
  Overstocked, Understocked, Stock-out, Overselling, Underselling, Rebuy pending, No PO,
  Run-down, Slow mover, NPD, plus Missing cost / Missing space data (off by default).
  Each shows as a chip on the product header and as a **quick-filter chip** (with
  counts) under the search; typing the tag name in the search filters by it.
  Thresholds and on/off are in **Settings → Tags**. The **NPD** tag lasts for the plan
  year a product was added for and drops automatically the next year (or clear it in
  Details).
- **Export the current search** — when the sidebar search matches products (a tag such
  as "overstocked", a code/name, or a supplier), an **⬇ Export N … products** button
  downloads them as a workbook (supplier, category, season, status, NPD, tags, stock,
  cover, stock-out week, remaining/year forecast, YTD vs plan, committed/proposed,
  ASP/landed/FOB, margin, stock value, space data), sorted by stock value.
- **PO & container matching** — with the WEBSA Open PO and Qlik Container uploads in,
  the Plan grid shows a **PO / container** row against each week's arrivals. Weeks with
  committed stock but no matching PO or booking are listed in **Settings → Data**, with
  a **Review & cycle** walk-through. **Apply container dates to plan** (Settings → Data
  for the whole year, or the toolbar's ⋯ menu for this supplier) **re-times** committed
  arrivals to each PO's arrival week (booked container date, else PO due date) — it only
  moves units, never invents or drops them or touches past weeks — with a preview first
  and one-click undo.
- **Sales channels** — the **Channel Index** import gives each customer's unit share per
  SKU (weighted by this year's units, falling back to last year early in the year).
  Customers are assigned to Marketplace / DSV / Direct / Ex-Works in **Settings → Data →
  Sales channels** (unassigned = Unmapped). Each product's **⇄ Channels** chip shows its
  forecast units and value by channel and customer, and the Summary shows the plan-wide
  split.
- **↻ History** (top bar) — every upload, supplier-form import, Load config, Revert
  orders and data-edit save is logged with a timestamp
  (bursts of the same edit within 10 minutes merge into one entry). The menu lists the
  last 5 revertable changes; the full log is in **Settings → Changelog**. Reverting
  restores the affected data files from a before-snapshot (kept for the 5 most recent
  changes) and reloads.
- **Theme** — the moon/sun button in the top bar switches light/dark mode (remembered).
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
  - A **forecast pill** beside the supplier name on the Plan page shows the year's active
    mode (e.g. *🌦️ Seasonality 100%*, *🎯 Target £15m*, *📈 Original forecast*); click it
    for the detail and the **adjust** / **revert** / **turn off** links. **Shared model settings** (own-shape weight, weather on/off + sensitivity
    + location, the **"Show calibrated curves"** diagnostic) apply to every year that uses
    Seasonality or Target. The **ⓘ button on each product** opens a per-SKU explainer
    (demand signals, blended baseline, applied curve, original→modelled result).
  - A product whose forecast is **pinned** (a stored `fc_pinned` flag, e.g. a
    sell-through plan for a line with no sales history — shown as "sell-through plan
    (pinned)" in Details) keeps its weekly plan as set: Seasonality and Target modes
    don't re-model it, and Target mode counts it as a fixed part of the target.
- **Forecast xlsx export** — the **Forecast xlsx** button (top bar) opens a choice of
  two versions, both with SKU in column A, Live/Not Live in B, W1–W53 (with
  week-commencing dates) and a Total, reflecting the multiplier and forecast mode:
  **Full demand forecast** (every forecast unit — the basis for buying) or
  **Achievable only, capped by stock** (the units behind the Sales Value row: banked
  actuals for traded weeks, then forecast capped at projected stock). **Not Live** lines
  are never rebought, so in both versions their current and future weeks are capped at
  the stock they hold. The dialog previews how many units each cap removes; your choice
  is remembered.
- **Stock xlsx export** — the **Stock xlsx** button (top bar) downloads each product's
  weekly closing stock (SKU, status, W1–W53): recorded history up to the current week,
  the plan's projection after, plus two columns counting **weeks at zero** (across the
  year, and across the last four completed weeks).
- **File Imports** (**Settings → File Imports**) — every Excel upload is a tile,
  coloured by freshness for its group's cadence (green current, amber one period old,
  red older/never); click to upload, or drag tiles to reorder or move them between
  groups. Each upload previews first, is logged in History and can be reverted.
  - **Weekly Sales** — the "WKnn Sales" export for one week: sets that week's actual
    units, closes the week's stock (chained forward, past-week gaps healed), moves the
    actuals/forecast boundary on a week, and stamps a **weekly ASP** chip beside the ASP
    pill.
  - **Buying Report** — catalogue status (Live/Not Live), outstanding purchases and
    live free stock for the ticked years. The stock reading is also **recorded against
    the week it was taken in**, and past closing stock re-anchors on each reading, so
    receipts never keyed as orders aren't lost. Lines turning Not Live lose their
    proposed rebuys.
  - **Sales / ASP**, **Landed Costs** (FOB + landed), **Duty Rates** — matched by code
    (see ASP below). The FOB and Landed pills are coloured by source and click-to-edit
    like the ASP chip; an **Est. landed** chip estimates future landed cost (FOB + sea
    freight per item → £ → + duty → + inland haulage) using the assumptions in
    **Settings → Capacity & weeks**.
  - **WEBSA Open PO**, **Qlik Container** — feed the Arrivals tab and PO matching.
  - **Channel Index** — the "Channel Sales" export (see Sales channels).
  - **Carton Sizes** — reads the workbook's **TIxHI** sheet to fill carton dimensions
    and pack size (leaving CBM, loading qty and pallet type alone until a product is
    reviewed in the carton editor).
- **Update selling prices (ASP)** — two ways to set each product's
  **Average Selling Price** (which feeds every sales-value figure and target tracking):
  - **Upload a sales file** (the **Sales / ASP** tile in **Settings → File Imports**) —
    a *"Product Sales by Account"* export (last 12 months),
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
  - **Manual prices** (**Settings → Data**) — a filterable list of every product with an editable price, for new
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
  threshold and colour is editable under **Settings → Weeks Cover**
  (add/remove bands, change the numbers and colours, or reset to default).
- **Settings** — tabbed:
  - **Forecast** — sales forecast multiplier (the old Value Splits P494 lever), the
    per-year forecast mode and shared model settings.
  - **Capacity & weeks** — current week, container CBM, warehouse capacities, pallet
    and stillage dimensions, landed-cost estimate assumptions.
  - **Weeks Cover** — the colour bands. **Tags** — condition tags and thresholds.
    **Rebuy** — rebuy parameters. **Supplier forms** — lead-time and payment terms in
    the exported forms.
  - **File Imports** — every upload tile. **Data** — sales-channel mapping, manual
    prices, manual CBM (hand-set item CBM; edited products are tagged manual), PO /
    container tools, and the button to restore the order quantities originally
    imported from Excel. **Changelog** — the full change history with revert.
- Changes **save automatically** (`data/<year>/orders.json` + `proposed.json`, and the
  global `data/settings.json`), and each save is logged in History.
  **Export CSV** downloads all order quantities.

## How the maths works (mirrors the Excel exactly)

For each SKU and week:

- **Sales Forecast** = (base forecast, or for future weeks the Seasonality/Target-modelled
  forecast) × multiplier (rounded)
- **Stock Closing** = weeks before current week: the recorded closing stock (chained
  from the weekly sales imports, re-anchored on each Buying Report stock reading);
  current week: this week's orders + live stock; future weeks:
  max(0, previous stock + orders − forecast)
- **Sales Value** = past weeks: actual units × ASP; future: min(forecast, stock) × ASP
  (you can't sell stock you don't have). **Selling units** are the same units without
  the × ASP.
- **Weeks Cover** = stock ÷ average forecast of the next 4 weeks
- **Order FOB / CBM / Stock Holding £** = orders × FOB, orders × unit CBM, stock × landed cost
- Containers = weekly CBM ÷ container size; warehouse spaces = ceil(stock ÷ full
  pallet qty) summed by pallet type (from the TIxHI tab)

The import was verified cell-by-cell against Excel's own calculated results:
100% match on forecast and sales value, 99.94% on stock (the only differences are
cells where a number had been typed over the formula in the workbook itself).

## Refreshing data

Weekly data is uploaded **inside the app** from **Settings → File Imports** (Weekly
Sales, Buying Report, Sales/ASP, Landed Costs, Duty Rates, Channel Index, Carton Sizes,
WEBSA Open PO, Qlik Container). Each upload previews first, is logged, and can be
reverted.

`python import_data.py [--year 2025|2026] [path]` re-imports a year **from scratch**
from its Excel workbook. It **replaces that year's `master.json`**, losing every price,
cost, status, stock reading, weekly actual, carton/duty update, NPD tag and product
added in the app since. Your order quantities in `orders.json` are kept unless you pass
`--reset-orders`. It asks you to type the year to confirm before replacing an existing
year (`--yes` skips this) and first copies each file it replaces to
`*.pre-import-<timestamp>`. It refuses a year with no workbook (generated forecast
years, 2027+) unless you give a path. Only use it to (re)seed a year.

## Files

| File | Purpose |
|---|---|
| `run.bat` | Starts the server |
| `server.py` | Local web server (Python stdlib + openpyxl): all `/api/*` routes, upload parsing, change log |
| `supplier_form.py` | Builds every Excel export (supplier Scheduling form, forecast, stock, arrivals, search) and parses the supplier form on import |
| `import_data.py` | Seeds / re-imports one year from its Excel workbook → `data/<year>/` (overwrites `master.json`) |
| `static/` | The app UI (HTML/CSS/JS, calculation engine in `app.js`) |
| `data/years.json` | Index of plan years + default |
| `data/<year>/master.json` | Reference + weekly data for the year (products, suppliers, costs, actuals, stock) |
| `data/<year>/orders.json` | **Your** committed order quantities (auto-saved) |
| `data/<year>/proposed.json` | The year's Proposed Rebuy layer |
| `data/<year>/orders_imported.json` | Order quantities as first imported / built (for Revert → Original) |
| `data/<year>/stockbase.json` | Forecast years only: the prior-year-end stock last chained from |
| `data/settings.json` | Global settings (multiplier, current week, capacities, rebuy, tags, channel map…) |
| `data/po_websa.json`, `data/po_containers.json` | WEBSA Open PO + Qlik container uploads (shared by all years) |
| `data/channel_index.json` | Channel Sales customer-share upload |
| `data/changelog.json`, `data/change_snapshots/` | Change history + before-snapshots for the 5 most recent revertable changes |
| `data/configs/` | Named **Save config** checkpoints |

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
