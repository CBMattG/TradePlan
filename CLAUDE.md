# Tradeplan app — Claude context

Standalone local buying/forecast app replacing the Charles Bentley "Tradeplan" Excel
workbook. A buyer edits weekly committed-order quantities per SKU and sees the live
effect on stock cover, sales value, FOB spend, CBM/containers, capacity, etc.
**In active daily use — treat the data under `data/` as production.**

## Stack (no build step; only external dep is openpyxl)
- `server.py` — Python stdlib `ThreadingHTTPServer` on **:8765**; all `/api/*` routes; upload parsing via openpyxl; change log + revert; stock re-chaining.
- `supplier_form.py` — every xlsx export builder (supplier Scheduling form, forecast, stock, arrivals, search) + the supplier-form import parser. Imported by `server.py`.
- `import_data.py` — seeds a year (2025/2026 layouts) from its Excel workbook. **Destructive**: overwrites that year's `master.json`.
- `static/app.js` — one large (~7.3k-line) vanilla-JS calc engine (mirrors the Excel formulas) + all UI. No framework/bundler.
- `static/index.html`, `static/style.css`.
- `data/<year>/` — per-year `master.json`, `orders.json`, `proposed.json`, `orders_imported.json`; `stockbase.json` on forecast years only.
  Global: `data/years.json`, `settings.json`, `po_websa.json`, `po_containers.json`, `channel_index.json`, `changelog.json` + `change_snapshots/`, `configs/` (Save config).

## Run
    python server.py                     # or run.bat → http://localhost:8765
    python import_data.py --year 2026    # DESTRUCTIVE: replaces data/2026/master.json (all app-side imports lost; orders.json kept). Prompts + backs up to *.pre-import-<ts>; refuses years with no workbook (2027+) unless given a path. Never run it to "test"

## Editing conventions (important)
- **Bump the cache-bust `?v=N`** on `app.js`/`style.css` in `index.html` for every edit, or the browser runs stale code.
- **Restart `server.py`** for server changes; a browser reload suffices for JS/CSS/HTML.
- **No `node`** on this machine: syntax-check the server with `python -m py_compile server.py`; validate JS by bracket-balancing the changed block + replaying parser logic in a throwaway Python script against the real source files.
- Read source `.xlsx` with **Python/openpyxl** (PowerShell file-open trips AMSI antivirus). Wrap reads in `warnings.catch_warnings()` — some sheets carry junk date serials that spam warnings.

## Data safety (the user edits live between turns)
- **Never force-correct totals.** The per-SKU committed + proposed total is the invariant; the baseline shifts as the user works.
- Before any test that writes to `data/`, **back up** the file and restore/verify after (or snapshot/restore in memory).
- Cost/status imports (`apply-asp|landed|buying|duty`) write additive `master.json` fields (`asp`, `fob`, `landed`, `status`, `os_purchases`, `duty_rate`), matched by product code; `asp`/`fob`/`landed` are tagged with a `*_src` (buying/duty aren't).
- `apply-buying` also writes `stock_now` / `stock_snap[week]`, re-chains `running_stock`, and **deletes Not Live entries from `proposed.json`** (`drop_not_live_proposals`; `apply-sku` does too when a line is set Not Live). Other `master.json` writers: `apply-wksales`, `apply-cbm`, `apply-cartons`, `apply-tixhi`, `apply-sku`, `add-product`, `set-removed`, `revert-change`. Each `apply-*` is logged via `record_change` with a before-snapshot.
- **Graveyard:** a product flagged `removed_at` (+ `removed_status` = its prior status; set Not Live) is hidden from the Plan, sidebar, tags, search and supplier forms but stays in every total/calc (`isRemoved()` client-side). Set by `apply-buying` (`remove` = codes the user ticked in the preview; anything back on the report auto-restores) or `set-removed` (manual, Graveyard tab / product dialog).
- **Server writes:** every data-file write goes through `write_atomic(path, text)` (temp file + `os.replace`) — never `Path.write_text` on a data file. Writing POST routes run under `WRITE_LOCK` (`/api/export-*` skip it). Any new route that changes data must `record_change(kind, label, detail, [every file it touches])` **before** writing, so History can undo it (a file absent at snapshot time is deleted on revert).
- Offline server tests: import `server`, point `S.DATA` / `S.CHANGELOG_PATH` / `S.SNAP_DIR` at a scratch **copy** of `data/`, and call `Handler.method(fake_self, qs)` with a fake exposing `send_json`, `headers`, `rfile`.

## Working style
- **Don't start the preview / restart the user's server to verify unless asked** — they run their own server on :8765 and restarting disrupts their session. Rely on code review + offline checks by default.

## Version control (git)
- Repo initialised; **code + docs are tracked, `data/` is gitignored** (see `.gitignore`). Git must never touch live plan data — branch switches/checkouts only move code, so the user's running server + `data/` are always safe.
- **Solo workflow:** small/low-risk changes (a settings option, label, colour, single-feature edit) are committed **directly on `main`** — no branch needed. Use a **branch** only for risky or long-running work where the user needs the stable version runnable meanwhile; merge to `main` and delete the branch when done. **Do not create per-feature "forks"/folder copies** — they fragment the live data and invite the data-loss juggling this project is prone to.
- **Data backup is NOT git** — it's the app's "Save config" (`data/configs/*.json`) + the portable zip. Keep it that way.
- Commit in small, working increments with a clear message. Only commit/branch when the user asks, or when finishing a discrete change they've approved.

## Patterns (copy an existing one)
- **Data imports** (ASP/landed/buying/duty/wksales share one shape; WEBSA/Qlik/Channel Index parse **and persist** in one call to global JSON, stamped by the file's `importedAt`): server `aggregate_X`/`parse_X` (`/api/parse-X`) + `apply_X` (`/api/apply-X`, writes `master.json` per year by code); client `XFileChosen`→`openXDialog` (preview)→`applyXUpdates`→`applyXToMemory`; state stamped in `SETTINGS.X_updated_at`. The upload button is a **tile in the "File Imports" settings tab** — register it in `IMPORT_DEFS` (`id`/`label`/`input`/`when()`); `renderUploadAges` renders all tiles into `#imports-groups`, coloured by group cadence.
- **Editable per-product chips**: ASP/FOB/Landed follow `aspChip`/`fobChip`/`landedChip` + `editAspInline`/`editCostInline` — a pill coloured by a `*_src` tag that click-edits via `prompt()` → `apply-*`. CBM (`cbmChip`) opens the carton dialog instead.
- Client feature state persists in `SETTINGS.*` via `markDirty()` (saves `settings.json` + the year's `orders.json`/`proposed.json`, logged as an "edit" change); per-year product attributes live in `master.json`.

## Gotchas
- Source `.xlsx` live on the user's redirect: `\\cb-fs1\folderredirection$\matthew.gwatkin\{Desktop,Downloads}\…`. To read/verify one, **write a `.py` script file** — inline `python -c` mangles the `\\` backslashes. Verify a parser by replaying it in Python against the real file (match rate vs `master.json` codes; the landed calc ties to the Excel to the penny).
- CSS: later blocks win on equal specificity — "APPLE REFRESH" (~2/3 down `style.css`) overrides the base rules, and the feature blocks after it (Settings/dialog refresh, arrivals, **PLAN REFRESH**, channels, changelog) override APPLE REFRESH; dark overrides may need id-scoping to beat earlier `[data-theme="dark"]` rules; give dialog `<fieldset>` `min-width:0` to stop horizontal overflow.
- **Target-mode quirk (known, deliberately left as is):** `buildTargetForecasts` hits the target as Σ **actuals** for traded weeks (`w+1 < current_week`) + Σ modelled **future** forecast, but `forecastDemandTotal()` (the "modelled demand" shown in the forecast pill and Settings → Forecast) sums `r.forecast × asp` over **all 53 weeks**, i.e. the *forecast* for traded weeks, not their actuals. The two drift apart by Σ traded-week (forecast − actual) × ASP (e.g. £15.00m target vs ~£17.15m demand). Don't "fix" either number silently — if asked, add a like-for-like actuals + future-forecast figure.
- **User's screen is a 1920×1200 laptop at Windows display scaling ≈ 1536 CSS px wide** (not 1920). Check layout changes at ~1536 px: the Plan toolbar, top bar and totals band all have `@media (max-width: 1800px)` / `(max-width: 1400px)` tightening. CSS `zoom` emulation doesn't trigger media queries or change `vw` — inject the media rules when emulating.
- Toolbar dropdowns are `<details class="dd">`: one global handler closes them (outside click / Escape / opening another / a `.dd-menu` action) and `placeDD()` flips a popover to the side with room (`#main` is `overflow:hidden`, so an off-edge popover gets clipped). To keep one open across a `renderPlan()`, set `reopenDD = '<id>'`. Clicks on elements a re-render replaced are ignored (`isConnected`). Popovers use a **solid** background — a `backdrop-filter` nested inside the blurred toolbar can't blur, so glass colours turn see-through.
- Visual/layout review without touching the user's server: run a scratch `sandbox_server.py` that imports `server`, points `S.DATA`/`S.CHANGELOG_PATH`/`S.SNAP_DIR` at a **copy** of `data/` and serves on :8799, then drive it with the Chrome tools.
- `preview_eval`: module-level `let` globals (`M`, `SETTINGS`, `ORDERS`, `PROPOSED`, `PO_WEBSA`, `PO_CONTAINERS`, `CHANNEL_INDEX`, `YEAR`, `YEARS`…) are barewords, not `window.*`; screenshots time out on the heavy grid — verify with `getComputedStyle`.
