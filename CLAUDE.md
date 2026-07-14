# Tradeplan app — Claude context

Standalone local buying/forecast app replacing the Charles Bentley "Tradeplan" Excel
workbook. A buyer edits weekly committed-order quantities per SKU and sees the live
effect on stock cover, sales value, FOB spend, CBM/containers, capacity, etc.
**In active daily use — treat the data under `data/` as production.**

## Stack (no build step; only external dep is openpyxl)
- `server.py` — Python stdlib `ThreadingHTTPServer` on **:8765**; all `/api/*` routes; xlsx parsing via openpyxl.
- `static/app.js` — one large (~4k-line) vanilla-JS calc engine (mirrors the Excel formulas) + all UI. No framework/bundler.
- `static/index.html`, `static/style.css`.
- `data/<year>/` — per-year `master.json`, `orders.json`, `proposed.json`, `stockbase.json`.
  Global: `data/settings.json`, `data/po_websa.json`, `data/po_containers.json`.

## Run
    python server.py                     # or run.bat → http://localhost:8765
    python import_data.py --year 2026    # (re)build a year from its Excel workbook

## Editing conventions (important)
- **Bump the cache-bust `?v=N`** on `app.js`/`style.css` in `index.html` for every edit, or the browser runs stale code.
- **Restart `server.py`** for server changes; a browser reload suffices for JS/CSS/HTML.
- **No `node`** on this machine: syntax-check the server with `python -m py_compile server.py`; validate JS by bracket-balancing the changed block + replaying parser logic in a throwaway Python script against the real source files.
- Read source `.xlsx` with **Python/openpyxl** (PowerShell file-open trips AMSI antivirus). Wrap reads in `warnings.catch_warnings()` — some sheets carry junk date serials that spam warnings.

## Data safety (the user edits live between turns)
- **Never force-correct totals.** The per-SKU committed + proposed total is the invariant; the baseline shifts as the user works.
- Before any test that writes to `data/`, **back up** the file and restore/verify after (or snapshot/restore in memory).
- Cost/status imports (`apply-asp|landed|buying|duty`) write additive `master.json` fields (`asp`, `fob`, `landed`, `status`, `os_purchases`, `duty_rate`), matched by product code, tagged with a `*_src`.

## Working style
- **Don't start the preview / restart the user's server to verify unless asked** — they run their own server on :8765 and restarting disrupts their session. Rely on code review + offline checks by default.

## Patterns (copy an existing one)
- **Data imports** (ASP/landed/buying/duty share one shape): server `aggregate_X`/`parse_X` (`/api/parse-X`) + `apply_X` (`/api/apply-X`, writes `master.json` per year by code); client `XFileChosen`→`openXDialog` (preview)→`applyXUpdates`→`applyXToMemory`; state stamped in `SETTINGS.X_updated_at`. The upload button is a **tile in the "File Imports" settings tab** — register it in `IMPORT_DEFS` (`id`/`label`/`input`/`when()`); `renderUploadAges` renders all tiles into `#imports-groups`, coloured by group cadence.
- **Editable per-product chips** (ASP/FOB/Landed) follow `aspChip`+`editAspInline`: a pill coloured by a `*_src` tag that click-edits via `prompt()` → `apply-*`.
- Client feature state persists in `SETTINGS.*` via `markDirty()` (writes `settings.json`); per-year product attributes live in `master.json`.

## Gotchas
- Source `.xlsx` live on the user's redirect: `\\cb-fs1\folderredirection$\matthew.gwatkin\{Desktop,Downloads}\…`. To read/verify one, **write a `.py` script file** — inline `python -c` mangles the `\\` backslashes. Verify a parser by replaying it in Python against the real file (match rate vs `master.json` codes; the landed calc ties to the Excel to the penny).
- CSS: the "APPLE REFRESH" block at the **end of `style.css`** wins on equal specificity; dark overrides may need id-scoping to beat earlier `[data-theme="dark"]` rules; give dialog `<fieldset>` `min-width:0` to stop horizontal overflow.
- `preview_eval`: module-level `let` globals (`M`, `SETTINGS`, `ORDERS`, `PO_WEBSA`, `YEAR`…) are barewords, not `window.*`; screenshots time out on the heavy grid — verify with `getComputedStyle`.
