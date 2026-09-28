"""Web app: ticker -> models shown in the browser; Excel available via a download button.
Single source of truth: the workbook is built (formulas), recalculated by headless LibreOffice,
then its VALUES are read back and sent to the page. The same recalculated file is the download."""
import os, re, shutil, tempfile, time, uuid

from flask import Flask, jsonify, request, send_file, send_from_directory
from openpyxl import load_workbook

from data_source import demo_company, fetch_company
from workbook import build_workbook

app = Flask(__name__, static_folder="static")
WORK = tempfile.mkdtemp(prefix="t2m_")
TTL = 3600
CACHE = {}                       # id -> {path, payload, ts, key}
TICKER_RE = re.compile(r"^[A-Z0-9.\-^=&]{1,20}$")


def calculate(src):
    """Pure-Python evaluation of every formula (no LibreOffice/Excel needed).
    Returns an openpyxl workbook whose formula cells are replaced by computed values."""
    import formulas
    sol = formulas.ExcelModel().loads(src).finish().calculate()
    name = os.path.basename(src)
    wb = load_workbook(src)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    k = f"'[{name}]{ws.title.upper()}'!{c.coordinate}"
                    v = sol[k].value[0, 0] if k in sol else None
                    if hasattr(v, "item"):
                        v = v.item()
                    if v is not None and not isinstance(v, (int, float, str, bool)):
                        v = None if str(v).lower() == "empty" else str(v)
                    c.value = v
    return wb


def show(v, nf):
    """Format a value the way Excel would display it (covers the formats used in workbook.py)."""
    if v is None:
        return ""
    if isinstance(v, bool) or not isinstance(v, (int, float)):
        return str(v)
    sec = (nf or "General").split(";")[0]
    dec = re.search(r"\.(0+)", sec)
    d = len(dec.group(1)) if dec else 0
    if v == 0 and '"-"' in (nf or ""):
        return "-"
    if "%" in sec:
        s = f"{abs(v) * 100:,.{d}f}%"
    elif '"x"' in sec:
        s = f"{abs(v):,.{d}f}x"
    elif nf in (None, "General"):
        return f"{v:,.4g}" if abs(v) < 1000 else f"{v:,.0f}"
    else:
        s = f"{abs(v):,.{d}f}"
    if v < 0:
        return f"({s})" if ";(" in (nf or "") else "-" + s
    return s


def grid(ws, source_ws=None):
    rows, errors = [], 0
    for row in ws.iter_rows():
        out = []
        for c in row:
            cell = {}
            t = show(c.value, c.number_format)
            if t.startswith("#") or t.startswith("Err:"):
                errors += 1
            if t:
                cell["t"] = t
            if isinstance(c.value, (int, float)) and not isinstance(c.value, bool):
                cell["n"] = 1
            if c.font and c.font.b:
                cell["b"] = 1
            rgb = getattr(c.font.color, "rgb", None) if c.font and c.font.color else None
            if isinstance(rgb, str) and rgb[-6:] not in ("000000",):
                cell["c"] = "#" + rgb[-6:]
            if c.fill and c.fill.fill_type == "solid" and isinstance(c.fill.start_color.rgb, str):
                cell["f"] = "#" + c.fill.start_color.rgb[-6:]
            # Keep the original formula for the browser's audit panel.  `calculate`
            # intentionally replaces formulas with values, so it has to be read from
            # the source workbook instead.
            if source_ws:
                original = source_ws[c.coordinate].value
                if isinstance(original, str) and original.startswith("="):
                    cell["x"] = original
            out.append(cell)
        rows.append(out)
    while rows and not any(rows[-1]):
        rows.pop()
    return rows, errors


def editable_inputs(wb):
    """Return user-facing inputs: yellow cells are deliberately the model levers."""
    out = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                fill = c.fill
                rgb = getattr(fill.start_color, "rgb", None) if fill and fill.fill_type == "solid" else None
                if not (isinstance(rgb, str) and rgb[-6:] == "FFFF00"):
                    continue
                if not isinstance(c.value, (int, float)) or isinstance(c.value, bool):
                    continue
                label = ws.cell(c.row, max(1, c.column - 1)).value
                note = ws.cell(c.row, c.column + 1).value if c.column < ws.max_column else None
                out.append(dict(id=f"{ws.title}!{c.coordinate}", sheet=ws.title,
                                cell=c.coordinate, label=str(label or c.coordinate),
                                value=c.value, format=c.number_format or "General",
                                note=str(note or "")))
    return out


def apply_inputs(wb, inputs):
    """Accept only numeric edits to existing yellow cells, never arbitrary cell writes."""
    accepted = {x["id"] for x in editable_inputs(wb)}
    for key, value in (inputs or {}).items():
        if key not in accepted or not isinstance(value, (int, float)) or isinstance(value, bool):
            continue
        sheet, cell = key.split("!", 1)
        wb[sheet][cell].value = value


METHODS = {
    "Assumptions": "These are the model's editable levers. Yellow cells are defaults chosen from market data or recent history; changing one and applying it rebuilds every dependent model.",
    "Forecast": "Revenue growth begins with the historical revenue CAGR, constrained to a sensible range, then fades toward terminal growth. Margins, capex and working-capital ratios start as historical averages. Scenario adjustments are applied to growth and EBITDA margin.",
    "ThreeStatement": "The income statement is driven by forecast revenue and margins. Balance-sheet operating items are forecast as a percentage of revenue, and cash flow links net income, non-cash D&A, working-capital changes, capex, dividends and debt changes. The balance-check row should be zero.",
    "DCF": "DCF uses FCFF: EBIT × (1 − tax) + D&A − capex − increase in net working capital. Each year is discounted at WACC. Terminal value uses the Gordon-growth formula: final-year FCFF × (1 + g) ÷ (WACC − g), then debt is subtracted and cash added to reach equity value.",
    "Comps": "Comparable-company valuation calculates EV as market cap + debt − cash, then applies peer-group median EV/Revenue, EV/EBITDA and P/E multiples to the target's corresponding metric. Peers and their inputs should be reviewed before relying on the result.",
    "Options": "Black-Scholes-Merton estimates European call and put values from spot, strike, time, risk-free rate, volatility and dividend yield. Historical volatility is only a starting point; implied volatility is normally better for matching traded options.",
    "Credit": "Credit combines Altman Z-score ratios, leverage and interest coverage with a simplified Merton distance-to-default calculation. It is a risk screen, not a credit rating.",
    "LBO": "The LBO estimates sponsor returns from an entry valuation, debt sizing, cash-flow debt paydown and exit multiple. It reports the maximum entry price that still reaches the selected target IRR.",
    "Summary": "Summary links the headline outputs from the valuation tabs. It is not a separate calculation.",
    "Notes": "Notes document data sources, methods and model limitations.",
}


def prune():
    for k in [k for k, v in CACHE.items() if time.time() - v["ts"] > TTL]:
        shutil.rmtree(os.path.dirname(os.path.dirname(CACHE[k]["path"])), ignore_errors=True)
        CACHE.pop(k)


@app.get("/")
def index():
    return send_from_directory("static", "index.html")


@app.post("/api/analyze")
def analyze():
    j = request.get_json(force=True, silent=True) or {}
    ticker = (j.get("ticker") or "").strip().upper()
    peers = [p.strip().upper() for p in (j.get("peers") or "").split(",") if p.strip()]
    if not TICKER_RE.match(ticker) or not all(TICKER_RE.match(p) for p in peers):
        return jsonify(error="Enter a valid ticker, e.g. AAPL or RELIANCE.NS"), 400
    prune()
    inputs = j.get("inputs") or {}
    key = f"{ticker}|{','.join(peers)}|{sorted(inputs.items())}"
    for rid, v in CACHE.items():
        if v["key"] == key:
            return jsonify(v["payload"])
    try:
        d = demo_company() if ticker == "DEMO" else fetch_company(ticker, peers or None)
    except Exception as e:
        return jsonify(error=str(e)), 400

    rid = uuid.uuid4().hex[:12]
    base = os.path.join(WORK, rid)
    os.makedirs(f"{base}/in"); os.makedirs(f"{base}/out")
    src = f"{base}/in/model.xlsx"
    build_workbook(d, src)
    # Apply only deliberately exposed (yellow) assumptions before recalculation.
    source_wb = load_workbook(src)
    apply_inputs(source_wb, inputs)
    source_wb.save(src)
    try:
        wb = calculate(src)
    except Exception as e:
        return jsonify(error=f"Calculation failed: {e}"), 500

    sheets, errors = [], 0
    for ws in wb.worksheets:
        rows, e = grid(ws, source_wb[ws.title])
        errors += e
        sheets.append(dict(name=ws.title, rows=rows, wA=ws.column_dimensions["A"].width or 30))
    s = wb["Summary"]
    headline = [dict(label=s[f"A{r}"].value, value=s[f"B{r}"].value,
                     up=s[f"C{r}"].value if isinstance(s[f"C{r}"].value, (int, float)) else None)
                for r in range(5, 12) if isinstance(s[f"B{r}"].value, (int, float))]
    payload = dict(id=rid, ticker=d["ticker"], name=d["name"], currency=d["currency"], price=d["price"],
                   warnings=d["warnings"], headline=headline, sheets=sheets, formula_errors=errors)
    payload["inputs"] = editable_inputs(source_wb)
    payload["methods"] = METHODS
    CACHE[rid] = dict(path=src, payload=payload, ts=time.time(), key=key)
    return jsonify(payload)


@app.get("/download/<rid>")
def download(rid):
    v = CACHE.get(rid)
    if not v:
        return "Expired: run the analysis again.", 404
    return send_file(v["path"], as_attachment=True, download_name=f"{v['payload']['ticker'].replace('.', '_')}_models.xlsx")


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=int(os.environ.get("PORT", 8000)), debug=False)
