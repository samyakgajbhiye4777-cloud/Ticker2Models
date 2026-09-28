"""Builds the Excel workbook. Python only writes INPUTS (blue) and FORMULAS (black/green);
every number you see in the model is computed by Excel, so users can audit and change anything."""
from __future__ import annotations

import re

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import column_index_from_string as CI
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation

NP = 5  # projection years
FONT = "Arial"
F_NUM = '#,##0;(#,##0);"-"'
F_PS = '#,##0.00;(#,##0.00);"-"'
F_PCT = '0.0%;(0.0%);"-"'
F_X = '0.0"x"'
F_D = '0.000'
NAVY, GREY, YELLOW = "1F3864", "D9E1F2", "FFFF00"
LINK_RE = re.compile(r"^=[A-Za-z]+!\$?[A-Z]+\$?\d+$")

# ---- fixed row maps (so sheets can reference each other before being written) ----
TS = dict(rev=6, ebitda=7, da=8, ebit=9, interest=10, other=11, ebt=12, tax=13, ni=14,
          cash=17, ar=18, inv=19, ppe=20, oth_a=21, ta=22, ap=23, debt=24, oth_l=25, tl=26, eq=27, tle=28, check=29,
          cf_ni=32, cf_da=33, cf_ar=34, cf_inv=35, cf_ap=36, cfo=37, capex=38, div=39, cf_debt=40, netchg=41)
FC = dict(scen=5, scen_name=6, bull_g=7, bear_g=8, bull_m=9, bear_m=10, adj_g=11, adj_m=12, hdr=14,
          g_base=15, g_act=16, m_base=17, m_act=18, da=19, capex=20, ar=21, inv=22, ap=23, tax=24,
          int=25, payout=26, dchg=27)
DC = dict(hdr=4, rev=5, ebitda=6, ebit=7, taxes=8, nopat=9, da=10, capex=11, nwc=12, dnwc=13, ufcf=14, n=15,
          df=16, pv=17, val=19, wacc=20, g=21, sumpv=22, tvcf=23, tv=24, pvtv=25, ev=26, debt=27, cash=28,
          eq=29, shares=30, ps=31, price=32, up=33, tvpct=34, impx=35, chk=36, x_sec=38, xm=39, xtv=40,
          xpv=41, xev=42, xeq=43, xps=44, s_sec=46, wstep=47, gstep=48, s_hdr=49, s0=50)


# ----------------------------------------------------------------------------- style helpers
def put(ws, ref, v, fmt=None, kind=None, bold=False, fill=None, italic=False, align=None, color=None):
    c = ws[ref]
    c.value = v
    if kind is None:
        if isinstance(v, str) and v.startswith("="):
            kind = "link" if LINK_RE.match(v) else "calc"
        elif isinstance(v, (int, float)) and not isinstance(v, bool):
            kind = "input"
        else:
            kind = "text"
    col = color or {"input": "0000FF", "link": "008000", "calc": "000000", "text": "000000"}[kind]
    c.font = Font(name=FONT, size=10, bold=bold, italic=italic, color=col)
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = PatternFill("solid", start_color=fill)
    if align:
        c.alignment = Alignment(horizontal=align, wrap_text=(align == "wrap"))
    return c


def title(ws, text, sub=None):
    put(ws, "A1", text, bold=True, kind="text").font = Font(name=FONT, size=14, bold=True, color=NAVY)
    if sub:
        put(ws, "A2", sub, italic=True, kind="text", color="595959")


def section(ws, row, text, ncols=9):
    for i in range(1, ncols + 1):
        ws.cell(row=row, column=i).fill = PatternFill("solid", start_color=GREY)
    put(ws, f"A{row}", text, bold=True, kind="text", fill=GREY)


def header(ws, row, labels, first_col=1):
    for i, lab in enumerate(labels):
        c = put(ws, f"{CL(first_col + i)}{row}", lab, kind="text", bold=True, fill=NAVY, color="FFFFFF", align="center")


def widths(ws, a=40, rest=13, n=12):
    ws.column_dimensions["A"].width = a
    for i in range(2, n + 1):
        ws.column_dimensions[CL(i)].width = rest


class Lay:
    def __init__(self, d):
        self.years = d["years"]
        nh = len(self.years)
        self.HC = [CL(2 + i) for i in range(nh)]
        self.PC = [CL(2 + nh + i) for i in range(NP)]
        self.LH, self.F, self.LP = self.HC[-1], self.PC[0], self.PC[-1]
        self.labels = [f"FY{y}A" for y in self.years] + [f"FY{self.years[-1] + i + 1}E" for i in range(NP)]
        self.cols = self.HC + self.PC

    def prev(self, c):
        return CL(CI(c) - 1)


def year_header(ws, lay, row, label="($ mm, reporting currency)"):
    header(ws, row, [label] + lay.labels)


# ----------------------------------------------------------------------------- defaults from history
def defaults(d):
    h, n = d["hist"], len(d["years"])
    inr = d["currency"] == "INR"
    rev = h["rev"]
    g_term = 0.05 if inr else 0.025
    cagr = (rev[-1] / rev[0]) ** (1 / (n - 1)) - 1 if rev[0] > 0 and rev[-1] > 0 else 0.05
    cagr = max(-0.05, min(0.25, cagr))
    avg = lambda xs: sum(xs) / len(xs) if xs else 0.0
    ratio = lambda a, b=rev: [x / y if y else 0.0 for x, y in zip(a, b)]
    clip = lambda x, lo, hi: max(lo, min(hi, x))

    tax_r = [t / (ni + t) for t, ni in zip(h["tax"], h["ni"]) if (ni + t) > 0]
    int_r = [i / dbt for i, dbt in zip(h["interest"], h["debt"]) if dbt > 0]
    pay_r = [dv / ni for dv, ni in zip(h["div"], h["ni"]) if ni > 0]
    ev_now = d["price"] * d["shares"] + h["debt"][-1] - h["cash"][-1]
    ev_x = ev_now / h["ebitda"][-1] if h["ebitda"][-1] > 0 else 12.0
    return dict(
        g_term=g_term, rf=0.065 if inr else 0.043, erp=0.07 if inr else 0.05,
        growth=[cagr + (g_term - cagr) * i / (NP - 1) for i in range(NP)],
        margin=avg(ratio(h["ebitda"])), da=avg(ratio(h["da"])), capex=avg(ratio(h["capex"])),
        ar=avg(ratio(h["ar"])), inv=avg(ratio(h["inv"])), ap=avg(ratio(h["ap"])),
        tax=clip(avg(tax_r), 0.10, 0.35) if tax_r else 0.25,
        int=clip(avg(int_r), 0.02, 0.12) if int_r else 0.06,
        payout=clip(avg(pay_r), 0.0, 1.0) if pay_r else 0.0,
        exit_x=round(clip(ev_x, 6, 20) * 2) / 2,
    )


# ----------------------------------------------------------------------------- Assumptions
def build_assumptions(wb, d, lay, df):
    ws = wb["Assumptions"]
    title(ws, "Assumptions & market data", "Blue = input you can change. Yellow = key assumption. Black = formula. Green = link to another sheet.")
    widths(ws, 40, 16, 3)
    ws.column_dimensions["C"].width = 90
    A, r = {}, [4]

    def sec(text):
        section(ws, r[0], text, 3)
        r[0] += 1

    def row(key, label, val, fmt=None, note=None, key_input=False, kind=None):
        put(ws, f"A{r[0]}", label, kind="text")
        put(ws, f"B{r[0]}", val, fmt=fmt, kind=kind, fill=YELLOW if key_input else None, align="right")
        if note:
            put(ws, f"C{r[0]}", note, italic=True, kind="text", color="595959")
        A[key] = f"Assumptions!$B${r[0]}"
        A[key + "_l"] = f"$B${r[0]}"
        r[0] += 1

    h, LH = d["hist"], lay.LH
    sec("Company")
    row("ticker", "Ticker", d["ticker"], kind="text")
    row("name", "Name", d["name"], kind="text")
    row("ccy", "Currency", d["currency"], kind="text", note="All statement values in millions of this currency.")
    row("sector", "Sector / industry", f"{d['sector']} / {d['industry']}", kind="text")
    row("asof", "Data as of", d["asof"], kind="text", note=f"Source: {d['source']}")
    r[0] += 1
    sec("Market data")
    row("price", "Share price", round(d["price"], 4), F_PS, "Latest price from data source.")
    row("shares", "Shares outstanding (mm)", round(d["shares"], 4), '#,##0.00', "From data source (basic shares; add dilution if relevant).")
    row("mcap", "Market cap ($ mm)", f"={A['price_l']}*{A['shares_l']}", F_NUM)
    row("debt", "Total debt ($ mm, latest FY)", f"=ThreeStatement!{LH}{TS['debt']}", F_NUM, "Latest fiscal-year balance sheet; may lag the current quarter.")
    row("cash", "Cash & ST investments ($ mm, latest FY)", f"=ThreeStatement!{LH}{TS['cash']}", F_NUM)
    row("netdebt", "Net debt ($ mm)", f"={A['debt_l']}-{A['cash_l']}", F_NUM)
    r[0] += 1
    sec("Scenario")
    row("scen", "Active scenario (1=Base, 2=Bull, 3=Bear)", 1, "0", "Adjustments per scenario are on the Forecast sheet.", key_input=True)
    dv = DataValidation(type="list", formula1='"1,2,3"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(A["scen_l"].replace("$", ""))
    r[0] += 1
    sec("Cost of capital (CAPM / WACC)")
    row("rf", "Risk-free rate", df["rf"], F_PCT, "PLACEHOLDER default - update to the current 10-yr government bond yield.", key_input=True)
    row("erp", "Equity risk premium", df["erp"], F_PCT, "PLACEHOLDER default (Damodaran-style). Update for your market/country risk.", key_input=True)
    row("beta", "Beta", round(d["beta"], 3), '0.00', "Yahoo Finance beta (or 1.0 if unavailable).", key_input=True)
    row("ke", "Cost of equity", f"={A['rf_l']}+{A['beta_l']}*{A['erp_l']}", F_PCT)
    row("kd", "Pre-tax cost of debt", round(df["int"], 4), F_PCT, "Default = historical interest / debt, clipped 2%-12%.", key_input=True)
    row("kd_tax", "Tax rate for debt shield", f"=Forecast!{lay.F}{FC['tax']}", F_PCT, "Linked to first projected year's tax rate.")
    row("kd_post", "After-tax cost of debt", f"={A['kd_l']}*(1-{A['kd_tax_l']})", F_PCT)
    row("we", "Equity weight (market cap)", f"=IFERROR({A['mcap_l']}/({A['mcap_l']}+{A['debt_l']}),1)", F_PCT)
    row("wd", "Debt weight", f"=1-{A['we_l']}", F_PCT)
    row("wacc_override", "WACC override (optional)", 0, F_PCT,
        "Set a positive value to use your own WACC; leave at 0 to use the CAPM / capital-structure calculation.", key_input=True)
    row("wacc", "WACC", f"=IF({A['wacc_override_l']}>0,{A['wacc_override_l']},{A['we_l']}*{A['ke_l']}+{A['wd_l']}*{A['kd_post_l']})", F_PCT,
        "Default = equity weight × CAPM cost of equity + debt weight × after-tax cost of debt. Feeds the DCF.")
    r[0] += 1
    sec("Terminal value")
    row("g", "Terminal growth rate", df["g_term"], F_PCT, "Must stay below WACC and long-run nominal GDP growth.", key_input=True)
    row("exit_x", "Exit EV/EBITDA multiple (cross-check)", df["exit_x"], F_X, "Default = current EV/EBITDA rounded to 0.5x, clipped 6x-20x.", key_input=True)
    ws.freeze_panes = "A4"
    return A


# ----------------------------------------------------------------------------- Three-statement
def build_three_statement(wb, d, lay):
    ws = wb["ThreeStatement"]
    title(ws, "Three-statement model", "Historicals (blue) from data source; projections are formulas driven by the Forecast sheet. Balance sheet balances by construction (check row).")
    widths(ws, 40, 12, 12)
    year_header(ws, lay, 4)
    h = d["hist"]
    R = TS

    for txt, rr in (("Income statement", 5), ("Balance sheet", 16), ("Cash flow", 31)):
        section(ws, rr, txt, 1 + len(lay.cols))
    labels = {
        "rev": "Revenue", "ebitda": "EBITDA", "da": "Depreciation & amortisation", "ebit": "EBIT",
        "interest": "Interest expense", "other": "Other income/(expense), net", "ebt": "Pre-tax income", "tax": "Income tax", "ni": "Net income",
        "cash": "Cash & ST investments", "ar": "Receivables", "inv": "Inventory", "ppe": "Net PP&E", "oth_a": "Other assets",
        "ta": "Total assets", "ap": "Payables", "debt": "Total debt", "oth_l": "Other liabilities", "tl": "Total liabilities",
        "eq": "Shareholders' equity", "tle": "Total liabilities & equity", "check": "Balance check (should be 0)",
        "cf_ni": "Net income", "cf_da": "Add: D&A", "cf_ar": "(Increase)/decrease in receivables", "cf_inv": "(Increase)/decrease in inventory",
        "cf_ap": "Increase/(decrease) in payables", "cfo": "Cash flow from operations", "capex": "Capital expenditure (outflow)",
        "div": "Dividends (outflow)", "cf_debt": "Net debt issued/(repaid)", "netchg": "Net change in cash"}
    bold = {"rev", "ebitda", "ebit", "ni", "ta", "tl", "tle", "cfo", "netchg", "eq"}
    for k, lab in labels.items():
        put(ws, f"A{R[k]}", lab, kind="text", bold=k in bold)

    # historical inputs
    plug_a = [ta - c - a - i - p for ta, c, a, i, p in zip(h["total_assets"], h["cash"], h["ar"], h["inv"], h["ppe"])]
    plug_l = [ta - e - ap - db for ta, e, ap, db in zip(h["total_assets"], h["equity"], h["ap"], h["debt"])]
    other = [ni - (eb - dep - itr - tx) for ni, eb, dep, itr, tx in zip(h["ni"], h["ebitda"], h["da"], h["interest"], h["tax"])]
    inputs = dict(rev=h["rev"], ebitda=h["ebitda"], da=h["da"], interest=h["interest"], other=other, tax=h["tax"],
                  cash=h["cash"], ar=h["ar"], inv=h["inv"], ppe=h["ppe"], oth_a=plug_a, ap=h["ap"], debt=h["debt"],
                  oth_l=plug_l, eq=h["equity"], capex=h["capex"], div=h["div"])
    for i, c in enumerate(lay.HC):
        for k, vals in inputs.items():
            fmt = F_NUM
            put(ws, f"{c}{R[k]}", round(vals[i], 3), fmt)
        put(ws, f"{c}{R['ebit']}", f"={c}{R['ebitda']}-{c}{R['da']}", F_NUM, bold=True)
        put(ws, f"{c}{R['ebt']}", f"={c}{R['ebit']}-{c}{R['interest']}+{c}{R['other']}", F_NUM)
        put(ws, f"{c}{R['ni']}", f"={c}{R['ebt']}-{c}{R['tax']}", F_NUM, bold=True)
        put(ws, f"{c}{R['ta']}", f"=SUM({c}{R['cash']}:{c}{R['oth_a']})", F_NUM, bold=True)
        put(ws, f"{c}{R['tl']}", f"=SUM({c}{R['ap']}:{c}{R['oth_l']})", F_NUM, bold=True)
        put(ws, f"{c}{R['tle']}", f"={c}{R['tl']}+{c}{R['eq']}", F_NUM, bold=True)
        put(ws, f"{c}{R['check']}", f"=ROUND({c}{R['ta']}-{c}{R['tle']},3)", F_NUM)

    # projections
    for c in lay.PC:
        p = lay.prev(c)
        f = lambda k: f"Forecast!{c}{FC[k]}"
        put(ws, f"{c}{R['rev']}", f"={p}{R['rev']}*(1+{f('g_act')})", F_NUM, bold=True)
        put(ws, f"{c}{R['ebitda']}", f"={c}{R['rev']}*{f('m_act')}", F_NUM, bold=True)
        put(ws, f"{c}{R['da']}", f"={c}{R['rev']}*{f('da')}", F_NUM)
        put(ws, f"{c}{R['ebit']}", f"={c}{R['ebitda']}-{c}{R['da']}", F_NUM, bold=True)
        put(ws, f"{c}{R['interest']}", f"={p}{R['debt']}*{f('int')}", F_NUM)
        put(ws, f"{c}{R['other']}", 0, F_NUM)
        put(ws, f"{c}{R['ebt']}", f"={c}{R['ebit']}-{c}{R['interest']}+{c}{R['other']}", F_NUM)
        put(ws, f"{c}{R['tax']}", f"=MAX(0,{c}{R['ebt']})*{f('tax')}", F_NUM)
        put(ws, f"{c}{R['ni']}", f"={c}{R['ebt']}-{c}{R['tax']}", F_NUM, bold=True)
        put(ws, f"{c}{R['cash']}", f"={p}{R['cash']}+{c}{R['netchg']}", F_NUM)
        put(ws, f"{c}{R['ar']}", f"={c}{R['rev']}*{f('ar')}", F_NUM)
        put(ws, f"{c}{R['inv']}", f"={c}{R['rev']}*{f('inv')}", F_NUM)
        put(ws, f"{c}{R['ppe']}", f"={p}{R['ppe']}+{c}{R['capex']}-{c}{R['da']}", F_NUM)
        put(ws, f"{c}{R['oth_a']}", f"={p}{R['oth_a']}", F_NUM)
        put(ws, f"{c}{R['ta']}", f"=SUM({c}{R['cash']}:{c}{R['oth_a']})", F_NUM, bold=True)
        put(ws, f"{c}{R['ap']}", f"={c}{R['rev']}*{f('ap')}", F_NUM)
        put(ws, f"{c}{R['debt']}", f"={p}{R['debt']}+{f('dchg')}", F_NUM)
        put(ws, f"{c}{R['oth_l']}", f"={p}{R['oth_l']}", F_NUM)
        put(ws, f"{c}{R['tl']}", f"=SUM({c}{R['ap']}:{c}{R['oth_l']})", F_NUM, bold=True)
        put(ws, f"{c}{R['eq']}", f"={p}{R['eq']}+{c}{R['ni']}-{c}{R['div']}", F_NUM, bold=True)
        put(ws, f"{c}{R['tle']}", f"={c}{R['tl']}+{c}{R['eq']}", F_NUM, bold=True)
        put(ws, f"{c}{R['check']}", f"=ROUND({c}{R['ta']}-{c}{R['tle']},3)", F_NUM)
        put(ws, f"{c}{R['cf_ni']}", f"={c}{R['ni']}", F_NUM)
        put(ws, f"{c}{R['cf_da']}", f"={c}{R['da']}", F_NUM)
        put(ws, f"{c}{R['cf_ar']}", f"={p}{R['ar']}-{c}{R['ar']}", F_NUM)
        put(ws, f"{c}{R['cf_inv']}", f"={p}{R['inv']}-{c}{R['inv']}", F_NUM)
        put(ws, f"{c}{R['cf_ap']}", f"={c}{R['ap']}-{p}{R['ap']}", F_NUM)
        put(ws, f"{c}{R['cfo']}", f"=SUM({c}{R['cf_ni']}:{c}{R['cf_ap']})", F_NUM, bold=True)
        put(ws, f"{c}{R['capex']}", f"={c}{R['rev']}*{f('capex')}", F_NUM)
        put(ws, f"{c}{R['div']}", f"=MAX(0,{c}{R['ni']})*{f('payout')}", F_NUM)
        put(ws, f"{c}{R['cf_debt']}", f"={c}{R['debt']}-{p}{R['debt']}", F_NUM)
        put(ws, f"{c}{R['netchg']}", f"={c}{R['cfo']}-{c}{R['capex']}-{c}{R['div']}+{c}{R['cf_debt']}", F_NUM, bold=True)

    put(ws, f"A{R['netchg'] + 2}", "Notes: 'Other assets/liabilities' historicals are plugs = reported totals less itemised lines. "
        "'Other income' historical = reported net income less (EBIT - interest - tax). Interest uses prior-year debt (no circularity). "
        "Projected cash can go negative if the plan under-funds itself: that signals a financing need, not an error.", italic=True, kind="text", color="595959")

    # Keep the reported statements visibly separate from the normalised model rows
    # above. This gives users the provider-sourced P&L, cash-flow and balance-sheet
    # figures they can reconcile before using the forecast.
    raw = d.get("raw", {})
    start = R["netchg"] + 4
    put(ws, f"A{start}", "Reported source statements — historical data fetched from provider", bold=True, kind="text", color=NAVY)
    put(ws, f"A{start + 1}", "These are reported historical line items where available; the model above uses normalised rows so the three statements link and forecast cleanly.",
        italic=True, kind="text", color="595959")

    def source_table(row, heading, items):
        section(ws, row, heading, 1 + len(lay.HC))
        header(ws, row + 1, ["($ mm, reporting currency)"] + [f"FY{y}A" for y in d["years"]])
        for i, (key, label) in enumerate(items):
            rr = row + 2 + i
            put(ws, f"A{rr}", label, kind="text", bold=key in ("revenue", "gross_profit", "ebitda", "net_income", "cfo", "total_assets", "equity"))
            vals = raw.get(key, [])
            for j, col in enumerate(lay.HC):
                value = vals[j] if j < len(vals) else None
                put(ws, f"{col}{rr}", "n/a" if value is None else round(value, 3), F_NUM if value is not None else None,
                    kind="text" if value is None else "input", bold=key in ("revenue", "gross_profit", "ebitda", "net_income", "cfo", "total_assets", "equity"))
        return row + 2 + len(items)

    last = source_table(start + 3, "Reported P&L / income statement", [
        ("revenue", "Revenue"), ("cogs", "Cost of revenue / COGS"), ("gross_profit", "Gross profit"),
        ("ebitda", "EBITDA"), ("ebit", "EBIT / operating income"), ("interest", "Interest expense"),
        ("tax", "Income tax"), ("net_income", "Net income")])
    last = source_table(last + 2, "Reported cash-flow statement", [
        ("cfo", "Cash flow from operations"), ("capex", "Capital expenditure"), ("dividends", "Dividends paid"),
        ("debt_issued", "Debt issued"), ("debt_repaid", "Debt repaid")])
    source_table(last + 2, "Reported balance sheet", [
        ("cash", "Cash & short-term investments"), ("receivables", "Accounts receivable"), ("inventory", "Inventory"),
        ("ppe", "Net PP&E"), ("total_assets", "Total assets"), ("payables", "Accounts payable"),
        ("debt", "Total debt"), ("equity", "Shareholders' equity")])
    ws.freeze_panes = "B5"


# ----------------------------------------------------------------------------- Forecast
def build_reported_financials(wb, d, lay):
    """Standalone source view: actual P&L, cash flow and balance-sheet figures."""
    ws = wb["Reported Financials"]
    title(ws, "Reported financial statements", "Historical financials fetched from the data provider. These are source figures, not forecast formulas.")
    widths(ws, 44, 14, len(lay.HC) + 1)
    raw = d.get("raw", {})

    def source_table(row, heading, items):
        section(ws, row, heading, 1 + len(lay.HC))
        header(ws, row + 1, ["($ mm, reporting currency)"] + [f"FY{y}A" for y in d["years"]])
        for i, (key, label) in enumerate(items):
            rr = row + 2 + i
            is_total = key in {"revenue", "gross_profit", "ebitda", "net_income", "cfo", "total_assets", "equity"}
            put(ws, f"A{rr}", label, kind="text", bold=is_total)
            values = raw.get(key, [])
            for j, col in enumerate(lay.HC):
                value = values[j] if j < len(values) else None
                put(ws, f"{col}{rr}", "n/a" if value is None else round(value, 3), F_NUM if value is not None else None,
                    kind="text" if value is None else "input", bold=is_total, align="right")
        return row + 2 + len(items)

    last = source_table(4, "P&L / income statement", [
        ("revenue", "Revenue"), ("cogs", "Cost of revenue / COGS"), ("gross_profit", "Gross profit"),
        ("ebitda", "EBITDA"), ("ebit", "EBIT / operating income"), ("interest", "Interest expense"),
        ("tax", "Income tax"), ("net_income", "Net income")])
    last = source_table(last + 2, "Cash-flow statement", [
        ("cfo", "Cash flow from operations"), ("capex", "Capital expenditure"), ("dividends", "Dividends paid"),
        ("debt_issued", "Debt issued"), ("debt_repaid", "Debt repaid")])
    source_table(last + 2, "Balance sheet", [
        ("cash", "Cash & short-term investments"), ("receivables", "Accounts receivable"), ("inventory", "Inventory"),
        ("ppe", "Net PP&E"), ("total_assets", "Total assets"), ("payables", "Accounts payable"),
        ("debt", "Total debt"), ("equity", "Shareholders' equity")])
    ws.freeze_panes = "B6"


# ----------------------------------------------------------------------------- Forecast
def build_forecast(wb, d, lay, df, A):
    ws = wb["Forecast"]
    title(ws, "Forecasting model: operating drivers & scenarios", "Historical columns show the ratio implied by the reported numbers; projected columns are the drivers (blue). Defaults are simple historical averages/fades: replace with your own view.")
    widths(ws, 44, 12, 12)
    F = FC
    section(ws, 4, "Scenario levers", 1 + len(lay.cols))
    put(ws, f"A{F['scen']}", "Active scenario (set on Assumptions)", kind="text")
    put(ws, f"B{F['scen']}", f"={A['scen']}", "0")
    put(ws, f"A{F['scen_name']}", "Scenario name", kind="text")
    put(ws, f"B{F['scen_name']}", f'=CHOOSE($B${F["scen"]},"Base","Bull","Bear")', kind="calc", align="right")
    for k, lab, v in (("bull_g", "Bull: revenue growth adjustment (pp per year)", 0.02), ("bear_g", "Bear: revenue growth adjustment (pp per year)", -0.03),
                      ("bull_m", "Bull: EBITDA margin adjustment (pp)", 0.02), ("bear_m", "Bear: EBITDA margin adjustment (pp)", -0.03)):
        put(ws, f"A{F[k]}", lab, kind="text")
        put(ws, f"B{F[k]}", v, F_PCT, fill=YELLOW)
    put(ws, f"A{F['adj_g']}", "Active growth adjustment", kind="text")
    put(ws, f"B{F['adj_g']}", f"=CHOOSE($B${F['scen']},0,$B${F['bull_g']},$B${F['bear_g']})", F_PCT)
    put(ws, f"A{F['adj_m']}", "Active margin adjustment", kind="text")
    put(ws, f"B{F['adj_m']}", f"=CHOOSE($B${F['scen']},0,$B${F['bull_m']},$B${F['bear_m']})", F_PCT)

    year_header(ws, lay, F["hdr"], "Drivers")
    rows = [("g_base", "Revenue growth: base case"), ("g_act", "Revenue growth: active scenario"),
            ("m_base", "EBITDA margin: base case"), ("m_act", "EBITDA margin: active scenario"),
            ("da", "D&A % revenue"), ("capex", "Capex % revenue"), ("ar", "Receivables % revenue"),
            ("inv", "Inventory % revenue"), ("ap", "Payables % revenue"), ("tax", "Effective tax rate"),
            ("int", "Interest rate on prior-year debt"), ("payout", "Dividend payout (% net income)"), ("dchg", "Net debt issued/(repaid) ($ mm)")]
    for k, lab in rows:
        put(ws, f"A{F[k]}", lab, kind="text", bold=k in ("g_act", "m_act"))

    T = "ThreeStatement!"
    ratios = dict(m_base=("ebitda", "rev"), da=("da", "rev"), capex=("capex", "rev"), ar=("ar", "rev"), inv=("inv", "rev"),
                  ap=("ap", "rev"), tax=("tax", "ebt"), int=("interest", "debt"), payout=("div", "ni"))
    for i, c in enumerate(lay.HC):
        if i > 0:
            p = lay.prev(c)
            put(ws, f"{c}{F['g_base']}", f"=IFERROR({T}{c}{TS['rev']}/{T}{p}{TS['rev']}-1,0)", F_PCT)
        for k, (num, den) in ratios.items():
            put(ws, f"{c}{F[k]}", f"=IFERROR({T}{c}{TS[num]}/{T}{c}{TS[den]},0)", F_PCT)
    inp = dict(m_base=df["margin"], da=df["da"], capex=df["capex"], ar=df["ar"], inv=df["inv"], ap=df["ap"],
               tax=df["tax"], int=df["int"], payout=df["payout"])
    for j, c in enumerate(lay.PC):
        put(ws, f"{c}{F['g_base']}", round(df["growth"][j], 4), F_PCT)
        put(ws, f"{c}{F['g_act']}", f"={c}{F['g_base']}+$B${F['adj_g']}", F_PCT, bold=True)
        put(ws, f"{c}{F['m_act']}", f"={c}{F['m_base']}+$B${F['adj_m']}", F_PCT, bold=True)
        for k, v in inp.items():
            put(ws, f"{c}{F[k]}", round(v, 4), F_PCT)
        put(ws, f"{c}{F['dchg']}", 0, F_NUM)
    put(ws, f"A{F['dchg'] + 2}", "Default logic: growth starts at the historical revenue CAGR (clipped -5%..25%) and fades linearly to the terminal growth rate; "
        "margins and working-capital/capex ratios are historical averages. Historical data is only 3 years, so treat defaults as a starting point.",
        italic=True, kind="text", color="595959")
    ws.freeze_panes = f"B{F['hdr'] + 1}"


# ----------------------------------------------------------------------------- DCF
def build_dcf(wb, lay, A):
    ws = wb["DCF"]
    title(ws, "Discounted cash flow (FCFF, end-of-year discounting)", "UFCF = EBIT x (1 - tax) + D&A - capex - increase in net working capital. WACC and terminal growth come from Assumptions.")
    widths(ws, 44, 13, 12)
    D, T = DC, "ThreeStatement!"
    year_header(ws, lay, D["hdr"])
    labs = dict(rev="Revenue", ebitda="EBITDA", ebit="EBIT", taxes="Less: taxes on EBIT", nopat="NOPAT", da="Add: D&A",
                capex="Less: capex", nwc="Net working capital (AR + inventory - AP)", dnwc="Less: increase in NWC",
                ufcf="Unlevered free cash flow", n="Discount period (years)", df="Discount factor", pv="PV of UFCF")
    for k, lab in labs.items():
        put(ws, f"A{D[k]}", lab, kind="text", bold=k in ("ufcf", "ebitda", "nopat"))
    for c in lay.HC:
        put(ws, f"{c}{D['rev']}", f"={T}{c}{TS['rev']}", F_NUM)
        put(ws, f"{c}{D['ebitda']}", f"={T}{c}{TS['ebitda']}", F_NUM)
        put(ws, f"{c}{D['nwc']}", f"={T}{c}{TS['ar']}+{T}{c}{TS['inv']}-{T}{c}{TS['ap']}", F_NUM)
    for j, c in enumerate(lay.PC):
        p = lay.prev(c)
        put(ws, f"{c}{D['rev']}", f"={T}{c}{TS['rev']}", F_NUM)
        put(ws, f"{c}{D['ebitda']}", f"={T}{c}{TS['ebitda']}", F_NUM, bold=True)
        put(ws, f"{c}{D['ebit']}", f"={T}{c}{TS['ebit']}", F_NUM)
        put(ws, f"{c}{D['taxes']}", f"=-{c}{D['ebit']}*Forecast!{c}{FC['tax']}", F_NUM)
        put(ws, f"{c}{D['nopat']}", f"={c}{D['ebit']}+{c}{D['taxes']}", F_NUM, bold=True)
        put(ws, f"{c}{D['da']}", f"={T}{c}{TS['da']}", F_NUM)
        put(ws, f"{c}{D['capex']}", f"=-{T}{c}{TS['capex']}", F_NUM)
        put(ws, f"{c}{D['nwc']}", f"={T}{c}{TS['ar']}+{T}{c}{TS['inv']}-{T}{c}{TS['ap']}", F_NUM)
        put(ws, f"{c}{D['dnwc']}", f"=-({c}{D['nwc']}-{p}{D['nwc']})", F_NUM)
        put(ws, f"{c}{D['ufcf']}", f"={c}{D['nopat']}+{c}{D['da']}+{c}{D['capex']}+{c}{D['dnwc']}", F_NUM, bold=True)
        put(ws, f"{c}{D['n']}", 1 if j == 0 else f"={p}{D['n']}+1", "0", kind="calc")
        put(ws, f"{c}{D['df']}", f"=1/(1+$B${D['wacc']})^{c}{D['n']}", F_D)
        put(ws, f"{c}{D['pv']}", f"={c}{D['ufcf']}*{c}{D['df']}", F_NUM)

    F, L = lay.F, lay.LP
    section(ws, D["val"], "Valuation: perpetuity growth method", 1 + len(lay.cols))
    v = lambda k: f"$B${D[k]}"
    rows = [
        ("wacc", "WACC", f"={A['wacc']}", F_PCT), ("g", "Terminal growth", f"={A['g']}", F_PCT),
        ("sumpv", "Sum of PV of UFCF", f"=SUM({F}{D['pv']}:{L}{D['pv']})", F_NUM),
        ("tvcf", "Terminal-year UFCF x (1+g)", f"={L}{D['ufcf']}*(1+{v('g')})", F_NUM),
        ("tv", "Terminal value", f"=IF({v('wacc')}>{v('g')},{v('tvcf')}/({v('wacc')}-{v('g')}),0)", F_NUM),
        ("pvtv", "PV of terminal value", f"={v('tv')}*{L}{D['df']}", F_NUM),
        ("ev", "Enterprise value", f"={v('sumpv')}+{v('pvtv')}", F_NUM),
        ("debt", "Less: total debt", f"={A['debt']}", F_NUM), ("cash", "Add: cash", f"={A['cash']}", F_NUM),
        ("eq", "Equity value", f"={v('ev')}-{v('debt')}+{v('cash')}", F_NUM),
        ("shares", "Shares outstanding (mm)", f"={A['shares']}", '#,##0.00'),
        ("ps", "Implied value per share", f"=IFERROR({v('eq')}/{v('shares')},0)", F_PS),
        ("price", "Current share price", f"={A['price']}", F_PS),
        ("up", "Upside / (downside)", f"=IFERROR({v('ps')}/{v('price')}-1,0)", F_PCT),
        ("tvpct", "PV of terminal value as % of EV", f"=IFERROR({v('pvtv')}/{v('ev')},0)", F_PCT),
        ("impx", "Implied terminal EV/EBITDA", f"=IFERROR({v('tv')}/{L}{D['ebitda']},0)", F_X),
        ("chk", "Check: WACC > terminal growth", f'=IF({v("wacc")}>{v("g")},"OK","WACC must exceed g")', None)]
    for k, lab, fml, fmt in rows:
        put(ws, f"A{D[k]}", lab, kind="text", bold=k in ("ev", "eq", "ps"))
        put(ws, f"B{D[k]}", fml, fmt, bold=k in ("ev", "eq", "ps"), align="right")

    section(ws, D["x_sec"], "Cross-check: exit multiple method", 1 + len(lay.cols))
    xrows = [("xm", "Exit EV/EBITDA multiple", f"={A['exit_x']}", F_X),
             ("xtv", "Terminal value (final-year EBITDA x multiple)", f"={L}{D['ebitda']}*{v('xm')}", F_NUM),
             ("xpv", "PV of terminal value", f"={v('xtv')}*{L}{D['df']}", F_NUM),
             ("xev", "Enterprise value", f"={v('sumpv')}+{v('xpv')}", F_NUM),
             ("xeq", "Equity value", f"={v('xev')}-{v('debt')}+{v('cash')}", F_NUM),
             ("xps", "Implied value per share", f"=IFERROR({v('xeq')}/{v('shares')},0)", F_PS)]
    for k, lab, fml, fmt in xrows:
        put(ws, f"A{D[k]}", lab, kind="text", bold=k == "xps")
        put(ws, f"B{D[k]}", fml, fmt, bold=k == "xps", align="right")

    section(ws, D["s_sec"], "Sensitivity: value per share, WACC (down) vs terminal growth (across)", 1 + len(lay.cols))
    put(ws, f"A{D['wstep']}", "WACC step", kind="text")
    put(ws, f"B{D['wstep']}", 0.01, F_PCT)
    put(ws, f"A{D['gstep']}", "Terminal growth step", kind="text")
    put(ws, f"B{D['gstep']}", 0.005, F_PCT)
    put(ws, f"B{D['s_hdr']}", "WACC \\ g", kind="text", bold=True, align="center")
    gcols = ["C", "D", "E", "F", "G"]
    for j, c in enumerate(gcols):
        put(ws, f"{c}{D['s_hdr']}", f"=$B${D['g']}+({j - 2})*$B${D['gstep']}", F_PCT, bold=True, fill=GREY)
    ufcf = f"${F}${D['ufcf']}:${L}${D['ufcf']}"
    nn = f"${F}${D['n']}:${L}${D['n']}"
    for i in range(5):
        rr = D["s0"] + i
        put(ws, f"B{rr}", f"=$B${D['wacc']}+({i - 2})*$B${D['wstep']}", F_PCT, bold=True, fill=GREY)
        for c in gcols:
            w, g = f"$B{rr}", f"{c}${D['s_hdr']}"
            fml = (f'=IF({w}<={g},"n/a",(SUMPRODUCT({ufcf}/(1+{w})^{nn})+${L}${D["ufcf"]}*(1+{g})/({w}-{g})/(1+{w})^${L}${D["n"]}'
                   f'-$B${D["debt"]}+$B${D["cash"]})/$B${D["shares"]})')
            put(ws, f"{c}{rr}", fml, F_PS)
    ws.freeze_panes = "B5"


# ----------------------------------------------------------------------------- Comps
def build_comps(wb, d, A):
    ws = wb["Comps"]
    title(ws, "Comparable company analysis", "Peers auto-selected by industry from the data source: REVIEW AND EDIT (blue cells). Multiples are unitless so cross-currency peers are tolerable, but size/geography differences matter.")
    widths(ws, 34, 12, 12)
    ws.column_dimensions["B"].width = 12
    header(ws, 4, ["Company", "Ticker", "Mkt cap", "Debt", "Cash", "EV", "Revenue", "EBITDA", "Net income", "EV/Revenue", "EV/EBITDA", "P/E"])
    t = d["target_ttm"]
    peers = d["peers"] or [dict(name="(enter peer)", ticker="", mcap=0, debt=0, cash=0, rev=0, ebitda=0, ni=0)] * 3

    def comp_row(rr, p, target=False):
        put(ws, f"A{rr}", p["name"], kind="text", bold=target, fill=None if d["peers"] or target else YELLOW)
        put(ws, f"B{rr}", p["ticker"], kind="text")
        if target:
            put(ws, f"C{rr}", f"={A['mcap']}", F_NUM)
        else:
            put(ws, f"C{rr}", round(p["mcap"], 1), F_NUM)
        for col, k in (("D", "debt"), ("E", "cash"), ("G", "rev"), ("H", "ebitda"), ("I", "ni")):
            put(ws, f"{col}{rr}", round(p[k], 1), F_NUM)
        put(ws, f"F{rr}", f"=C{rr}+D{rr}-E{rr}", F_NUM)
        put(ws, f"J{rr}", f'=IF(G{rr}>0,F{rr}/G{rr},"n/m")', F_X, align="right")
        put(ws, f"K{rr}", f'=IF(H{rr}>0,F{rr}/H{rr},"n/m")', F_X, align="right")
        put(ws, f"L{rr}", f'=IF(I{rr}>0,C{rr}/I{rr},"n/m")', F_X, align="right")

    comp_row(5, t | {"name": d["name"], "ticker": d["ticker"]}, target=True)
    put(ws, "A6", "Peer group (TTM figures, $ mm)", bold=True, kind="text")
    r0 = 7
    for i, p in enumerate(peers):
        comp_row(r0 + i, p)
    r1 = r0 + len(peers) - 1
    stats = {}
    for j, (lab, fn) in enumerate((("Median", "MEDIAN"), ("Mean", "AVERAGE"), ("High", "MAX"), ("Low", "MIN"))):
        rr = r1 + 2 + j
        stats[lab] = rr
        put(ws, f"A{rr}", lab, kind="text", bold=True)
        for col in "JKL":
            put(ws, f"{col}{rr}", f'=IFERROR({fn}({col}{r0}:{col}{r1}),"n/a")', F_X, bold=True, align="right")
    m = stats["Median"]

    b = stats["Low"] + 3
    section(ws, b, "Implied valuation of target (peer median multiple)", 12)
    header(ws, b + 1, ["Method", "Median multiple", "Target metric", "Implied EV", "Net debt", "Implied equity", "Per share", "vs price"])
    CR = {}
    specs = (("EV/EBITDA", f"K{m}", "H5", True), ("EV/Revenue", f"J{m}", "G5", True), ("P/E", f"L{m}", "I5", False))
    for i, (lab, mult, metric, is_ev) in enumerate(specs):
        rr = b + 2 + i
        put(ws, f"A{rr}", lab, kind="text", bold=True)
        put(ws, f"B{rr}", f"={mult}", F_X, align="right")
        put(ws, f"C{rr}", f"={metric}", F_NUM)
        if is_ev:
            put(ws, f"D{rr}", f'=IFERROR(B{rr}*C{rr},"n/a")', F_NUM, align="right")
            put(ws, f"E{rr}", "=D5-E5", F_NUM)
            put(ws, f"F{rr}", f'=IFERROR(D{rr}-E{rr},"n/a")', F_NUM, align="right")
        else:
            put(ws, f"D{rr}", "n/a", kind="text", align="right")
            put(ws, f"E{rr}", "n/a", kind="text", align="right")
            put(ws, f"F{rr}", f'=IFERROR(B{rr}*C{rr},"n/a")', F_NUM, align="right")
        put(ws, f"G{rr}", f'=IFERROR(F{rr}/{A["shares"]},"n/a")', F_PS, bold=True, align="right")
        put(ws, f"H{rr}", f'=IFERROR(G{rr}/{A["price"]}-1,"n/a")', F_PCT, align="right")
        CR[lab] = f"Comps!$G${rr}"
    put(ws, f"A{b + 6}", "Negative or zero denominators are shown as n/m and excluded from statistics. P/E uses net income to common; EV uses market cap + debt - cash.",
        italic=True, kind="text", color="595959")
    ws.freeze_panes = "C5"
    return CR


# ----------------------------------------------------------------------------- Options
def build_options(wb, d, A):
    ws = wb["Options"]
    title(ws, "Option pricing: Black-Scholes-Merton (European, continuous dividend yield)", "Prices a call and put on the stock. For a live listed option, compare with the market quote below. American early-exercise premium is ignored.")
    widths(ws, 44, 14, 12)
    o = d["option"]
    price = d["price"]
    step = 5 if price >= 100 else 1
    strike = o["strike"] if o else round(price / step) * step
    T = round(o["days"] / 365, 4) if o else 0.25
    section(ws, 4, "Inputs", 6)
    inputs = [("S", "Spot price", f"={A['price']}", F_PS, None),
              ("K", "Strike", strike, F_PS, "Nearest-ATM strike from option chain." if o else "Default ATM (rounded)."),
              ("T", "Time to expiry (years)", T, '0.000', "Days / 365." if o else "Default 3 months."),
              ("r", "Risk-free rate", f"={A['rf']}", F_PCT, "Linked to Assumptions; use a rate matching the expiry."),
              ("sig", "Volatility (annualised)", round(d["vol"], 4), F_PCT, "1-year historical close-to-close vol. Replace with implied vol for market-consistent prices."),
              ("q", "Dividend yield", round(d["div_yield"], 4), F_PCT, "Trailing dividend / price.")]
    O = {}
    for i, (k, lab, v, fmt, note) in enumerate(inputs):
        rr = 5 + i
        O[k] = f"$B${rr}"
        put(ws, f"A{rr}", lab, kind="text")
        put(ws, f"B{rr}", v, fmt, fill=YELLOW if k in ("K", "T", "sig") else None)
        if note:
            put(ws, f"C{rr}", note, italic=True, kind="text", color="595959")
    S, K, Tt, r, sg, q = (O[k] for k in ("S", "K", "T", "r", "sig", "q"))
    section(ws, 12, "Outputs", 6)
    put(ws, "A13", "d1", kind="text")
    put(ws, "B13", f"=(LN({S}/{K})+({r}-{q}+{sg}^2/2)*{Tt})/({sg}*SQRT({Tt}))", F_D)
    put(ws, "A14", "d2", kind="text")
    put(ws, "B14", f"=B13-{sg}*SQRT({Tt})", F_D)
    put(ws, "A15", "Call price", kind="text", bold=True)
    put(ws, "B15", f"={S}*EXP(-{q}*{Tt})*NORMSDIST(B13)-{K}*EXP(-{r}*{Tt})*NORMSDIST(B14)", F_PS, bold=True)
    put(ws, "A16", "Put price", kind="text", bold=True)
    put(ws, "B16", f"={K}*EXP(-{r}*{Tt})*NORMSDIST(-B14)-{S}*EXP(-{q}*{Tt})*NORMSDIST(-B13)", F_PS, bold=True)
    put(ws, "A17", "Put-call parity check (should be 0)", kind="text")
    put(ws, "B17", f"=ROUND(B15-B16-({S}*EXP(-{q}*{Tt})-{K}*EXP(-{r}*{Tt})),6)", F_D)

    header(ws, 19, ["Greeks", "Call", "Put"])
    pdf = f"EXP(-B13^2/2)/SQRT(2*PI())"
    greeks = [
        ("Delta", f"=EXP(-{q}*{Tt})*NORMSDIST(B13)", f"=EXP(-{q}*{Tt})*(NORMSDIST(B13)-1)"),
        ("Gamma", f"=EXP(-{q}*{Tt})*{pdf}/({S}*{sg}*SQRT({Tt}))", "=B21"),
        ("Vega (per 1 vol point)", f"={S}*EXP(-{q}*{Tt})*{pdf}*SQRT({Tt})/100", "=B22"),
        ("Theta (per calendar day)",
         f"=(-{S}*EXP(-{q}*{Tt})*{pdf}*{sg}/(2*SQRT({Tt}))-{r}*{K}*EXP(-{r}*{Tt})*NORMSDIST(B14)+{q}*{S}*EXP(-{q}*{Tt})*NORMSDIST(B13))/365",
         f"=(-{S}*EXP(-{q}*{Tt})*{pdf}*{sg}/(2*SQRT({Tt}))+{r}*{K}*EXP(-{r}*{Tt})*NORMSDIST(-B14)-{q}*{S}*EXP(-{q}*{Tt})*NORMSDIST(-B13))/365"),
        ("Rho (per 1% rate)", f"={K}*{Tt}*EXP(-{r}*{Tt})*NORMSDIST(B14)/100", f"=-{K}*{Tt}*EXP(-{r}*{Tt})*NORMSDIST(-B14)/100")]
    for i, (lab, cf, pf) in enumerate(greeks):
        rr = 20 + i
        put(ws, f"A{rr}", lab, kind="text")
        put(ws, f"B{rr}", cf, F_D)
        put(ws, f"C{rr}", pf, F_D)

    section(ws, 26, "Market quote (nearest-ATM option, ~1 month out) vs model", 6)
    if o:
        put(ws, "A27", "Expiry", kind="text"); put(ws, "B27", o["expiry"], kind="text", align="right")
        put(ws, "A28", "Market call / put (mid or last)", kind="text")
        put(ws, "B28", round(o["call"] or 0, 4), F_PS); put(ws, "C28", round(o["put"] or 0, 4), F_PS)
        put(ws, "A29", "Market implied vol (call / put)", kind="text")
        put(ws, "B29", round(o["call_iv"] or 0, 4), F_PCT); put(ws, "C29", round(o["put_iv"] or 0, 4), F_PCT)
        put(ws, "A30", "Model minus market", kind="text")
        put(ws, "B30", "=B15-B28", F_PS); put(ws, "C30", "=B16-C28", F_PS)
        put(ws, "A31", "Gap is mostly historical-vs-implied vol. Set volatility above to the market IV to reconcile.", italic=True, kind="text", color="595959")
    else:
        put(ws, "A27", "No option chain available for this listing (common for non-US tickers). Enter a market quote here to compare.", italic=True, kind="text", color="595959")

    section(ws, 33, "Strike ladder", 6)
    header(ws, 34, ["Strike (% of spot)", "Strike", "d1", "d2", "Call", "Put"])
    for i, pct in enumerate((0.8, 0.9, 1.0, 1.1, 1.2)):
        rr = 35 + i
        put(ws, f"A{rr}", pct, '0%', align="left")
        put(ws, f"B{rr}", f"={S}*A{rr}", F_PS)
        put(ws, f"C{rr}", f"=(LN({S}/B{rr})+({r}-{q}+{sg}^2/2)*{Tt})/({sg}*SQRT({Tt}))", F_D)
        put(ws, f"D{rr}", f"=C{rr}-{sg}*SQRT({Tt})", F_D)
        put(ws, f"E{rr}", f"={S}*EXP(-{q}*{Tt})*NORMSDIST(C{rr})-B{rr}*EXP(-{r}*{Tt})*NORMSDIST(D{rr})", F_PS)
        put(ws, f"F{rr}", f"=B{rr}*EXP(-{r}*{Tt})*NORMSDIST(-D{rr})-{S}*EXP(-{q}*{Tt})*NORMSDIST(-C{rr})", F_PS)
    return dict(sigma=f"Options!{sg}")


# ----------------------------------------------------------------------------- Credit
def build_credit(wb, d, lay, A, OR):
    ws = wb["Credit"]
    title(ws, "Credit risk model", "Altman Z-scores, leverage/coverage ratios and a naive Merton distance-to-default. Latest fiscal year. Not valid for banks/insurers.")
    widths(ws, 46, 16, 3)
    ws.column_dimensions["C"].width = 80
    h, LH, T = d["hist"], lay.LH, "ThreeStatement!"
    C = {}

    def row(rr, key, label, val, fmt=None, note=None, bold=False):
        C[key] = f"$B${rr}"
        put(ws, f"A{rr}", label, kind="text", bold=bold)
        put(ws, f"B{rr}", val, fmt, bold=bold, align="right")
        if note:
            put(ws, f"C{rr}", note, italic=True, kind="text", color="595959")

    section(ws, 4, "Inputs (latest fiscal year, $ mm)", 3)
    row(5, "ta", "Total assets", f"={T}{LH}{TS['ta']}", F_NUM)
    row(6, "tl", "Total liabilities", f"={T}{LH}{TS['tl']}", F_NUM)
    row(7, "be", "Book equity", f"={T}{LH}{TS['eq']}", F_NUM)
    row(8, "ca", "Current assets", round(h["cur_assets"][-1], 3), F_NUM, "From balance sheet (0 if not reported).")
    row(9, "cl", "Current liabilities", round(h["cur_liab"][-1], 3), F_NUM)
    row(10, "wc", "Working capital", "=B8-B9", F_NUM)
    row(11, "re", "Retained earnings", round(h["retained"][-1], 3), F_NUM)
    row(12, "ebit", "EBIT", f"={T}{LH}{TS['ebit']}", F_NUM)
    row(13, "ebitda", "EBITDA", f"={T}{LH}{TS['ebitda']}", F_NUM)
    row(14, "int", "Interest expense", f"={T}{LH}{TS['interest']}", F_NUM)
    row(15, "sales", "Sales", f"={T}{LH}{TS['rev']}", F_NUM)
    row(16, "debt", "Total debt", f"={A['debt']}", F_NUM)
    row(17, "cash", "Cash", f"={A['cash']}", F_NUM)
    row(18, "mc", "Market capitalisation", f"={A['mcap']}", F_NUM)

    section(ws, 20, "Altman Z-score (original, public manufacturers)", 3)
    row(21, "x1", "X1 = Working capital / Total assets", "=IFERROR(B10/B5,0)", F_D)
    row(22, "x2", "X2 = Retained earnings / Total assets", "=IFERROR(B11/B5,0)", F_D)
    row(23, "x3", "X3 = EBIT / Total assets", "=IFERROR(B12/B5,0)", F_D)
    row(24, "x4", "X4 = Market cap / Total liabilities", "=IFERROR(B18/B6,0)", F_D)
    row(25, "x5", "X5 = Sales / Total assets", "=IFERROR(B15/B5,0)", F_D)
    row(26, "z", "Z = 1.2 X1 + 1.4 X2 + 3.3 X3 + 0.6 X4 + 1.0 X5", "=1.2*B21+1.4*B22+3.3*B23+0.6*B24+1*B25", '0.00', bold=True)
    row(27, "zz", "Zone (>2.99 safe, <1.81 distress)", '=IF(B26>2.99,"Safe",IF(B26<1.81,"Distress","Grey"))', bold=True)
    row(28, "z2", "Z'' (non-manufacturers/EM) = 6.56 X1 + 3.26 X2 + 6.72 X3 + 1.05 (Book equity/TL)", "=6.56*B21+3.26*B22+6.72*B23+1.05*IFERROR(B7/B6,0)", '0.00', bold=True)
    row(29, "z2z", "Zone (>2.60 safe, <1.10 distress)", '=IF(B28>2.6,"Safe",IF(B28<1.1,"Distress","Grey"))', bold=True)

    section(ws, 31, "Leverage & coverage", 3)
    row(32, "nde", "Net debt / EBITDA", "=IFERROR((B16-B17)/B13,0)", F_X)
    row(33, "cov", "Interest coverage (EBIT / interest)", '=IF(B14>0,B12/B14,"no interest")', F_X)
    row(34, "de", "Debt / equity", "=IFERROR(B16/B7,0)", F_X)
    row(35, "cr", "Current ratio", "=IFERROR(B8/B9,0)", F_X)

    section(ws, 37, "Naive Merton distance-to-default (Bharath-Shumway style)", 3)
    row(38, "E", "Equity value (market cap)", "=B18", F_NUM)
    row(39, "Dbt", "Debt face value (proxy for default point)", "=B16", F_NUM, "Simplification: total debt; many practitioners use ST debt + 0.5 x LT debt.")
    row(40, "sE", "Equity volatility", f"={OR['sigma']}", F_PCT)
    row(41, "sD", "Debt volatility = 5% + 0.25 x equity vol", "=0.05+0.25*B40", F_PCT)
    row(42, "sV", "Asset volatility", "=IFERROR(B38/(B38+B39)*B40+B39/(B38+B39)*B41,0)", F_PCT)
    row(43, "V", "Asset value = E + D", "=B38+B39", F_NUM)
    row(44, "rr", "Drift (risk-free rate)", f"={A['rf']}", F_PCT)
    row(45, "t", "Horizon (years)", 1, '0', "Blue input.")
    row(46, "dd", "Distance to default", '=IF(AND(B39>0,B42>0),(LN(B43/B39)+(B44-B42^2/2)*B45)/(B42*SQRT(B45)),"no debt")', '0.00', bold=True)
    row(47, "pd", "Implied probability of default", '=IF(ISNUMBER(B46),NORMSDIST(-B46),0)', '0.00%', "Model-implied, not a rating. Naive and typically overstates/understates versus agency PDs.", bold=True)
    return dict(z="Credit!$B$26", zone="Credit!$B$27", z2="Credit!$B$28", z2z="Credit!$B$29", dd="Credit!$B$46", pd="Credit!$B$47",
                nde="Credit!$B$32", cov="Credit!$B$33")


# ----------------------------------------------------------------------------- LBO
def build_lbo(wb, d, lay, A, df):
    ws = wb["LBO"]
    title(ws, "Illustrative LBO: 'what could a sponsor pay?'", "A screening model, not a deal model: single tranche mechanics, 100% cash sweep, no fees amortisation, no management rollover.")
    widths(ws, 46, 13, 12)
    L, T = lay, "ThreeStatement!"
    section(ws, 4, "Transaction assumptions", 1 + len(L.cols))
    inp = [
        (5, "prem", "Offer premium to current price", 0.25, F_PCT, True), (6, "offer", "Offer price per share", f"={A['price']}*(1+B5)", F_PS, False),
        (7, "eqp", "Equity purchase price ($ mm)", f"=B6*{A['shares']}", F_NUM, False),
        (8, "nd", "Existing net debt (refinanced)", f"={A['netdebt']}", F_NUM, False),
        (9, "ev", "Entry enterprise value", "=B7+B8", F_NUM, False),
        (10, "ltm", "LTM EBITDA", f"={T}{L.LH}{TS['ebitda']}", F_NUM, False),
        (11, "emult", "Entry EV / EBITDA", "=IFERROR(B9/B10,0)", F_X, False),
        (12, "sx", "Senior debt (x EBITDA)", 3.0, F_X, True), (13, "ux", "Subordinated debt (x EBITDA)", 1.5, F_X, True),
        (14, "sr", "Senior interest rate", 0.085, F_PCT, True), (15, "ur", "Subordinated interest rate", 0.115, F_PCT, True),
        (16, "fee", "Transaction fees (% of EV)", 0.02, F_PCT, True),
        (17, "xm", "Exit EV / EBITDA", df["exit_x"], F_X, True), (18, "irr_t", "Target sponsor IRR", 0.20, F_PCT, True),
        (19, "tax", "Tax rate", f"=Forecast!{L.F}{FC['tax']}", F_PCT, False)]
    for rr, k, lab, v, fmt, key in inp:
        put(ws, f"A{rr}", lab, kind="text")
        put(ws, f"B{rr}", v, fmt, fill=YELLOW if key else None)
    section(ws, 21, "Sources & uses ($ mm)", 1 + len(L.cols))
    su = [(22, "Senior debt", "=B12*B10"), (23, "Subordinated debt", "=B13*B10"), (24, "Sponsor equity (plug)", "=B28-B22-B23"),
          (25, "Total sources", "=B22+B23+B24"), (26, "Purchase of equity + refinance net debt (EV)", "=B9"),
          (27, "Fees", "=B9*B16"), (28, "Total uses", "=B26+B27"), (29, "Sponsor equity % of total", "=IFERROR(B24/B28,0)")]
    for rr, lab, f in su:
        put(ws, f"A{rr}", lab, kind="text", bold=rr in (25, 28))
        put(ws, f"B{rr}", f, F_PCT if rr == 29 else F_NUM, bold=rr in (25, 28))
    year_header(ws, L, 31, "Cash flow & debt schedule ($ mm)")
    rows = {"ebitda": 32, "capex": 33, "dnwc": 34, "tax": 35, "int": 36, "fcf": 37, "sb": 39, "sr_": 40, "se": 41,
            "ub": 42, "ur_": 43, "ue": 44, "cb": 45, "ce": 46, "nd": 47}
    labs = {"ebitda": "EBITDA", "capex": "Less: capex", "dnwc": "Less: increase in NWC", "tax": "Less: cash taxes (on EBIT less interest)",
            "int": "Less: cash interest (on opening balances)", "fcf": "Free cash flow available for debt paydown",
            "sb": "Senior: opening", "sr_": "Senior: repayment", "se": "Senior: closing", "ub": "Sub: opening", "ur_": "Sub: repayment",
            "ue": "Sub: closing", "cb": "Cash: opening", "ce": "Cash: closing", "nd": "Net debt (closing)"}
    for k, rr in rows.items():
        put(ws, f"A{rr}", labs[k], kind="text", bold=k in ("fcf", "nd"))
    R = rows
    for j, c in enumerate(L.PC):
        p = L.prev(c)
        put(ws, f"{c}{R['ebitda']}", f"={T}{c}{TS['ebitda']}", F_NUM)
        put(ws, f"{c}{R['capex']}", f"=-{T}{c}{TS['capex']}", F_NUM)
        put(ws, f"{c}{R['dnwc']}", f"=DCF!{c}{DC['dnwc']}", F_NUM)
        put(ws, f"{c}{R['tax']}", f"=-MAX(0,{T}{c}{TS['ebit']}+{c}{R['int']})*$B$19", F_NUM)
        put(ws, f"{c}{R['int']}", f"=-({c}{R['sb']}*$B$14+{c}{R['ub']}*$B$15)", F_NUM)
        put(ws, f"{c}{R['fcf']}", f"=SUM({c}{R['ebitda']}:{c}{R['int']})", F_NUM, bold=True)
        put(ws, f"{c}{R['sb']}", "=$B$22" if j == 0 else f"={p}{R['se']}", F_NUM)
        put(ws, f"{c}{R['sr_']}", f"=MIN({c}{R['sb']},MAX(0,{c}{R['fcf']}))", F_NUM)
        put(ws, f"{c}{R['se']}", f"={c}{R['sb']}-{c}{R['sr_']}", F_NUM)
        put(ws, f"{c}{R['ub']}", "=$B$23" if j == 0 else f"={p}{R['ue']}", F_NUM)
        put(ws, f"{c}{R['ur_']}", f"=MIN({c}{R['ub']},MAX(0,{c}{R['fcf']}-{c}{R['sr_']}))", F_NUM)
        put(ws, f"{c}{R['ue']}", f"={c}{R['ub']}-{c}{R['ur_']}", F_NUM)
        put(ws, f"{c}{R['cb']}", 0 if j == 0 else f"={p}{R['ce']}", F_NUM, kind="calc")
        put(ws, f"{c}{R['ce']}", f"={c}{R['cb']}+{c}{R['fcf']}-{c}{R['sr_']}-{c}{R['ur_']}", F_NUM)
        put(ws, f"{c}{R['nd']}", f"={c}{R['se']}+{c}{R['ue']}-{c}{R['ce']}", F_NUM, bold=True)
    LP = L.LP
    section(ws, 49, "Returns", 1 + len(L.cols))
    ret = [(50, "hold", "Holding period (years)", f"=DCF!{LP}{DC['n']}", '0'), (51, "xe", "Exit-year EBITDA", f"={LP}{R['ebitda']}", F_NUM),
           (52, "xev", "Exit enterprise value", "=B51*B17", F_NUM), (53, "xnd", "Less: net debt at exit", f"={LP}{R['nd']}", F_NUM),
           (54, "xeq", "Exit equity value", "=B52-B53", F_NUM), (55, "moic", "MOIC", "=IFERROR(B54/B24,0)", '0.00"x"'),
           (56, "irr", "Sponsor IRR", "=IF(B55>0,B55^(1/B50)-1,-1)", F_PCT),
           (58, "maxeq", "Max sponsor equity for target IRR", "=B54/(1+B18)^B50", F_NUM),
           (59, "maxev", "Max entry EV at target IRR (debt held constant)", "=(B22+B23+B58)/(1+B16)", F_NUM),
           (60, "maxps", "Max offer price per share at target IRR", f"=IFERROR((B59-B8)/{A['shares']},0)", F_PS),
           (61, "maxprem", "Implied premium / (discount) to current price", f"=IFERROR(B60/{A['price']}-1,0)", F_PCT)]
    for rr, k, lab, f, fmt in ret:
        put(ws, f"A{rr}", lab, kind="text", bold=k in ("irr", "maxps"))
        put(ws, f"B{rr}", f, fmt, bold=k in ("irr", "maxps"))
    put(ws, "A63", "Row 59-61 answer: 'what is the most a sponsor could pay and still hit the target IRR?' Exit equity is independent of entry price because debt is sized off EBITDA.",
        italic=True, kind="text", color="595959")
    ws.freeze_panes = "B5"
    return dict(maxps="LBO!$B$60", irr="LBO!$B$56", moic="LBO!$B$55")


# ----------------------------------------------------------------------------- Summary & notes
def build_summary(wb, d, lay, A, CR, CRD, LR):
    ws = wb["Summary"]
    title(ws, f"{d['name']} ({d['ticker']}): valuation summary", "All values in reporting currency; per-share values in currency units. Every number links to the model tabs.")
    widths(ws, 44, 16, 14)
    ws.column_dimensions["D"].width = 70
    header(ws, 4, ["Method", "Value per share", "vs current price", "Comment"])
    rows = [("Current share price", f"={A['price']}", None, "Market"),
            ("DCF: perpetuity growth", f"=DCF!$B${DC['ps']}", True, "Driven by Forecast sheet + WACC/g assumptions"),
            ("DCF: exit multiple", f"=DCF!$B${DC['xps']}", True, "Cross-check on terminal value"),
            ("Comps: EV/EBITDA", f"={CR['EV/EBITDA']}", True, "Peer median: review the peer list"),
            ("Comps: EV/Revenue", f"={CR['EV/Revenue']}", True, "Less meaningful when margins differ"),
            ("Comps: P/E", f"={CR['P/E']}", True, "Needs positive earnings"),
            ("LBO: max price at target IRR", f"={LR['maxps']}", True, "Financial-buyer ability-to-pay floor")]
    for i, (lab, f, cmp_, note) in enumerate(rows):
        rr = 5 + i
        put(ws, f"A{rr}", lab, kind="text", bold=i == 0)
        put(ws, f"B{rr}", f, F_PS, bold=i == 0, align="right")
        if cmp_:
            put(ws, f"C{rr}", f'=IFERROR(B{rr}/$B$5-1,"n/a")', F_PCT, align="right")
        put(ws, f"D{rr}", note, italic=True, kind="text", color="595959")

    section(ws, 13, "Key drivers & risk indicators", 4)
    k = [("Active scenario", f"=Forecast!$B${FC['scen_name']}", None), ("WACC", f"={A['wacc']}", F_PCT), ("Terminal growth", f"={A['g']}", F_PCT),
         ("Terminal value % of DCF EV", f"=DCF!$B${DC['tvpct']}", F_PCT),
         ("Revenue CAGR (proj.)", f"=IFERROR((ThreeStatement!{lay.LP}{TS['rev']}/ThreeStatement!{lay.LH}{TS['rev']})^(1/{NP})-1,0)", F_PCT),
         ("Altman Z / zone", f'=TEXT({CRD["z"]},"0.00")&" / "&{CRD["zone"]}', None), ("Altman Z'' / zone", f'=TEXT({CRD["z2"]},"0.00")&" / "&{CRD["z2z"]}', None),
         ("Net debt / EBITDA", f"={CRD['nde']}", F_X), ("Merton distance-to-default", f"={CRD['dd']}", '0.00'),
         ("Merton implied PD (1 yr)", f"={CRD['pd']}", '0.00%'), ("LBO sponsor IRR at offer premium", f"={LR['irr']}", F_PCT)]
    for i, (lab, f, fmt) in enumerate(k):
        put(ws, f"A{14 + i}", lab, kind="text")
        put(ws, f"B{14 + i}", f, fmt, align="right")
    rr = 14 + len(k) + 1
    section(ws, rr, "Model integrity checks", 4)
    chk = TS["check"]
    put(ws, f"A{rr + 1}", "Balance sheet balances (all years)", kind="text")
    put(ws, f"B{rr + 1}", f'=IF(SUMPRODUCT(ABS(ThreeStatement!{lay.HC[0]}{chk}:{lay.LP}{chk}))<0.01,"OK","ERROR")', align="right")
    put(ws, f"A{rr + 2}", "DCF: WACC > terminal growth", kind="text")
    put(ws, f"B{rr + 2}", f"=DCF!$B${DC['chk']}", align="right")
    put(ws, f"A{rr + 3}", "Option put-call parity", kind="text")
    put(ws, f"B{rr + 3}", '=IF(ABS(Options!$B$17)<0.0001,"OK","ERROR")', align="right")
    rr += 5
    section(ws, rr, "Warnings from the data layer", 4)
    for i, w in enumerate(d["warnings"] or ["None"]):
        put(ws, f"A{rr + 1 + i}", w, kind="text", color="C00000")

    ch = BarChart()
    ch.type, ch.style, ch.title = "col", 10, "Value per share by method"
    ch.add_data(Reference(ws, min_col=2, min_row=5, max_row=11), titles_from_data=False)
    ch.set_categories(Reference(ws, min_col=1, min_row=5, max_row=11))
    ch.legend, ch.height, ch.width = None, 8.5, 18
    ws.add_chart(ch, "F4")
    ws.freeze_panes = "A5"


def build_notes(wb, d):
    ws = wb["Notes"]
    title(ws, "Notes, method & limitations")
    ws.column_dimensions["A"].width = 150
    lines = [
        "COLOUR CODE: blue = hard-coded input; black = formula; green = link to another sheet; yellow fill = key assumption to review.",
        f"DATA: {d['source']}, pulled {d['asof']}. Last {len(d['years'])} fiscal years of statements; TTM/market data for comps.",
        "",
        "MODELS INCLUDED (ticker-only): Three-statement, Forecast (drivers + scenarios), DCF (perpetuity + exit multiple + sensitivity), Comps, Option pricing (BSM), Credit risk (Altman, Merton), LBO (illustrative).",
        "NOT INCLUDED (need extra inputs): Precedent transactions (deal database), SOTP (segment data), M&A (second company + synergies), IPO, Budget, Consolidation, Project finance.",
        "",
        "KEY LIMITATIONS",
        "- Default forecasts are mechanical (historical CAGR fading to terminal growth; historical average margins). They are a starting point, not a view.",
        "- Risk-free rate and equity risk premium defaults are placeholders: update them to today's values for your market.",
        "- Banks, insurers and REITs need different models; the DCF/Altman/EV multiples here are not appropriate for them.",
        "- Historical 'other assets/liabilities' and 'other income' are plugs so the historical statements tie to reported totals.",
        "- Comps peers are auto-selected by industry and may be irrelevant: edit them. Mixed-currency peers are fine for multiples but not for size comparisons.",
        "- LBO ignores fee amortisation, minimum cash, tranche-specific covenants, and management rollover. Treat as an ability-to-pay screen.",
        "- Option pricing uses historical volatility by default; real option prices depend on implied vol. American exercise is ignored.",
        "",
        "This workbook is a modelling aid, not investment advice or a research recommendation. Verify inputs against filings before relying on it.",
    ]
    for i, t in enumerate(lines):
        put(ws, f"A{3 + i}", t, kind="text", bold=t.isupper() or t.startswith("KEY"))


# ----------------------------------------------------------------------------- entry point
def build_workbook(d, path):
    lay, df = Lay(d), defaults(d)
    wb = Workbook()
    wb.remove(wb.active)
    for n in ("Summary", "Assumptions", "Forecast", "ThreeStatement", "Reported Financials", "DCF", "Comps", "Options", "Credit", "LBO", "Notes"):
        wb.create_sheet(n)
    A = build_assumptions(wb, d, lay, df)
    build_three_statement(wb, d, lay)
    build_reported_financials(wb, d, lay)
    build_forecast(wb, d, lay, df, A)
    build_dcf(wb, lay, A)
    CR = build_comps(wb, d, A)
    OR = build_options(wb, d, A)
    CRD = build_credit(wb, d, lay, A, OR)
    LR = build_lbo(wb, d, lay, A, df)
    build_summary(wb, d, lay, A, CR, CRD, LR)
    build_notes(wb, d)
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)
    return path
