"""Data layer: pulls statements + market data for one ticker and normalises them.

All money values are returned in MILLIONS of reporting currency.
Swap this module out (e.g. for SEC EDGAR, FMP, or an NSE/BSE feed) without touching workbook.py:
just return a dict with the same shape as demo_company().
"""
from __future__ import annotations

import datetime as dt
import math

MM = 1e6


# ----------------------------------------------------------------------------- helpers
def _num(x):
    try:
        if x is None:
            return None
        x = float(x)
        return None if math.isnan(x) or math.isinf(x) else x
    except (TypeError, ValueError):
        return None


def _series(df, names, years, scale=MM):
    """First matching row in a yfinance statement, aligned to calendar `years` (oldest -> newest)."""
    out = [None] * len(years)
    if df is None or getattr(df, "empty", True):
        return out
    for name in names:
        if name in df.index:
            by_year = {c.year: _num(v) for c, v in df.loc[name].items()}
            vals = [by_year.get(y) for y in years]
            if any(v is not None for v in vals):
                return [None if v is None else v / scale for v in vals]
    return out


def _fill(xs, default=0.0):
    return [default if v is None else v for v in xs]


def _need(name, xs):
    if any(v is None for v in xs):
        raise ValueError(f"Missing required line item '{name}' for one or more years "
                         f"(got {xs}). This ticker's statements are too sparse for the model.")
    return xs


# ----------------------------------------------------------------------------- market data
def _hist_vol(t):
    try:
        px = t.history(period="1y")["Close"].dropna()
        r = (px / px.shift(1)).apply(math.log).dropna()
        if len(r) > 60:
            return float(r.std() * math.sqrt(252))
    except Exception:
        pass
    return None


def _option_snapshot(t, price):
    """Nearest ATM call/put ~1 month out. Returns None if the listing has no options data."""
    try:
        today = dt.date.today()
        chosen = None
        for e in t.options:
            days = (dt.date.fromisoformat(e) - today).days
            if days >= 25:
                chosen = (e, days)
                break
        if not chosen:
            return None
        ch = t.option_chain(chosen[0])
        strike = float(ch.calls.iloc[(ch.calls["strike"] - price).abs().argsort()[:1]]["strike"].iloc[0])
        c = ch.calls[ch.calls["strike"] == strike].iloc[0]
        p = ch.puts[ch.puts["strike"] == strike].iloc[0]

        def px(row):
            bid, ask = _num(row.get("bid")), _num(row.get("ask"))
            return (bid + ask) / 2 if bid and ask else _num(row.get("lastPrice"))

        return dict(expiry=chosen[0], days=chosen[1], strike=strike, call=px(c), put=px(p),
                    call_iv=_num(c.get("impliedVolatility")), put_iv=_num(p.get("impliedVolatility")))
    except Exception:
        return None


def _snapshot(info):
    """Comps row from a yfinance info dict (TTM figures, millions)."""
    def g(k):
        v = _num(info.get(k))
        return 0.0 if v is None else v / MM
    return dict(name=info.get("shortName") or info.get("longName") or "", ticker=info.get("symbol", ""),
                mcap=g("marketCap"), debt=g("totalDebt"), cash=g("totalCash"),
                rev=g("totalRevenue"), ebitda=g("ebitda"), ni=g("netIncomeToCommon"))


def _find_peers(yf, info, self_symbol, n):
    syms = []
    for kind, key in (("Industry", info.get("industryKey")), ("Sector", info.get("sectorKey"))):
        if not key or len(syms) >= n:
            continue
        try:
            top = getattr(yf, kind)(key).top_companies
            syms += [s for s in top.index.tolist() if s != self_symbol and s not in syms]
        except Exception:
            continue
    return syms[:n]


# ----------------------------------------------------------------------------- main entry
def fetch_company(ticker, peers=None, n_peers=8, max_hist=3):
    import yfinance as yf

    t = yf.Ticker(ticker)
    info = t.get_info()
    inc, bal, cf = t.income_stmt, t.balance_sheet, t.cashflow
    if inc is None or inc.empty:
        raise ValueError(f"No financial statements found for '{ticker}'. Check the symbol "
                         f"(Indian stocks need a suffix, e.g. RELIANCE.NS or TCS.BO).")

    years = sorted({c.year for c in inc.columns})[-max_hist:]
    if len(years) < 2:
        raise ValueError("Need at least 2 years of statements.")

    ebit = _series(inc, ["EBIT", "Operating Income"], years)
    da = _series(cf, ["Depreciation And Amortization", "Depreciation Amortization Depletion"], years)
    if all(v is None for v in da):
        da = _series(inc, ["Reconciled Depreciation"], years)
    ebitda = _series(inc, ["EBITDA", "Normalized EBITDA"], years)
    ebitda = [e if e is not None else (b + a if b is not None and a is not None else None)
              for e, b, a in zip(ebitda, ebit, da)]
    da = [a if a is not None else (e - b if e is not None and b is not None else None)
          for a, e, b in zip(da, ebitda, ebit)]

    h = dict(
        rev=_need("Total Revenue", _series(inc, ["Total Revenue", "Operating Revenue"], years)),
        ebitda=_need("EBITDA", ebitda),
        da=_need("Depreciation", da),
        interest=[abs(v) for v in _fill(_series(inc, ["Interest Expense", "Interest Expense Non Operating"], years))],
        tax=_fill(_series(inc, ["Tax Provision"], years)),
        ni=_need("Net Income", _series(inc, ["Net Income", "Net Income Common Stockholders"], years)),
        capex=[abs(v) for v in _fill(_series(cf, ["Capital Expenditure"], years))],
        div=[abs(v) for v in _fill(_series(cf, ["Cash Dividends Paid", "Common Stock Dividend Paid"], years))],
        cash=_fill(_series(bal, ["Cash Cash Equivalents And Short Term Investments", "Cash And Cash Equivalents"], years)),
        ar=_fill(_series(bal, ["Accounts Receivable", "Receivables"], years)),
        inv=_fill(_series(bal, ["Inventory"], years)),
        ppe=_fill(_series(bal, ["Net PPE"], years)),
        ap=_fill(_series(bal, ["Accounts Payable", "Payables"], years)),
        debt=_fill(_series(bal, ["Total Debt"], years)),
        equity=_need("Stockholders Equity", _series(bal, ["Stockholders Equity", "Common Stock Equity"], years)),
        total_assets=_need("Total Assets", _series(bal, ["Total Assets"], years)),
        retained=_fill(_series(bal, ["Retained Earnings"], years)),
        cur_assets=_fill(_series(bal, ["Current Assets"], years)),
        cur_liab=_fill(_series(bal, ["Current Liabilities"], years)),
    )

    price = _num(info.get("currentPrice")) or _num(info.get("regularMarketPrice")) or _num(info.get("previousClose"))
    if not price:
        try:
            price = _num(t.fast_info.get("lastPrice") or t.fast_info.get("last_price"))
        except Exception:
            pass
    if not price:
        try:
            closes = t.history(period="5d")["Close"].dropna()
            price = _num(closes.iloc[-1]) if len(closes) else None
        except Exception:
            pass
    if not price:
        raise ValueError(f"Could not read a current share price for '{ticker}'. "
                         f"Yahoo Finance may be rate-limiting this server, or the symbol may be wrong "
                         f"(Indian stocks need .NS or .BO). Try again in a minute or try a different ticker.")
    mcap = _num(info.get("marketCap"))
    shares_raw = _num(info.get("sharesOutstanding")) or _num(info.get("impliedSharesOutstanding"))
    if not shares_raw:
        try:
            fi = t.fast_info
            shares_raw = _num(fi.get("shares") or fi.get("shares_outstanding"))
            mcap = mcap or _num(fi.get("market_cap") or fi.get("marketCap"))
        except Exception:
            pass
    if not shares_raw and mcap:
        shares_raw = mcap / price
    shares = (shares_raw or 0) / MM
    if not shares:
        raise ValueError(f"Could not determine shares outstanding for '{ticker}'. "
                         f"Yahoo Finance may be withholding this field for this listing or rate-limiting "
                         f"this server. Try again shortly, or try a different ticker.")

    warnings = []
    cur, fcur = info.get("currency", "USD"), info.get("financialCurrency")
    if fcur and fcur != cur:
        warnings.append(f"Trading currency ({cur}) differs from reporting currency ({fcur}); per-share values "
                        f"will be wrong until you convert. Model is in {fcur}.")
    if info.get("sector") in ("Financial Services", "Real Estate"):
        warnings.append(f"Sector '{info.get('sector')}': a standard FCFF DCF, Altman Z and EV multiples are not "
                        f"appropriate for banks/insurers/REITs. Treat outputs as illustrative only.")
    vol = _hist_vol(t)
    if vol is None:
        vol = 0.30
        warnings.append("Could not compute historical volatility; defaulted to 30%.")
    if info.get("beta") is None:
        warnings.append("No beta available; defaulted to 1.0.")

    peer_syms = [p.strip() for p in peers] if peers else _find_peers(yf, info, ticker, n_peers)
    peer_rows = []
    for s in peer_syms:
        try:
            row = _snapshot(yf.Ticker(s).get_info())
            row["ticker"] = row["ticker"] or s
            if row["mcap"] > 0:
                peer_rows.append(row)
        except Exception:
            continue
    if not peer_rows:
        warnings.append("No peers found automatically. Re-run with --peers A,B,C or fill the yellow rows in Comps.")

    div_rate = _num(info.get("dividendRate")) or 0.0
    return dict(
        ticker=ticker, name=info.get("longName") or info.get("shortName") or ticker, currency=cur,
        sector=info.get("sector", ""), industry=info.get("industry", ""), price=price, shares=shares,
        beta=_num(info.get("beta")) or 1.0, vol=vol, div_yield=div_rate / price if price else 0.0,
        years=years, hist=h, target_ttm=_snapshot(info), peers=peer_rows,
        option=_option_snapshot(t, price), warnings=warnings,
        source="Yahoo Finance via yfinance (unofficial; personal use only)", asof=dt.date.today().isoformat(),
    )


# ----------------------------------------------------------------------------- synthetic demo
def demo_company():
    """Entirely synthetic company so the workbook can be built/tested offline. NOT real data."""
    h = dict(
        rev=[4200, 4650, 5100], ebitda=[840, 930, 1020], da=[210, 225, 240], interest=[60, 58, 55],
        tax=[95, 110, 125], ni=[470, 540, 605], capex=[260, 280, 300], div=[100, 120, 140],
        cash=[400, 450, 520], ar=[560, 620, 680], inv=[420, 460, 500], ppe=[1800, 1900, 2000],
        ap=[380, 420, 460], debt=[900, 850, 800], equity=[2100, 2520, 2990],
        total_assets=[3600, 3950, 4350], retained=[1200, 1550, 1980],
        cur_assets=[1500, 1650, 1850], cur_liab=[700, 750, 800],
    )
    def peer(n, s, mc, d, c, r, e, ni):
        return dict(name=n, ticker=s, mcap=mc, debt=d, cash=c, rev=r, ebitda=e, ni=ni)
    peers = [peer("Alpha Industries", "ALP", 21000, 1500, 600, 6800, 1350, 780),
             peer("Beta Systems", "BET", 9800, 700, 450, 3900, 640, 380),
             peer("Gamma Corp", "GAM", 30500, 3200, 900, 9100, 1900, 1010),
             peer("Delta Holdings", "DEL", 12400, 2100, 300, 5200, 780, -40),
             peer("Epsilon Ltd", "EPS", 7300, 400, 500, 2900, 470, 290),
             peer("Zeta Group", "ZET", 15900, 1200, 700, 5600, 1010, 610)]
    return dict(
        ticker="DEMO", name="Demo Manufacturing Co (SYNTHETIC)", currency="USD", sector="Industrials",
        industry="Specialty Industrial Machinery", price=150.0, shares=100.0, beta=1.1, vol=0.28,
        div_yield=0.008, years=[2023, 2024, 2025], hist=h,
        target_ttm=dict(name="Demo", ticker="DEMO", mcap=15000, debt=800, cash=520, rev=5250, ebitda=1060, ni=630),
        peers=peers,
        option=dict(expiry="synthetic", days=32, strike=150.0, call=6.9, put=5.8, call_iv=0.27, put_iv=0.28),
        warnings=["SYNTHETIC DEMO DATA - not a real company."],
        source="Synthetic", asof=dt.date.today().isoformat(),
    )
