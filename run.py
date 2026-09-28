"""Usage:  python run.py AAPL            (US)
          python run.py RELIANCE.NS     (India NSE; .BO for BSE)
          python run.py MSFT --peers GOOGL,AMZN,ORCL,SAP
          python run.py --demo          (synthetic data, works offline)"""
import argparse
import sys

from data_source import demo_company, fetch_company
from workbook import build_workbook


def main():
    ap = argparse.ArgumentParser(description="Ticker -> valuation workbook")
    ap.add_argument("ticker", nargs="?", default=None)
    ap.add_argument("--peers", help="comma-separated peer tickers (overrides auto-selection)")
    ap.add_argument("--out", help="output .xlsx path")
    ap.add_argument("--demo", action="store_true", help="use synthetic data (offline)")
    a = ap.parse_args()
    if not a.demo and not a.ticker:
        ap.error("give a ticker or use --demo")
    try:
        d = demo_company() if a.demo else fetch_company(a.ticker.upper(), peers=a.peers.split(",") if a.peers else None)
    except Exception as e:
        sys.exit(f"Data error: {e}")
    out = a.out or f"{d['ticker'].replace('.', '_')}_models.xlsx"
    build_workbook(d, out)
    print(f"Wrote {out}")
    for w in d["warnings"]:
        print("WARNING:", w)


if __name__ == "__main__":
    main()
