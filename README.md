# ticker2models

Type a ticker, get an Excel workbook with live formulas: Three-Statement, Forecast, DCF, Comps, Option Pricing, Credit Risk and an illustrative LBO.

## Run the website
    pip install -r requirements.txt
    python app.py                     # open http://127.0.0.1:8000
    # or: docker build -t t2m . && docker run -p 8000:8000 t2m

Results render in the browser; the 'Download Excel' button gives the same model with live formulas.

## Run from the command line (Excel only)
    pip install -r requirements.txt
    python run.py AAPL                  # US
    python run.py RELIANCE.NS           # India (NSE); use .BO for BSE
    python run.py MSFT --peers GOOGL,AMZN,ORCL,SAP
    python run.py --demo                # synthetic data, works offline

## Files
- data_source.py : fetch + normalise (yfinance). Replace this to change data vendor; keep the dict shape of demo_company().
- workbook.py    : all model logic, written as Excel formulas (Python never does the valuation maths).
- run.py         : CLI (Excel only).
- app.py         : Flask server: builds workbook -> pure-Python "formulas" library calculates -> values shown in browser; same file is the download.
- static/index.html : the page.

## Sheets
Summary, Assumptions, Forecast, ThreeStatement, DCF, Comps, Options, Credit, LBO, Notes.
Blue = input, black = formula, green = cross-sheet link, yellow = key assumption. Change the scenario on Assumptions (1/2/3).

## Known limits
- yfinance is unofficial and free-tier; fine for prototyping, not for a commercial product (licence a data vendor first).
- Risk-free rate / ERP defaults are placeholders. Update them.
- Banks, insurers and REITs get a warning: these models do not suit them.
- Only ~3-4 years of history from the free feed; forecasts are mechanical defaults, not views.
- Formulas have no cached values until opened in Excel/LibreOffice (they calculate on open).
