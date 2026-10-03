# Youth Consumer Spend Index

A custom stock index tracking the companies India's youth actually spend on, built in Python with Streamlit.

The index tracks a capped market-cap basket of youth-facing consumer brands (Zomato/Eternal, Paytm, Nykaa, Delhivery, Titan, ABFRL) and benchmarks it against Nifty 50 and Sensex. The app includes TradingView-styled charts, extended risk analysis (drawdown, beta, correlation, rolling volatility), search-and-add for custom constituents, and a one-click Word report generator.

## Run on the web

The hosted version runs on Streamlit Community Cloud. See the deployment section below.

## Run locally

```bash
python -m venv venv
source venv/bin/activate         # Mac/Linux
# venv\Scripts\activate          # Windows

pip install -r requirements.txt
streamlit run app.py
```

Opens at http://localhost:8501.

## Package as a desktop app

The repo includes a PyInstaller spec file (`YouthIndex.spec`) and a launcher (`run_app.py`) that package the app as a double-clickable `.app` on Mac or `.exe` folder on Windows. From an activated venv with `pyinstaller` installed:

```bash
pyinstaller YouthIndex.spec --clean --noconfirm
```

The packaged version persists user customisations to `~/.youth_index/user_config.json`. The hosted web version runs session-only (changes reset on reload) because the server filesystem is shared across visitors.

## Project structure

- `app.py` — the Streamlit app (UI, index math, charts, report generator)
- `run_app.py` — launcher used by PyInstaller for the packaged version
- `index_core.py` — terminal-only version of the index math, useful for debugging
- `YouthIndex.spec` — PyInstaller build configuration
- `requirements.txt` — Python dependencies (also used by Streamlit Cloud)

## Honest caveats

The index has historically outperformed Nifty by a wide margin, but with roughly twice the volatility and material drawdowns. The start date materially affects the result (several constituents IPO'd into weak markets in 2021-2022). Treat the numbers as a demonstration of thematic construction, not as investment advice.
