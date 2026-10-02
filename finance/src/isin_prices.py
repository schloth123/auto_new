"""Write latest ISIN prices converted to euros to a CSV file."""

from __future__ import annotations

import argparse
import csv
import math
from pathlib import Path
from typing import Any

import yfinance as yf
import yaml


FINANCE_DIR = Path(__file__).resolve().parents[1]
OUTPUT_PATH = FINANCE_DIR / "workspace" / "isin_prices.csv"
CSV_COLUMNS = ("isin", "name", "current rate", "currency")
SUPPORTED_QUOTE_TYPES = {"EQUITY", "ETF", "MUTUALFUND"}


def read_isins(input_path: Path) -> list[str]:
    text = input_path.read_text(encoding="utf-8-sig")
    loaded = yaml.safe_load(text)

    if isinstance(loaded, list) and all(isinstance(item, str) for item in loaded):
        isins = [item.strip().upper() for item in loaded if item.strip()]
    else:
        isins = [line.split("#", 1)[0].strip().upper() for line in text.splitlines()]
        isins = [isin for isin in isins if isin]

    if not isins:
        raise ValueError("Die Eingabedatei enthält keine ISINs.")
    if any(len(isin) != 12 or not isin.isalnum() for isin in isins):
        raise ValueError("Jede ISIN muss aus genau 12 Buchstaben oder Ziffern bestehen.")
    return isins


def latest_close(ticker: yf.Ticker) -> float:
    history = ticker.history(period="5d", interval="1d", auto_adjust=False, actions=False)
    closes = history["Close"].dropna()
    if closes.empty:
        raise ValueError(f"Kein aktueller Schlusskurs für {ticker.ticker} verfügbar.")

    rate = float(closes.iloc[-1])
    if not math.isfinite(rate) or rate <= 0:
        raise ValueError(f"Ungültiger Schlusskurs für {ticker.ticker}: {rate}.")
    return rate


def rate_in_eur(rate: float, currency: str) -> float:
    divisor = 100 if currency in {"GBp", "GBX", "GBPENCE"} else 1
    currency = "GBP" if divisor == 100 else currency.upper()
    if currency in {"EUR", "EURO"}:
        return rate

    direct = yf.Ticker(f"{currency}EUR=X")
    try:
        fx_rate = latest_close(direct)
        return rate / divisor * fx_rate
    except (ValueError, KeyError, TypeError):
        inverse = yf.Ticker(f"EUR{currency}=X")
        fx_rate = latest_close(inverse)
        return rate / divisor / fx_rate


def fetch_isin(isin: str) -> dict[str, Any]:
    search = yf.Search(isin, max_results=10, news_count=0)
    quotes = [quote for quote in search.quotes if quote.get("quoteType") in SUPPORTED_QUOTE_TYPES]
    if not quotes:
        raise ValueError(f"Yahoo Finance hat kein handelbares Wertpapier für {isin} gefunden.")

    exact_matches = [quote for quote in quotes if quote.get("isin", "").upper() == isin]
    quote = (exact_matches or quotes)[0]
    symbol = quote.get("symbol")
    if not symbol:
        raise ValueError(f"Für {isin} wurde kein Börsensymbol gefunden.")

    ticker = yf.Ticker(symbol)
    details = ticker.get_info()
    reported_isin = details.get("isin")
    if reported_isin and reported_isin.upper() != isin:
        raise ValueError(f"ISIN-Abweichung für {symbol}: erwartet {isin}, erhalten {reported_isin}.")

    currency = details.get("currency")
    if not isinstance(currency, str) or not currency.strip():
        raise ValueError(f"Keine Währung für {isin} ({symbol}) verfügbar; Export wird abgebrochen.")

    source_rate = latest_close(ticker)
    euro_rate = rate_in_eur(source_rate, currency)
    name = details.get("longName") or details.get("shortName") or quote.get("longname") or quote.get("shortname")
    if not name:
        raise ValueError(f"Kein Name für {isin} ({symbol}) verfügbar.")

    return {"isin": isin, "name": name, "current rate": f"{euro_rate:.8f}", "currency": "EUR"}


def create_csv(input_path: Path) -> Path:
    isins = read_isins(input_path)
    rows = []
    errors = []
    for isin in isins:
        try:
            rows.append(fetch_isin(isin))
        except Exception as error:
            errors.append(f"{isin}: {error}")

    if errors:
        raise ValueError("Kursdaten konnten nicht zuverlässig ermittelt werden:\n" + "\n".join(errors))

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with OUTPUT_PATH.open("w", encoding="utf-8-sig", newline="") as output_file:
        writer = csv.DictWriter(output_file, fieldnames=CSV_COLUMNS)
        writer.writeheader()
        writer.writerows(rows)
    return OUTPUT_PATH


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("yaml_path", type=Path, help="Pfad zur Eingabedatei mit ISINs")
    args = parser.parse_args()
    if not args.yaml_path.is_file():
        parser.error(f"Datei nicht gefunden: {args.yaml_path}")

    try:
        output_path = create_csv(args.yaml_path)
    except (OSError, ValueError, yaml.YAMLError) as error:
        parser.error(str(error))
    print(f"CSV geschrieben nach {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())