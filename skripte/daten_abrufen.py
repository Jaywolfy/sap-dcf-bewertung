"""Lädt alle Eingangsdaten für das Modell und schreibt sie nach daten/sap_kennzahlen.json.

Quellen:
  Yahoo Finance   Jahresabschlusszahlen der SAP SE (IFRS, Mio. EUR), Schlusskurse von SAP und DAX
  EZB             Rendite zehnjähriger Staatsanleihen des Euroraums mit AAA-Rating (risikofreier Zins)

Das Beta wird hier selbst berechnet: Kovarianz der wöchentlichen Renditen von SAP und DAX geteilt durch die
Varianz der DAX-Renditen, über fünf Jahre.
"""

import csv
import gzip
import io
import json
import math
import time
import urllib.request
from datetime import UTC, date, datetime
from pathlib import Path

OUT = Path(__file__).resolve().parents[1] / "daten" / "sap_kennzahlen.json"
HEADERS = {"User-Agent": "Mozilla/5.0"}
# Aus den Geschäftsberichten von SAP (Dateien sap-investor-YYYY-annual-report.pdf bzw. sap-YYYY-integrated-report.pdf
# unter sap.com/docs/download/investors/YYYY/). Je Jahr zählt der Wert aus dem Bericht des Folgejahres, falls der
# Vorjahreswert dort angepasst wurde. Yahoo liefert nur die letzten vier Jahre.
# Umsatz in Mio. EUR, Gesamtumsatz nach IFRS. 1997 bis 1999 fehlen: die Berichte nennen DM und keine eindeutige Summe.
UMSATZ_2000_2021 = {
    "2000": 6265, "2001": 7341, "2002": 7413, "2003": 7025, "2004": 7514, "2005": 8513, "2006": 9393,
    "2007": 10242, "2008": 11575, "2009": 10672, "2010": 12464, "2011": 14233, "2012": 16223, "2013": 16815,
    "2014": 17560, "2015": 20793, "2016": 22062, "2017": 23461, "2018": 24708, "2019": 27553, "2020": 27338,
    "2021": 27842,
}
# Betriebsergebnis nach IFRS in Mio. EUR, 2011 bis 2021 aus den Berichten. Vor 2011 ist keine einheitliche
# IFRS-Reihe in den Berichten, die ich geprüft habe.
BETRIEBSERGEBNIS_IFRS_2011_2021 = {
    "2011": 4881, "2012": 4041, "2013": 4479, "2014": 4331, "2015": 4252, "2016": 5135, "2017": 4877,
    "2018": 5703, "2019": 4473, "2020": 6623, "2021": 4656,
}
FIELDS = {
    "umsatz": "annualTotalRevenue",
    "betriebsergebnis": "annualOperatingIncome",
    "ergebnis_vor_steuern": "annualPretaxIncome",
    "steuern": "annualTaxProvision",
    "jahresueberschuss": "annualNetIncome",
    "abschreibungen": "annualReconciledDepreciation",
    "investitionen": "annualCapitalExpenditure",
    "operativer_cashflow": "annualOperatingCashFlow",
    "free_cashflow": "annualFreeCashFlow",
    "aktienbasierte_verguetung": "annualStockBasedCompensation",
    "finanzschulden": "annualTotalDebt",
    "liquide_mittel": "annualCashCashEquivalentsAndShortTermInvestments",
    "aktien_ausstehend": "annualOrdinarySharesNumber",
}


def get(url: str) -> bytes:
    with urllib.request.urlopen(urllib.request.Request(url, headers=HEADERS), timeout=30) as response:
        body = response.read()
    # Die EZB liefert manche Antworten gzip-komprimiert, auch ohne Accept-Encoding-Header
    return gzip.decompress(body) if body[:2] == b"\x1f\x8b" else body


def fundamentals(symbol: str) -> dict[str, dict[str, float]]:
    """Jahreswerte je Kennzahl in Mio. EUR, zum Beispiel {"umsatz": {"2025": 36800.0}}."""
    url = (f"https://query1.finance.yahoo.com/ws/fundamentals-timeseries/v1/finance/timeseries/{symbol}"
           f"?symbol={symbol}&type={','.join(FIELDS.values())}&period1=1577836800&period2={int(time.time())}")
    result = json.loads(get(url))["timeseries"]["result"]
    by_type = {entry["meta"]["type"][0]: entry for entry in result}
    table = {}
    for name, key in FIELDS.items():
        rows = by_type.get(key, {}).get(key) or []
        table[name] = {row["asOfDate"][:4]: round(row["reportedValue"]["raw"] / 1e6, 1) for row in rows if row}
    return table


def weekly_closes(symbol: str) -> dict[date, float]:
    url = f"https://query1.finance.yahoo.com/v8/finance/chart/{symbol}?range=5y&interval=1wk"
    result = json.loads(get(url))["chart"]["result"][0]
    closes = result["indicators"]["quote"][0]["close"]
    return {datetime.fromtimestamp(stamp, tz=UTC).date(): close
            for stamp, close in zip(result["timestamp"], closes, strict=True) if close}


def monthly_closes(symbol: str) -> list[list]:
    """Monatsschlusskurse seit dem frühesten Kurs, den Yahoo führt (für SAP ist das 1998)."""
    url = f"https://query1.finance.yahoo.com/v8/finance/chart/{symbol}?range=max&interval=1mo"
    result = json.loads(get(url))["chart"]["result"][0]
    closes = result["indicators"]["quote"][0]["close"]
    return [[datetime.fromtimestamp(stamp, tz=UTC).date().isoformat(), round(close, 2)]
            for stamp, close in zip(result["timestamp"], closes, strict=True) if close]


def last_close(symbol: str) -> tuple[str, float]:
    url = f"https://query1.finance.yahoo.com/v8/finance/chart/{symbol}?range=10d&interval=1d"
    result = json.loads(get(url))["chart"]["result"][0]
    pairs = [(stamp, close) for stamp, close in zip(result["timestamp"], result["indicators"]["quote"][0]["close"], strict=True) if close]
    stamp, close = pairs[-1]
    return datetime.fromtimestamp(stamp + result["meta"]["gmtoffset"], tz=UTC).date().isoformat(), round(close, 2)


def beta(stock: dict[date, float], market: dict[date, float]) -> tuple[float, int]:
    days = sorted(set(stock) & set(market))
    r_stock = [math.log(stock[b] / stock[a]) for a, b in zip(days, days[1:], strict=False)]
    r_market = [math.log(market[b] / market[a]) for a, b in zip(days, days[1:], strict=False)]
    mean_s, mean_m = sum(r_stock) / len(r_stock), sum(r_market) / len(r_market)
    covariance = sum((s - mean_s) * (m - mean_m) for s, m in zip(r_stock, r_market, strict=True))
    variance = sum((m - mean_m) ** 2 for m in r_market)
    return round(covariance / variance, 2), len(r_stock)


def risk_free() -> tuple[str, float]:
    """Kassazins zehnjähriger AAA-Staatsanleihen des Euroraums aus der Zinsstrukturkurve der EZB."""
    url = "https://data-api.ecb.europa.eu/service/data/YC/B.U2.EUR.4F.G_N_A.SV_C_YM.SR_10Y?lastNObservations=1&format=csvdata"
    row = next(csv.DictReader(io.StringIO(get(url).decode("utf-8-sig"))))
    return row["TIME_PERIOD"], round(float(row["OBS_VALUE"]) / 100, 4)


def main() -> None:
    price_day, price = last_close("SAP.DE")
    sap_weekly, dax_weekly = weekly_closes("SAP.DE"), weekly_closes("^GDAXI")
    beta_value, weeks = beta(sap_weekly, dax_weekly)
    rf_day, rf = risk_free()
    data = {
        "abgerufen_am": date.today().isoformat(),
        "unternehmen": "SAP SE",
        "einheit": "Mio. EUR",
        "abschluss": fundamentals("SAP.DE"),
        "kurs": {"tag": price_day, "schluss_eur": price},
        "kurs_monate": monthly_closes("SAP.DE"),
        "umsatz_lang": {
            "einheit": "Mio. EUR",
            "quelle": "2000 bis 2021: Geschäftsberichte der SAP (IFRS), manuell übernommen, nicht über eine API abgerufen. "
                      "2022 bis 2025: Yahoo Finance (siehe Blatt Historie).",
            "werte": UMSATZ_2000_2021,
        },
        "betriebsergebnis_lang": {
            "einheit": "Mio. EUR",
            "quelle": "2011 bis 2021: Geschäftsberichte der SAP (IFRS), manuell übernommen. "
                      "2022 bis 2025: Yahoo Finance (siehe Blatt Historie).",
            "werte": BETRIEBSERGEBNIS_IFRS_2011_2021,
        },
        "kurs_wochen": [{"datum": day.isoformat(), "sap": round(sap_weekly[day], 2), "dax": round(dax_weekly[day], 2)}
                        for day in sorted(set(sap_weekly) & set(dax_weekly))],
        "beta": {"wert": beta_value, "wochen": weeks, "gegen": "DAX, wöchentliche Renditen, 5 Jahre"},
        "risikofreier_zins": {"tag": rf_day, "wert": rf, "quelle": "EZB, Zinsstrukturkurve AAA-Staatsanleihen Euroraum, 10 Jahre"},
    }
    OUT.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(f"geschrieben: {OUT}")


if __name__ == "__main__":
    main()
