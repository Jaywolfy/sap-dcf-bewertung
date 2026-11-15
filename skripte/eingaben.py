"""Setzt die abgerufenen Daten und die Annahmen zu den Eingaben der Bewertung zusammen."""

import json
from pathlib import Path

import annahmen
from dcf import Inputs

ROOT = Path(__file__).resolve().parents[1]
BASE_YEAR = "2025"


def load_data() -> dict:
    return json.loads((ROOT / "daten" / "sap_kennzahlen.json").read_text(encoding="utf-8"))


def build_inputs(data: dict) -> Inputs:
    statements = data["abschluss"]
    return Inputs(
        revenue=statements["umsatz"][BASE_YEAR],
        revenue_growth=annahmen.REVENUE_GROWTH,
        ebit_margin=annahmen.EBIT_MARGIN,
        tax_rate=annahmen.TAX_RATE,
        depreciation_rate=annahmen.DEPRECIATION_RATE,
        capex_rate=annahmen.CAPEX_RATE,
        nwc_rate=annahmen.NWC_RATE,
        risk_free=data["risikofreier_zins"]["wert"],
        beta=data["beta"]["wert"],
        equity_risk_premium=annahmen.EQUITY_RISK_PREMIUM,
        cost_of_debt=annahmen.COST_OF_DEBT,
        terminal_growth=annahmen.TERMINAL_GROWTH,
        share_price=data["kurs"]["schluss_eur"],
        shares=statements["aktien_ausstehend"][BASE_YEAR],
        debt=statements["finanzschulden"][BASE_YEAR],
        cash=statements["liquide_mittel"][BASE_YEAR],
    )


if __name__ == "__main__":
    from dcf import valuation

    for key, value in valuation(build_inputs(load_data())).items():
        print(f"{key:24} {value:,.4f}")
