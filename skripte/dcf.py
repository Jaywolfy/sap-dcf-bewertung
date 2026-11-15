"""Die Bewertung als reine Rechnung, ohne Excel. Dient als zweite, unabhängige Rechnung neben der Arbeitsmappe:
Die Tests prüfen, dass beide dasselbe Ergebnis liefern.

Methode: Discounted Cash Flow. Der Wert des Unternehmens ist die Summe aller künftigen freien Cashflows,
abgezinst mit den gewichteten Kapitalkosten (WACC).
"""

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class Inputs:
    revenue: float                 # Umsatz des letzten Ist-Jahres, Mio. EUR
    revenue_growth: tuple[float, ...]
    ebit_margin: tuple[float, ...]
    tax_rate: float
    depreciation_rate: float
    capex_rate: float
    nwc_rate: float
    risk_free: float
    beta: float
    equity_risk_premium: float
    cost_of_debt: float
    terminal_growth: float
    share_price: float             # EUR
    shares: float                  # Mio. Stück
    debt: float                    # Finanzschulden, Mio. EUR
    cash: float                    # Liquide Mittel, Mio. EUR


def cost_of_equity(inputs: Inputs) -> float:
    """Eigenkapitalkosten nach dem CAPM: risikofreier Zins plus Beta mal Marktrisikoprämie."""
    return inputs.risk_free + inputs.beta * inputs.equity_risk_premium


def wacc(inputs: Inputs) -> float:
    """Gewichtete Kapitalkosten. Gewichtet wird mit Marktwerten; Fremdkapitalzinsen mindern die Steuer."""
    equity = inputs.share_price * inputs.shares
    weight_equity = equity / (equity + inputs.debt)
    after_tax_debt = inputs.cost_of_debt * (1 - inputs.tax_rate)
    return weight_equity * cost_of_equity(inputs) + (1 - weight_equity) * after_tax_debt


def free_cash_flows(inputs: Inputs) -> list[float]:
    """Freier Cashflow je Prognosejahr: Betriebsergebnis nach Steuern plus Abschreibungen
    minus Investitionen minus Aufbau von Umlaufvermögen."""
    flows, revenue = [], inputs.revenue
    for growth, margin in zip(inputs.revenue_growth, inputs.ebit_margin, strict=True):
        previous, revenue = revenue, revenue * (1 + growth)
        nopat = revenue * margin * (1 - inputs.tax_rate)
        flows.append(nopat + revenue * inputs.depreciation_rate - revenue * inputs.capex_rate
                     - (revenue - previous) * inputs.nwc_rate)
    return flows


def value_per_share(inputs: Inputs, discount_rate: float | None = None, terminal_growth: float | None = None) -> float:
    """Wert je Aktie in EUR. Zinssatz und ewiges Wachstum lassen sich für die Sensitivitätstabelle überschreiben."""
    return valuation(inputs, discount_rate, terminal_growth)["value_per_share"]


def valuation(inputs: Inputs, discount_rate: float | None = None, terminal_growth: float | None = None) -> dict[str, float]:
    rate = wacc(inputs) if discount_rate is None else discount_rate
    growth = inputs.terminal_growth if terminal_growth is None else terminal_growth
    if rate <= growth:
        raise ValueError("Der Zinssatz muss über dem ewigen Wachstum liegen, sonst wäre der Endwert unendlich.")
    flows = free_cash_flows(inputs)
    present_values = sum(flow / (1 + rate) ** year for year, flow in enumerate(flows, start=1))
    terminal_value = flows[-1] * (1 + growth) / (rate - growth)
    present_terminal = terminal_value / (1 + rate) ** len(flows)
    enterprise_value = present_values + present_terminal
    equity_value = enterprise_value + inputs.cash - inputs.debt
    return {
        "wacc": rate,
        "present_value_flows": present_values,
        "terminal_value": terminal_value,
        "present_value_terminal": present_terminal,
        "enterprise_value": enterprise_value,
        "equity_value": equity_value,
        "value_per_share": equity_value / inputs.shares,
        "upside": equity_value / inputs.shares / inputs.share_price - 1,
        "terminal_share": present_terminal / enterprise_value,
    }
