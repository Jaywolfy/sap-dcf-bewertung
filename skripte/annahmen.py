"""Die Annahmen des Modells an einer Stelle. Alles hier ist eine Einschätzung, keine Tatsache.

Wer eine Annahme ändern will, ändert sie hier und baut das Modell neu (python skripte/modell_bauen.py),
oder überschreibt die blau geschriebene Zelle direkt in Excel.
"""

FORECAST_YEARS = (2026, 2027, 2028, 2029, 2030)

# Umsatzwachstum je Jahr. SAP ist 2023 bis 2025 um 5,7 %, 9,5 % und 7,7 % gewachsen; der Umstieg der Kunden
# in die Cloud trägt das Wachstum noch einige Jahre und läuft dann aus.
REVENUE_GROWTH = (0.09, 0.09, 0.08, 0.07, 0.06)

# Operative Marge (Betriebsergebnis durch Umsatz). 2025 lag sie bei 26,1 %. Cloud-Umsatz skaliert, deshalb
# steigt sie hier langsam weiter.
EBIT_MARGIN = (0.27, 0.28, 0.285, 0.29, 0.295)

TAX_RATE = 0.28            # Steuerquote 2025: 28,7 %
DEPRECIATION_RATE = 0.036  # Abschreibungen in Prozent vom Umsatz, wie 2025
CAPEX_RATE = 0.02          # Investitionen in Prozent vom Umsatz, wie 2025
NWC_RATE = 0.05            # Zusätzliches Umlaufvermögen je Euro Umsatzzuwachs (vorsichtige Annahme)

EQUITY_RISK_PREMIUM = 0.05  # Mehrrendite von Aktien gegenüber Staatsanleihen; üblich sind 4,5 bis 5,5 %
COST_OF_DEBT = 0.035        # Zins vor Steuern für neue Schulden eines Unternehmens mit dieser Bonität
TERMINAL_GROWTH = 0.025     # Wachstum nach 2030 für immer; darf nicht über dem der Gesamtwirtschaft liegen
