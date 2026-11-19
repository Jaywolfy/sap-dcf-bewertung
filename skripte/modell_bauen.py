"""Baut die Arbeitsmappe modell/SAP_DCF.xlsx aus den abgerufenen Daten und den Annahmen.

Jede Zahl im Modell ist entweder eine Eingabe (blaue Schrift) oder eine Excel-Formel. Wer in Excel eine blaue
Zelle ändert, sieht sofort den neuen Wert je Aktie. Farben der Kopfzeilen wie jasonwolf.de.
"""

from datetime import date

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import annahmen
from eingaben import BASE_YEAR, ROOT, build_inputs, load_data

NAVY, GOLD, PAPER, INPUT_FILL = "182337", "97742F", "F6F4EF", "F3E9D2"
FONT = "Arial"
MIO = '#,##0;(#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
EUR = '#,##0.00'
MULT = '0.0"x"'

HEAD = Font(name=FONT, bold=True, color="FFFFFF", size=10)
BOLD = Font(name=FONT, bold=True, color=NAVY, size=10)
TEXT = Font(name=FONT, color="000000", size=10)
NOTE = Font(name=FONT, color="6B7488", size=9, italic=True)
INPUT = Font(name=FONT, color="0000FF", size=10)
LINK = Font(name=FONT, color="008000", size=10)
TITLE = Font(name=FONT, bold=True, color=NAVY, size=16)
RESULT = Font(name=FONT, bold=True, color=NAVY, size=12)
HEAD_FILL = PatternFill("solid", fgColor=NAVY)
IN_FILL = PatternFill("solid", fgColor=INPUT_FILL)
RESULT_FILL = PatternFill("solid", fgColor=PAPER)
RULE = Border(bottom=Side(style="thin", color=GOLD))
TOP_RULE = Border(top=Side(style="thin", color=NAVY))


def put(ws, ref, value, font=TEXT, fmt=None, fill=None, align=None, border=None):
    cell = ws[ref]
    cell.value, cell.font = value, font
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    if align:
        cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=align == "wrap")
    if border:
        cell.border = border
    return cell


def header(ws, row, labels, first_col=1):
    for offset, label in enumerate(labels):
        put(ws, f"{get_column_letter(first_col + offset)}{row}", label, HEAD, fill=HEAD_FILL,
            align="left" if offset == 0 else "right")


def title(ws, text, subtitle):
    put(ws, "A1", text, TITLE)
    put(ws, "A2", subtitle, NOTE)
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 26


def widths(ws, first, others, count):
    ws.column_dimensions["A"].width = first
    for index in range(2, count + 2):
        ws.column_dimensions[get_column_letter(index)].width = others


def sheet_history(wb, data):
    ws = wb.create_sheet("Historie")
    years = ["2022", "2023", "2024", "2025"]
    title(ws, "Historie", f"SAP SE, IFRS, Mio. EUR. Quelle: Yahoo Finance, abgerufen am {data['abgerufen_am']}. "
                          "Vor einer Verwendung mit dem Geschäftsbericht der SAP abgleichen.")
    header(ws, 3, ["Mio. EUR", *years])
    statements = data["abschluss"]
    rows = [  # (Zeile, Beschriftung, Schlüssel, Vorzeichen)
        (4, "Umsatz", "umsatz", 1), (6, "Betriebsergebnis", "betriebsergebnis", 1),
        (8, "Ergebnis vor Steuern", "ergebnis_vor_steuern", 1), (9, "Ertragsteuern", "steuern", 1),
        (11, "Jahresüberschuss", "jahresueberschuss", 1), (12, "Abschreibungen", "abschreibungen", 1),
        (14, "Investitionen", "investitionen", -1), (16, "Operativer Cashflow", "operativer_cashflow", 1),
        (17, "Free Cashflow (operativer Cashflow minus Investitionen)", "free_cashflow", 1),
        (19, "Aktienbasierte Vergütung", "aktienbasierte_verguetung", 1),
        (21, "Finanzschulden", "finanzschulden", 1), (22, "Liquide Mittel und kurzfristige Anlagen", "liquide_mittel", 1),
        (24, "Ausstehende Aktien (Mio. Stück)", "aktien_ausstehend", 1),
    ]
    for row, label, key, sign in rows:
        put(ws, f"A{row}", label)
        for col, year in zip("BCDE", years, strict=True):
            put(ws, f"{col}{row}", sign * statements[key][year], INPUT, "#,##0.0" if row == 24 else MIO)
    ratios = [  # (Zeile, Beschriftung, Formel mit {c} für die Spalte)
        (7, "  Operative Marge", "={c}6/{c}4"), (10, "  Steuerquote", "={c}9/{c}8"),
        (13, "  in % vom Umsatz", "={c}12/{c}4"), (15, "  in % vom Umsatz", "={c}14/{c}4"),
        (18, "  Free-Cashflow-Marge", "={c}17/{c}4"),
    ]
    for row, label, formula in ratios:
        put(ws, f"A{row}", label, NOTE)
        for col in "BCDE":
            put(ws, f"{col}{row}", formula.format(c=col), TEXT, PCT)
    put(ws, "A5", "  Wachstum zum Vorjahr", NOTE)
    for previous, col in zip("BCD", "CDE", strict=True):
        put(ws, f"{col}5", f"={col}4/{previous}4-1", TEXT, PCT)
    put(ws, "A23", "Nettoliquidität (liquide Mittel minus Finanzschulden)", BOLD)
    for col in "BCDE":
        put(ws, f"{col}23", f"={col}22-{col}21", BOLD, MIO, border=TOP_RULE)
    put(ws, "A26", "Hinweis: Das Betriebsergebnis von Yahoo Finance enthält 2024 nicht die Restrukturierungskosten "
                   "von rund 3 Mrd. EUR; das berichtete IFRS-Betriebsergebnis 2024 liegt entsprechend niedriger.", NOTE)
    widths(ws, 54, 13, 4)


def sheet_assumptions(wb, data, inputs):
    ws = wb.create_sheet("Annahmen")
    title(ws, "Annahmen", "Blaue Schrift auf hellem Grund: Eingaben, die sich ändern lassen. Alles andere rechnet sich daraus.")
    header(ws, 3, ["Marktdaten", "Wert", "Quelle oder Begründung"])
    put(ws, "A4", "Aktienkurs (EUR)")
    put(ws, "B4", inputs.share_price, INPUT, EUR, IN_FILL)
    put(ws, "C4", f"Xetra-Schlusskurs vom {data['kurs']['tag']}, Yahoo Finance", NOTE)
    for row, label, ref, fmt in ((5, "Ausstehende Aktien (Mio. Stück)", "E24", "#,##0.0"), (6, "Finanzschulden (Mio. EUR)", "E21", MIO),
                                 (7, "Liquide Mittel (Mio. EUR)", "E22", MIO)):
        put(ws, f"A{row}", label)
        put(ws, f"B{row}", f"=Historie!{ref}", LINK, fmt)
        put(ws, f"C{row}", f"Bilanz zum 31.12.{BASE_YEAR}, Blatt Historie", NOTE)

    header(ws, 9, ["Kapitalkosten", "Wert", "Quelle oder Begründung"])
    capital = [
        (10, "Risikofreier Zins", inputs.risk_free, f"{data['risikofreier_zins']['quelle']}, Stand {data['risikofreier_zins']['tag']}"),
        (11, "Beta", inputs.beta, f"Eigene Berechnung: {data['beta']['gegen']} ({data['beta']['wochen']} Wochen)"),
        (12, "Marktrisikoprämie", inputs.equity_risk_premium, "Annahme; übliche Spanne für Deutschland 4,5 bis 5,5 %"),
        (14, "Fremdkapitalzins vor Steuern", inputs.cost_of_debt, "Annahme für ein Unternehmen mit sehr guter Bonität"),
        (15, "Steuersatz", inputs.tax_rate, "Annahme; Steuerquote 2025 siehe Blatt Historie"),
    ]
    for row, label, value, source in capital:
        put(ws, f"A{row}", label)
        put(ws, f"B{row}", value, INPUT, "0.00" if row == 11 else "0.00%", IN_FILL)
        put(ws, f"C{row}", source, NOTE)
    derived = [
        (13, "Eigenkapitalkosten (CAPM)", "=B10+B11*B12", "0.00%", "Risikofreier Zins + Beta × Marktrisikoprämie"),
        (16, "Fremdkapitalzins nach Steuern", "=B14*(1-B15)", "0.00%", "Zinsen mindern die Steuerlast"),
        (17, "Marktwert Eigenkapital (Mio. EUR)", "=B4*B5", MIO, "Kurs × Aktien"),
        (18, "Anteil Eigenkapital", "=B17/(B17+B6)", "0.0%", "Gewichtung nach Marktwerten"),
        (19, "Anteil Fremdkapital", "=1-B18", "0.0%", ""),
    ]
    for row, label, formula, fmt, source in derived:
        put(ws, f"A{row}", label)
        put(ws, f"B{row}", formula, TEXT, fmt)
        put(ws, f"C{row}", source, NOTE)
    put(ws, "A21", "WACC (gewichtete Kapitalkosten)", BOLD)
    put(ws, "B21", "=B18*B13+B19*B16", BOLD, "0.00%", RESULT_FILL, border=TOP_RULE)
    put(ws, "C21", "Mit diesem Satz werden alle künftigen Cashflows abgezinst", NOTE)

    header(ws, 23, ["Operative Annahmen", "Wert", "Quelle oder Begründung"])
    operating = [
        (24, f"Ewiges Wachstum nach {annahmen.FORECAST_YEARS[-1]}", inputs.terminal_growth, "Annahme; darf nicht über dem langfristigen Wirtschaftswachstum liegen"),
        (25, "Abschreibungen in % vom Umsatz", inputs.depreciation_rate, "Wie 2025 (Blatt Historie)"),
        (26, "Investitionen in % vom Umsatz", inputs.capex_rate, "Wie 2025 (Blatt Historie)"),
        (27, "Umlaufvermögen je Euro Umsatzzuwachs", inputs.nwc_rate, "Vorsichtige Annahme; SAP erhält Zahlungen oft im Voraus"),
    ]
    for row, label, value, source in operating:
        put(ws, f"A{row}", label)
        put(ws, f"B{row}", value, INPUT, "0.0%", IN_FILL)
        put(ws, f"C{row}", source, NOTE)

    header(ws, 29, ["Prognose je Jahr", *[str(year) for year in annahmen.FORECAST_YEARS]])
    put(ws, "A30", "Umsatzwachstum")
    put(ws, "A31", "Operative Marge")
    forecast_cols = [get_column_letter(index) for index in range(2, 2 + len(annahmen.FORECAST_YEARS))]  # B bis P
    for col, growth, margin in zip(forecast_cols, inputs.revenue_growth, inputs.ebit_margin, strict=True):
        put(ws, f"{col}30", growth, INPUT, "0.0%", IN_FILL)
        put(ws, f"{col}31", margin, INPUT, "0.0%", IN_FILL)
    put(ws, "A32", "Zum Vergleich: SAP wuchs 2023 bis 2025 um 5,7 %, 9,5 % und 7,7 %; die operative Marge lag 2025 bei 26,1 %.", NOTE)
    widths(ws, 40, 13, len(forecast_cols))
    ws.column_dimensions["C"].width = 13


def sheet_dcf(wb):
    ws = wb.create_sheet("DCF")
    title(ws, "DCF-Bewertung", "Mio. EUR. Bewertungsstichtag 31.12.2025; alle Cashflows fallen am Jahresende an.")
    years = [f"{BASE_YEAR} (Ist)", *[f"{year} (Plan)" for year in annahmen.FORECAST_YEARS]]
    header(ws, 4, ["Mio. EUR", *years])
    plan = [get_column_letter(index) for index in range(3, 3 + len(annahmen.FORECAST_YEARS))]  # C bis Q
    last = plan[-1]
    put(ws, "A5", "Jahre ab Stichtag", NOTE)
    for index, col in enumerate(plan, start=1):
        put(ws, f"{col}5", index, NOTE, "0")
    labels = {6: "Umsatz", 7: "  Wachstum", 8: "Betriebsergebnis (EBIT)", 9: "  Operative Marge", 10: "Steuern auf das Betriebsergebnis",
              11: "Betriebsergebnis nach Steuern (NOPAT)", 12: "+ Abschreibungen", 13: "− Investitionen",
              14: "− Aufbau Umlaufvermögen", 15: "Freier Cashflow", 16: "Abzinsungsfaktor", 17: "Barwert des freien Cashflows"}
    for row, label in labels.items():
        put(ws, f"A{row}", label, BOLD if row in (15, 17) else NOTE if label.startswith("  ") else TEXT)
    put(ws, "B6", "=Historie!E4", LINK, MIO)
    put(ws, "B8", "=Historie!E6", LINK, MIO)
    put(ws, "B9", "=B8/B6", TEXT, PCT)
    for index, col in enumerate(plan):
        previous = "B" if index == 0 else plan[index - 1]
        driver = get_column_letter(index + 2)  # Spalte des Jahres auf dem Blatt Annahmen, ab B
        put(ws, f"{col}6", f"={previous}6*(1+Annahmen!{driver}30)", TEXT, MIO)
        put(ws, f"{col}7", f"={col}6/{previous}6-1", TEXT, PCT)
        put(ws, f"{col}8", f"={col}6*Annahmen!{driver}31", TEXT, MIO)
        put(ws, f"{col}9", f"={col}8/{col}6", TEXT, PCT)
        put(ws, f"{col}10", f"=-{col}8*Annahmen!$B$15", TEXT, MIO)
        put(ws, f"{col}11", f"={col}8+{col}10", TEXT, MIO)
        put(ws, f"{col}12", f"={col}6*Annahmen!$B$25", TEXT, MIO)
        put(ws, f"{col}13", f"=-{col}6*Annahmen!$B$26", TEXT, MIO)
        put(ws, f"{col}14", f"=-({col}6-{previous}6)*Annahmen!$B$27", TEXT, MIO)
        put(ws, f"{col}15", f"=SUM({col}11:{col}14)", BOLD, MIO, border=TOP_RULE)
        put(ws, f"{col}16", f"=1/(1+Annahmen!$B$21)^{col}5", TEXT, "0.000")
        put(ws, f"{col}17", f"={col}15*{col}16", BOLD, MIO)

    header(ws, 19, ["Vom Cashflow zum Wert je Aktie", "Mio. EUR"])
    steps = [
        (20, f"Summe der Barwerte {annahmen.FORECAST_YEARS[0]} bis {annahmen.FORECAST_YEARS[-1]}",
         f"=SUM({plan[0]}17:{last}17)", MIO, TEXT),
        (21, f"Endwert {annahmen.FORECAST_YEARS[-1]} (alle Cashflows danach)",
         f"={last}15*(1+Annahmen!B24)/(Annahmen!B21-Annahmen!B24)", MIO, TEXT),
        (22, "Barwert des Endwerts", f"=B21*{last}16", MIO, TEXT),
        (23, "Unternehmenswert", "=B20+B22", MIO, BOLD),
        (24, "+ Nettoliquidität", "=Annahmen!B7-Annahmen!B6", MIO, TEXT),
        (25, "Wert des Eigenkapitals", "=B23+B24", MIO, BOLD),
        (26, "÷ Ausstehende Aktien (Mio. Stück)", "=Annahmen!B5", "#,##0.0", TEXT),
    ]
    for row, label, formula, fmt, font in steps:
        put(ws, f"A{row}", label, font)
        put(ws, f"B{row}", formula, font, fmt, border=TOP_RULE if font is BOLD else None)
    put(ws, "A27", "Wert je Aktie (EUR)", RESULT)
    put(ws, "B27", "=B25/B26", RESULT, EUR, RESULT_FILL, border=RULE)
    put(ws, "A28", "Aktienkurs (EUR)")
    put(ws, "B28", "=Annahmen!B4", LINK, EUR)
    put(ws, "A29", "Abweichung des Werts vom Kurs", BOLD)
    put(ws, "B29", "=B27/B28-1", BOLD, '+0.0%;-0.0%;0.0%')

    header(ws, 31, ["Plausibilität", "Wert"])
    checks = [
        (32, "Anteil des Endwerts am Unternehmenswert", "=B22/B23", "0.0%"),
        (33, "Unternehmenswert / Betriebsergebnis 2026", "=B23/C8", MULT),
        (34, "Freier Cashflow 2026 / Unternehmenswert", "=C15/B23", "0.0%"),
        (35, "Freier Cashflow 2026 gegenüber Free Cashflow 2025", "=C15/Historie!E17-1", '+0.0%;-0.0%;0.0%'),
    ]
    for row, label, formula, fmt in checks:
        put(ws, f"A{row}", label)
        put(ws, f"B{row}", formula, TEXT, fmt)
    widths(ws, 46, 14, len(plan) + 1)


def sheet_sensitivity(wb):
    ws = wb.create_sheet("Sensitivität")
    title(ws, "Sensitivität", "Wert je Aktie in EUR, wenn sich Kapitalkosten (Zeilen) und ewiges Wachstum (Spalten) ändern.")
    put(ws, "A4", "WACC ↓   Wachstum →", HEAD, fill=HEAD_FILL)
    for col, step in zip("BCDEF", (-0.01, -0.005, 0, 0.005, 0.01), strict=True):
        put(ws, f"{col}4", f"=Annahmen!$B$24{step:+}", HEAD, "0.0%", HEAD_FILL, "right")
    for row, step in zip(range(5, 10), (-0.01, -0.005, 0, 0.005, 0.01), strict=True):
        put(ws, f"A{row}", f"=Annahmen!$B$21{step:+}", BOLD, "0.00%")
        for col in "BCDEF":
            centre = row == 7 and col == "D"
            put(ws, f"{col}{row}",
                f"=(SUMPRODUCT(DCF!$C$15:$Q$15/(1+$A{row})^DCF!$C$5:$Q$5)"
                f"+DCF!$Q$15*(1+{col}$4)/($A{row}-{col}$4)/(1+$A{row})^{len(annahmen.FORECAST_YEARS)}+DCF!$B$24)/DCF!$B$26",
                RESULT if centre else TEXT, EUR, RESULT_FILL if centre else None)
    put(ws, "A11", "Aktienkurs zum Vergleich (EUR)")
    put(ws, "B11", "=Annahmen!B4", LINK, EUR)
    put(ws, "A13", "Lesehilfe: Das mittlere Feld ist der Wert aus dem Blatt DCF. Ein halber Prozentpunkt mehr oder weniger "
                   "bei den Kapitalkosten verschiebt den Wert stärker als jede andere einzelne Annahme.", NOTE)
    widths(ws, 30, 13, 5)


def sheet_overview(wb, data):
    ws = wb.active
    ws.title = "Übersicht"
    title(ws, "SAP SE: Unternehmensbewertung nach dem DCF-Verfahren",
          f"Jason Thomas Wolf · Daten vom {data['abgerufen_am']} · Studienprojekt, keine Anlageberatung")
    header(ws, 4, ["Ergebnis", "Wert"])
    summary = [
        (5, "Wert je Aktie laut Modell (EUR)", "=DCF!B27", EUR, RESULT),
        (6, "Aktienkurs (EUR)", "=Annahmen!B4", EUR, LINK),
        (7, "Abweichung des Werts vom Kurs", "=DCF!B29", '+0.0%;-0.0%;0.0%', LINK),
        (8, "Unternehmenswert (Mio. EUR)", "=DCF!B23", MIO, LINK),
        (9, "WACC", "=Annahmen!B21", "0.00%", LINK),
        (10, "Ewiges Wachstum", "=Annahmen!B24", "0.0%", LINK),
        (11, "Anteil des Endwerts am Unternehmenswert", "=DCF!B32", "0.0%", LINK),
    ]
    for row, label, formula, fmt, font in summary:
        put(ws, f"A{row}", label, RESULT if row == 5 else TEXT)
        put(ws, f"B{row}", formula, font, fmt, RESULT_FILL if row == 5 else None)
    header(ws, 13, ["Aufbau der Mappe", ""])
    guide = [
        ("Historie", "Jahresabschlüsse 2022 bis 2025 und daraus abgeleitete Quoten"),
        ("Annahmen", "Alle Eingaben: Marktdaten, Kapitalkosten, Wachstum und Marge je Jahr"),
        ("DCF", f"Prognose der freien Cashflows {annahmen.FORECAST_YEARS[0]} bis {annahmen.FORECAST_YEARS[-1]} "
                "und Herleitung des Werts je Aktie"),
        ("Sensitivität", "Wert je Aktie bei anderen Kapitalkosten und anderem ewigen Wachstum"),
        ("Zeitverlauf", "Kennzahlen über die Jahre mit Diagrammen und Kursverlauf der letzten fünf Jahre"),
        ("Langfristig", "Umsatz und Betriebsergebnis seit 2000, Kursverlauf seit 1998 und Diagramme"),
    ]
    for row, (name, text) in enumerate(guide, start=14):
        put(ws, f"A{row}", name, BOLD)
        put(ws, f"B{row}", text)
    header(ws, 20, ["Farben", ""])
    put(ws, "A21", "Blaue Schrift auf hellem Grund", INPUT, fill=IN_FILL)
    put(ws, "B21", "Eingabe; lässt sich ändern")
    put(ws, "A22", "Schwarze Schrift")
    put(ws, "B22", "Formel innerhalb des Blatts")
    put(ws, "A23", "Grüne Schrift", LINK)
    put(ws, "B23", "Übernahme von einem anderen Blatt")
    put(ws, "A25", "Einschränkung: Der Stichtag der Bilanz ist der 31.12.2025, der Aktienkurs stammt vom "
                   f"{data['kurs']['tag']}. Was SAP seit dem Stichtag verdient und ausgeschüttet hat, ist nicht berücksichtigt.", NOTE)
    widths(ws, 46, 18, 1)
    ws.column_dimensions["B"].width = 70
    for row in range(5, 12):
        ws[f"B{row}"].alignment = Alignment(horizontal="left")


def style_chart(chart, heading, y_format, width=16, height=7.5):
    chart.title = heading
    chart.width, chart.height = width, height
    chart.legend.position = "b"
    chart.y_axis.number_format = y_format
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.x_axis.tickLblPos = "low"


def color_series(chart, colors, line=False):
    for series, color in zip(chart.series, colors, strict=True):
        if line:
            series.graphicalProperties.line.solidFill = color
            series.graphicalProperties.line.width = 22000
            series.smooth = False
        else:
            series.graphicalProperties.solidFill = color
            series.graphicalProperties.line.solidFill = color


def sheet_timeline(wb, data):
    ws = wb.create_sheet("Zeitverlauf")
    years = ["2022", "2023", "2024", "2025"]
    title(ws, "Zeitverlauf", "SAP SE. Jahreswerte aus dem Blatt Historie (grüne Schrift), Kurse wöchentlich über fünf "
                             f"Jahre. Quelle: Yahoo Finance, abgerufen am {data['abgerufen_am']}.")
    header(ws, 4, ["Mio. EUR", *years])
    links = [  # (Zeile, Beschriftung, Formel mit {c} für die Spalte in Historie, Format)
        (5, "Umsatz", "=Historie!{c}4", MIO),
        (7, "Betriebsergebnis", "=Historie!{c}6", MIO),
        (9, "Jahresüberschuss", "=Historie!{c}11", MIO),
        (10, "Free Cashflow", "=Historie!{c}17", MIO),
        (12, "Nettoliquidität", "=Historie!{c}23", MIO),
        (13, "Ausstehende Aktien (Mio. Stück)", "=Historie!{c}24", "#,##0.0"),
    ]
    for row, label, formula, fmt in links:
        put(ws, f"A{row}", label)
        for col in "BCDE":
            put(ws, f"{col}{row}", formula.format(c=col), LINK, fmt)
    ratios = [  # (Zeile, Beschriftung, Formel mit {c} für die Spalte, Format)
        (8, "  Operative Marge", "={c}7/{c}5", PCT),
        (11, "  Free-Cashflow-Marge", "={c}10/{c}5", PCT),
        (14, "Free Cashflow je Aktie (EUR)", "={c}10/{c}13", EUR),
        (15, "Jahresüberschuss je Aktie (EUR)", "={c}9/{c}13", EUR),
    ]
    for row, label, formula, fmt in ratios:
        put(ws, f"A{row}", label, NOTE if label.startswith("  ") else TEXT)
        for col in "BCDE":
            put(ws, f"{col}{row}", formula.format(c=col), TEXT, fmt)
    put(ws, "A6", "  Wachstum zum Vorjahr", NOTE)
    for previous, col in zip("BCD", "CDE", strict=True):
        put(ws, f"{col}6", f"={col}5/{previous}5-1", TEXT, PCT)
    put(ws, "A16", "Das Betriebsergebnis 2024 ist ohne Restrukturierungskosten (siehe Historie). Nettoliquidität ist "
                   "liquide Mittel minus Finanzschulden.", NOTE)

    years_ref = Reference(ws, min_col=2, max_col=5, min_row=4)
    bars = BarChart()
    bars.type = "col"
    for row, label in ((5, "Umsatz"), (7, "Betriebsergebnis")):
        bars.series.append(Series(Reference(ws, min_col=2, max_col=5, min_row=row), title=label))
    bars.set_categories(years_ref)
    style_chart(bars, "Umsatz und Betriebsergebnis (Mio. EUR)", "#,##0")
    color_series(bars, [NAVY, GOLD])
    ws.add_chart(bars, "A18")

    margins = LineChart()
    for row, label in ((8, "Operative Marge"), (11, "Free-Cashflow-Marge")):
        margins.series.append(Series(Reference(ws, min_col=2, max_col=5, min_row=row), title=label))
    margins.set_categories(years_ref)
    style_chart(margins, "Margen in % vom Umsatz", "0%")
    color_series(margins, [NAVY, GOLD], line=True)
    ws.add_chart(margins, "A34")

    cash = BarChart()
    cash.type = "col"
    for row, label in ((10, "Free Cashflow"), (12, "Nettoliquidität")):
        cash.series.append(Series(Reference(ws, min_col=2, max_col=5, min_row=row), title=label))
    cash.set_categories(years_ref)
    style_chart(cash, "Free Cashflow und Nettoliquidität (Mio. EUR)", "#,##0")
    color_series(cash, [NAVY, GOLD])
    ws.add_chart(cash, "A50")

    per_share = LineChart()
    for row, label in ((14, "Free Cashflow je Aktie"), (15, "Jahresüberschuss je Aktie")):
        per_share.series.append(Series(Reference(ws, min_col=2, max_col=5, min_row=row), title=label))
    per_share.set_categories(years_ref)
    style_chart(per_share, "Je Aktie (EUR)", "0.00")
    color_series(per_share, [NAVY, GOLD], line=True)
    ws.add_chart(per_share, "A66")

    weeks = data["kurs_wochen"]
    first = 96
    last = first + len(weeks) - 1
    put(ws, "A84", "Kurs über fünf Jahre (wöchentliche Schlusskurse)", BOLD)
    stats = [  # (Zeile, Beschriftung, Formel, Format)
        (85, "Letzter Schlusskurs (EUR)", f"=B{last}", EUR),
        (86, "Hoch, fünf Jahre (EUR)", f"=MAX(B{first}:B{last})", EUR),
        (87, "Tief, fünf Jahre (EUR)", f"=MIN(B{first}:B{last})", EUR),
        (88, "Durchschnitt, fünf Jahre (EUR)", f"=AVERAGE(B{first}:B{last})", EUR),
        (89, "Hoch, letzte 52 Wochen (EUR)", f"=MAX(B{last - 51}:B{last})", EUR),
        (90, "Tief, letzte 52 Wochen (EUR)", f"=MIN(B{last - 51}:B{last})", EUR),
        (91, "Veränderung, letzte 52 Wochen", f"=B{last}/B{last - 52}-1", '+0.0%;-0.0%;0.0%'),
        (92, "Veränderung, fünf Jahre", f"=B{last}/B{first}-1", '+0.0%;-0.0%;0.0%'),
        (93, "SAP minus DAX seit Beginn (Prozentpunkte)", f"=(B{last}/B{first}-1)-(C{last}/C{first}-1)",
         '+0.0%;-0.0%;0.0%'),
    ]
    for row, label, formula, fmt in stats:
        put(ws, f"A{row}", label)
        put(ws, f"B{row}", formula, TEXT, fmt)
    put(ws, "A94", "DAX skaliert = DAX-Stand geteilt durch den DAX-Startwert, mal SAP-Startkurs. Die Hoch- und "
                   "Tiefwerte beruhen auf Wochenschlüssen.", NOTE)
    header(ws, 95, ["Woche (Freitag)", "SAP (EUR)", "DAX (Punkte)", "DAX skaliert (EUR)"])
    for offset, week in enumerate(weeks):
        row = first + offset
        put(ws, f"A{row}", date.fromisoformat(week["datum"]), INPUT, "dd.mm.yyyy", align="left")
        put(ws, f"B{row}", week["sap"], INPUT, EUR)
        put(ws, f"C{row}", week["dax"], INPUT, "#,##0.00")
        put(ws, f"D{row}", f"=C{row}/$C${first}*$B${first}", TEXT, EUR)

    price = LineChart()
    for col, label in ((2, "SAP (EUR)"), (4, "DAX skaliert (EUR)")):
        price.series.append(Series(Reference(ws, min_col=col, min_row=first, max_row=last), title=label))
    price.set_categories(Reference(ws, min_col=1, min_row=first, max_row=last))
    style_chart(price, "SAP-Kurs über fünf Jahre im Vergleich zum DAX", "#,##0", width=22, height=10)
    price.x_axis.number_format = "mmm yy"
    price.x_axis.tickLblSkip = 26
    color_series(price, [NAVY, GOLD], line=True)
    ws.add_chart(price, "F84")

    widths(ws, 46, 13, 4)


def sheet_long(wb, data):
    ws = wb.create_sheet("Langfristig")
    title(ws, "Langfristig", "SAP SE. Umsatz und Betriebsergebnis in Mio. EUR, Kurse in EUR. Quellen: Geschäftsberichte "
                             f"der SAP für 2000 bis 2021, Yahoo Finance für 2022 bis 2025 und die Kurse "
                             f"(abgerufen am {data['abgerufen_am']}).")
    header(ws, 4, ["Jahr", "Umsatz", "Wachstum", "Betriebsergebnis", "Marge", "Quelle"])
    revenue = data["umsatz_lang"]["werte"]
    operating = data["betriebsergebnis_lang"]["werte"]
    recent = {"2022": "B", "2023": "C", "2024": "D", "2025": "E"}  # Spalten im Blatt Historie
    years = [str(year) for year in range(2000, 2026)]
    for offset, year in enumerate(years):
        row = 5 + offset
        put(ws, f"A{row}", year, TEXT, align="left")
        if year in recent:
            put(ws, f"B{row}", f"=Historie!{recent[year]}4", LINK, MIO)
            put(ws, f"D{row}", f"=Historie!{recent[year]}6", LINK, MIO)
            put(ws, f"F{row}", "Yahoo Finance, Blatt Historie", NOTE)
        else:
            put(ws, f"B{row}", revenue[year], INPUT, MIO)
            if year in operating:
                put(ws, f"D{row}", operating[year], INPUT, MIO)
            put(ws, f"F{row}", "Geschäftsbericht der SAP", NOTE)
        if offset:
            put(ws, f"C{row}", f"=B{row}/B{row - 1}-1", TEXT, PCT)
        if year in operating or year in recent:
            put(ws, f"E{row}", f"=D{row}/B{row}", TEXT, PCT)
    last = 5 + len(years) - 1
    put(ws, f"A{last + 2}", "Jährliches Umsatzwachstum 2000 bis 2025 (CAGR)")
    put(ws, f"B{last + 2}", f"=(B{last}/B5)^(1/{len(years) - 1})-1", TEXT, PCT)
    put(ws, f"A{last + 3}", "Jährliches Umsatzwachstum 2015 bis 2025 (CAGR)")
    put(ws, f"B{last + 3}", f"=(B{last}/B20)^(1/10)-1", TEXT, PCT)
    put(ws, f"A{last + 5}", "Umsatz 2000 bis 2021 aus den Geschäftsberichten (IFRS, jeweils der neueste Ausweis). 1997 bis 1999 "
                            "fehlen, weil die Berichte in D-Mark sind und keine eindeutige Summe nennen.", NOTE)
    put(ws, f"A{last + 6}", "Betriebsergebnis 2011 bis 2021 nach IFRS aus den Geschäftsberichten. Ab 2022 steht Yahoo "
                            "Finance, das vom IFRS-Wert abweichen kann (siehe Blatt Historie).", NOTE)

    header(ws, 4, ["Monat", "SAP-Schluss (EUR)"], first_col=8)
    monthly = data["kurs_monate"]
    for offset, (day, close) in enumerate(monthly):
        row = 5 + offset
        put(ws, f"H{row}", date.fromisoformat(day), INPUT, "mm.yyyy", align="left")
        put(ws, f"I{row}", close, INPUT, EUR)
    price_last = 5 + len(monthly) - 1

    revenue_chart = BarChart()
    revenue_chart.type = "col"
    revenue_chart.series.append(Series(Reference(ws, min_col=2, min_row=5, max_row=last), title="Umsatz"))
    revenue_chart.series.append(Series(Reference(ws, min_col=4, min_row=5, max_row=last), title="Betriebsergebnis"))
    revenue_chart.set_categories(Reference(ws, min_col=1, min_row=5, max_row=last))
    style_chart(revenue_chart, "Umsatz und Betriebsergebnis 2000 bis 2025 (Mio. EUR)", "#,##0", width=22, height=9)
    color_series(revenue_chart, [NAVY, GOLD])
    ws.add_chart(revenue_chart, "K4")

    price_chart = LineChart()
    price_chart.series.append(Series(Reference(ws, min_col=9, min_row=5, max_row=price_last), title="SAP-Schluss (EUR)"))
    price_chart.set_categories(Reference(ws, min_col=8, min_row=5, max_row=price_last))
    style_chart(price_chart, "SAP-Kurs seit 1998 (Monatsschluss, EUR)", "#,##0", width=22, height=10)
    price_chart.x_axis.number_format = "yyyy"
    price_chart.x_axis.tickLblSkip = 24
    color_series(price_chart, [NAVY], line=True)
    ws.add_chart(price_chart, "K24")

    widths(ws, 44, 13, 4)
    ws.column_dimensions["E"].width = 10
    ws.column_dimensions["F"].width = 28
    ws.column_dimensions["G"].width = 3
    ws.column_dimensions["H"].width = 12
    ws.column_dimensions["I"].width = 16


def build():
    data = load_data()
    inputs = build_inputs(data)
    wb = Workbook()
    sheet_overview(wb, data)
    sheet_history(wb, data)
    sheet_assumptions(wb, data, inputs)
    sheet_dcf(wb)
    sheet_sensitivity(wb)
    sheet_timeline(wb, data)
    sheet_long(wb, data)
    for ws in wb.worksheets:
        ws.sheet_properties.tabColor = NAVY
        ws.freeze_panes = "B5" if ws.title in ("Historie", "DCF") else None
    out = ROOT / "modell" / "SAP_DCF.xlsx"
    wb.save(out)
    print(f"geschrieben: {out}")


if __name__ == "__main__":
    build()
