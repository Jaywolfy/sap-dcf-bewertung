# SAP SE: Unternehmensbewertung nach dem DCF-Verfahren

Was ist eine SAP-Aktie wert, wenn man nur die künftigen freien Cashflows betrachtet? Dieses Projekt beantwortet die
Frage mit einem Discounted-Cashflow-Modell in Excel und einer zweiten, unabhängigen Rechnung in Python.

**Ergebnis:** Mit den Annahmen unten ergibt das Modell einen Wert von rund **169 EUR je Aktie**. Der Xetra-Schlusskurs
vom 13.11.2026 lag bei 213,45 EUR, also etwa 21 % darüber. Der Markt rechnet demnach mit mehr Wachstum oder höheren
Margen, als hier angenommen sind.

Das ist ein Studienprojekt und keine Anlageberatung.

## Inhalt

| Datei | Inhalt |
|---|---|
| [`modell/SAP_DCF.xlsx`](modell/SAP_DCF.xlsx) | Arbeitsmappe mit fünf Blättern: Übersicht, Historie, Annahmen, DCF, Sensitivität. Alle Werte sind Formeln; Eingaben stehen in blauer Schrift. |
| `skripte/daten_abrufen.py` | lädt Abschlusszahlen, Kurs, Beta-Grundlage und risikofreien Zins |
| `skripte/annahmen.py` | alle Annahmen mit Begründung an einer Stelle |
| `skripte/dcf.py` | dieselbe Bewertung als reine Python-Rechnung |
| `skripte/modell_bauen.py` | erzeugt die Arbeitsmappe |
| `daten/sap_kennzahlen.json` | Stand der abgerufenen Daten |

## Methode

1. **Freier Cashflow 2026 bis 2030:** Betriebsergebnis nach Steuern plus Abschreibungen minus Investitionen minus
   Aufbau von Umlaufvermögen.
2. **Kapitalkosten (WACC):** Eigenkapitalkosten nach dem CAPM, Fremdkapitalkosten nach Steuern, gewichtet nach
   Marktwerten.
3. **Endwert:** Cashflow 2030, der danach für immer mit einer festen Rate wächst.
4. **Wert je Aktie:** Barwerte plus Barwert des Endwerts ergeben den Unternehmenswert; plus Nettoliquidität, geteilt
   durch die ausstehenden Aktien.

## Annahmen

| Annahme | Wert | Begründung |
|---|---|---|
| Umsatzwachstum 2026 bis 2030 | 9 %, 9 %, 8 %, 7 %, 6 % | 2023 bis 2025 waren es 5,7 %, 9,5 % und 7,7 % |
| Operative Marge 2026 bis 2030 | 27 % bis 29,5 % | 2025 lag sie bei 26,1 % |
| Steuersatz | 28 % | Steuerquote 2025: 28,7 % |
| Risikofreier Zins | 3,37 % | EZB, zehnjährige AAA-Staatsanleihen des Euroraums, 12.11.2026 |
| Beta | 0,95 | eigene Berechnung gegen den DAX, wöchentliche Renditen über fünf Jahre |
| Marktrisikoprämie | 5,0 % | Annahme |
| WACC | 7,96 % | ergibt sich aus den Zeilen darüber |
| Ewiges Wachstum | 2,5 % | Annahme |

## Sensitivität

Wert je Aktie in EUR bei anderen Kapitalkosten (Zeilen) und anderem ewigen Wachstum (Spalten):

| WACC | 1,5 % | 2,0 % | 2,5 % | 3,0 % | 3,5 % |
|---|---|---|---|---|---|
| 6,96 % | 175 | 190 | 207 | 230 | 259 |
| 7,46 % | 160 | 172 | 186 | 204 | 226 |
| **7,96 %** | 147 | 157 | **169** | 183 | 200 |
| 8,46 % | 137 | 145 | 155 | 166 | 180 |
| 8,96 % | 127 | 134 | 143 | 152 | 163 |

Der Endwert macht 79 % des Unternehmenswerts aus. Das Ergebnis hängt deshalb stark an zwei Zahlen, den
Kapitalkosten und dem ewigen Wachstum.

## Grenzen

- Die Abschlusszahlen stammen von Yahoo Finance und sind nicht mit dem Geschäftsbericht der SAP abgeglichen.
  Das Betriebsergebnis 2024 enthält dort nicht die Restrukturierungskosten.
- Stichtag der Bilanz ist der 31.12.2025, der Kurs stammt vom 13.11.2026. Was dazwischen verdient und
  ausgeschüttet wurde, ist nicht berücksichtigt.
- Das Beta gegen den DAX ist nach oben verzerrt, weil SAP selbst einen großen Teil des DAX ausmacht.
- Ein Vergleich mit Wettbewerbern über Multiplikatoren fehlt noch.

## Neu rechnen

```bash
python -m venv .venv && .venv/bin/pip install -r requirements.txt
cd skripte
../.venv/bin/python daten_abrufen.py   # aktuelle Daten laden
../.venv/bin/python eingaben.py        # Ergebnis der Python-Rechnung
../.venv/bin/python modell_bauen.py    # Arbeitsmappe neu erzeugen
```

## Lizenz

[MIT](LICENSE) © 2026 Jason Thomas Wolf
