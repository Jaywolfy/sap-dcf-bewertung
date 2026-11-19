# SAP SE: Unternehmensbewertung nach dem DCF-Verfahren

Was ist eine SAP-Aktie wert, wenn man nur die künftigen freien Cashflows betrachtet? Dieses Projekt beantwortet die
Frage mit einem Discounted-Cashflow-Modell in Excel und einer zweiten, unabhängigen Rechnung in Python.

**Ergebnis:** Mit den Annahmen unten ergibt das Modell einen Wert von rund **185 EUR je Aktie**. Der Xetra-Schlusskurs
vom 19.11.2026 lag bei 187,02 EUR, die Abweichung beträgt also etwa −1 %. Der Markt preist damit ungefähr die
Wachstums- und Margenannahmen dieses Modells ein.

Das ist ein Studienprojekt und keine Anlageberatung.

## Inhalt

| Datei | Inhalt |
|---|---|
| [`modell/SAP_DCF.xlsx`](modell/SAP_DCF.xlsx) | Arbeitsmappe mit sieben Blättern: Übersicht, Historie, Annahmen, DCF, Sensitivität, Zeitverlauf, Langfristig. Alle Werte sind Formeln; Eingaben stehen in blauer Schrift. |
| `skripte/daten_abrufen.py` | lädt Abschlusszahlen, Kurse, Beta-Grundlage und risikofreien Zins; Umsatz und Betriebsergebnis aus den Geschäftsberichten 2000 bis 2021 sind fest hinterlegt |
| `skripte/annahmen.py` | alle Annahmen mit Begründung an einer Stelle |
| `skripte/dcf.py` | dieselbe Bewertung als reine Python-Rechnung |
| `skripte/modell_bauen.py` | erzeugt die Arbeitsmappe |
| `daten/sap_kennzahlen.json` | Stand der abgerufenen Daten, einschließlich Wochen- und Monatskursen |

## Methode

1. **Freier Cashflow 2026 bis 2040:** Betriebsergebnis nach Steuern plus Abschreibungen minus Investitionen minus
   Aufbau von Umlaufvermögen.
2. **Kapitalkosten (WACC):** Eigenkapitalkosten nach dem CAPM, Fremdkapitalkosten nach Steuern, gewichtet nach
   Marktwerten.
3. **Endwert:** Cashflow 2040, der danach für immer mit einer festen Rate wächst.
4. **Wert je Aktie:** Barwerte plus Barwert des Endwerts ergeben den Unternehmenswert; plus Nettoliquidität, geteilt
   durch die ausstehenden Aktien.

## Annahmen

| Annahme | Wert | Begründung |
|---|---|---|
| Umsatzwachstum 2026 bis 2030 | 9 %, 9 %, 8 %, 7 %, 6 % | 2023 bis 2025 waren es 5,7 %, 9,5 % und 7,7 % |
| Umsatzwachstum 2031 bis 2040 | linear von 5,5 % auf 3,0 % | Auslaufen des Cloud-Wachstums; Annahme |
| Operative Marge 2026 bis 2030 | 27 % bis 29,5 %, dann 30 % | 2025 lag sie bei 26,1 %; danach kein weiterer Anstieg |
| Steuersatz | 28 % | Steuerquote 2025: 28,7 % |
| Risikofreier Zins | 3,41 % | EZB, zehnjährige AAA-Staatsanleihen des Euroraums, 18.11.2026 |
| Beta | 0,97 | eigene Berechnung gegen den DAX, wöchentliche Renditen über fünf Jahre |
| Marktrisikoprämie | 5,0 % | Annahme |
| WACC | 8,07 % | ergibt sich aus den Zeilen darüber |
| Ewiges Wachstum | 2,5 % | Annahme |

## Sensitivität

Wert je Aktie in EUR bei anderen Kapitalkosten (Zeilen) und anderem ewigen Wachstum (Spalten):

| WACC | 1,5 % | 2,0 % | 2,5 % | 3,0 % | 3,5 % |
|---|---|---|---|---|---|
| 7,07 % | 204 | 215 | 228 | 244 | 264 |
| 7,57 % | 186 | 194 | 204 | 216 | 231 |
| **8,07 %** | 171 | 177 | **185** | 194 | 205 |
| 8,57 % | 157 | 163 | 169 | 176 | 185 |
| 9,07 % | 146 | 150 | 155 | 161 | 168 |

Der Endwert macht 49 % des Unternehmenswerts aus; bei der früheren Fassung mit Prognose nur bis 2030 waren es 79 %,
weil jetzt mehr Jahre explizit gerechnet werden. Das Ergebnis hängt trotzdem stark an den Kapitalkosten und dem
ewigen Wachstum.

## Grenzen

- Die Abschlusszahlen 2022 bis 2025 stammen von Yahoo Finance. Die Quelle stockanalysis.com nennt für 2025 ein
  höheres Betriebsergebnis (10.418 statt 9.620 Mio. EUR). Die Unterschiede sind nicht geklärt und nicht mit dem
  Geschäftsbericht abgeglichen. Das Betriebsergebnis 2024 enthält bei Yahoo nicht die Restrukturierungskosten.
- Umsatz 2000 bis 2021 und Betriebsergebnis 2011 bis 2021 stammen aus den Geschäftsberichten der SAP
  (2001 bis 2015: `sap-investor-YYYY-annual-report.pdf` bzw. `sap-YYYY-annual-report.pdf`, 2016 bis 2021:
  `sap-YYYY-integrated-report.pdf`, jeweils unter `sap.com/docs/download/investors/YYYY/`). Die Werte wurden von
  Hand aus dem Text der PDFs übernommen. Bei Anpassungen zählt der Ausweis aus dem Folgejahr.
- Die Geschäftsberichte 1997 bis 1999 nennen D-Mark und keine eindeutige Umsatzsumme; sie fehlen deshalb.
  Berichte vor 1997 liegen auf sap.com nicht vor.
- Die Quellen stockanalysis.com und wallstreetnumbers.com weichen von den Geschäftsberichten ab (z. B. Umsatz 2021:
  26.953 gegenüber 27.842 Mio. EUR laut Bericht). Sie werden nicht verwendet.
- Die Prognose bis 2040 ist eine Einschätzung ohne Gewähr. Sie beruht auf den Annahmen oben.
- Stichtag der Bilanz ist der 31.12.2025, der Kurs stammt vom 19.11.2026. Was dazwischen verdient und
  ausgeschüttet wurde, ist nicht berücksichtigt.
- Das Beta gegen den DAX ist nach oben verzerrt, weil SAP selbst einen großen Teil des DAX ausmacht.
- Ein Vergleich mit Wettbewerbern über Multiplikatoren fehlt noch.
- Historische Daten vor 2000 fehlen im Umsatz, Kurse liegen ab 1998 vor. Das Betriebsergebnis fehlt vor 2011, weil ich
  dafür in den geprüften Berichten keine einheitliche IFRS-Reihe gefunden habe.

## Neu rechnen

```bash
python -m venv .venv && .venv/bin/pip install -r requirements.txt
cd skripte
../.venv/bin/python daten_abrufen.py   # aktuelle Daten laden
../.venv/bin/python eingaben.py        # Ergebnis der Python-Rechnung
../.venv/bin/python modell_bauen.py    # Arbeitsmappe neu erzeugen
```

Die Arbeitsmappe ist danach ohne Excel nur ohne berechnete Werte. Zum Neuberechnen öffnet man sie in Excel
oder LibreOffice und speichert sie einmal.

## Lizenz

[MIT](LICENSE) © 2026 Jason Thomas Wolf
