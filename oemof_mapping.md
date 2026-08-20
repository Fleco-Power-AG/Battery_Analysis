# Batterieanalyse: Excel-Input-Mapping auf oemof.solph & Vorschläge für Erweiterungen

Stand: 2026-08-20 (aktualisiert). Basis: `Inputs.xlsx` (Sheets `Parameter`, `Hilfen_Listen`, `Zeitreihen`, 15-Min-Auflösung 2025).

## Architektur: Code ortsunabhängig, Daten am vom Nutzer gewählten Ort
Wichtige Design-Entscheidung (finalisiert): der CODE (alle .py-Dateien) kann irgendwo liegen (z.B. wohin `git clone` das Repo legt) — das ist irrelevant. Alle Daten (Inputs.xlsx, input_lp.xlsx, Ergebnis) landen dort, wo der NUTZER seine Inputs.xlsx hinlegt/auswählt, nicht im Code-Ordner. Kein Pfad ist irgendwo hartcodiert (auch keine Personennamen/Benutzerpfade — geprüft).

## Einstiegspunkt für alle Fleco-Nutzer
- **`run_batterieanalyse.py`**: orchestriert den ganzen Ablauf (fill_missing_inputs -> battery_optimization -> Ergebnis-Export). Nimmt den Input-Pfad als Kommandozeilen-Argument ODER öffnet (falls keiner angegeben) einen Datei-Auswahl-Dialog (tkinter, Standardbibliothek, kein Zusatzpaket nötig). `input_lp.xlsx` und `ergebnis.xlsx` (Format noch nicht final definiert — aktuell einfacher Excel-Export der Zeitreihen + Modul-Erträge) landen automatisch im selben Ordner wie die gewählte Inputs.xlsx.
- **`run_batterieanalyse.bat`**: Doppelklick-Startpunkt für Windows. Öffnet bei einfachem Doppelklick den Dateidialog; per Drag & Drop einer Inputs.xlsx auf das Icon wird der Pfad automatisch übernommen. Sucht automatisch nach einer lokalen `.venv`/`venv`-virtuellen Umgebung neben der .bat-Datei, sonst System-Python.
- Getestet (mit gestubbtem `battery_optimization`, da oemof.solph in dieser Sandbox nicht installierbar ist): Pfad-Ableitung und Schritt 1 (fill_missing_inputs) laufen korrekt durch, `input_lp.xlsx`/`ergebnis.xlsx` landen korrekt im Input-Ordner.

## Datei-Pfade in den Einzelskripten (Fallback für Standalone-Aufruf)
`fill_missing_inputs.py` und `battery_optimization.py` haben weiterhin eigene Default-Pfade für den Fall, dass sie direkt (nicht über den Orchestrator) aufgerufen werden: Default ist derselbe Ordner wie das jeweilige Skript, überschreibbar via Umgebungsvariable `BATTERIE_DATA_DIR`. Der Orchestrator (`run_batterieanalyse.py`) übersteuert das aber ohnehin immer mit expliziten Pfaden basierend auf der Nutzer-Auswahl.

## Last/Verbrauch
`Zeitreihen!I` (Last, in **kWh pro 15-Min-Intervall**, NICHT kW) + `Parameter!C43` ("Ja"/"Nein"). `battery_optimization.py::build_last_profile()`: C43=="Ja" -> Spalte I, umgerechnet in kW via `/DT_HOURS` (=0.25); C43=="Nein"/leer -> Last=0 überall.

## Sauber mit oemof.solph abbildbar
Zeitindex/Auflösung, PV-Produktionsprofil (Source mit fix-Profil, Mapping siehe unten), Batterie-Grundmodell (`GenericStorage`, Start/End-SOC=50%), Trafo-Wirkungsgrad (`Converter`), Investitionskosten/Lebenszeit (Annuität), Vollzyklen-Durchsatzbegrenzung (`generic_integral_limit`, Einheiten korrekt MWh/Jahr -> kWh), Last/Maximale Einspeisung/Maximaler Bezug als `nominal_value`/`fix` auf Flows.

## Braucht Zusatzarbeit ausserhalb oemof
SwissIX via ENTSO-E API (UTC→CET fix UTC+1, EUR/MWh→CHF/kWh), HT/NT-Tarifschema + SPOT-Auf-/Abschlag (pandas-Vorverarbeitung), PV-Referenzprofil "CH_PV_normiert_1kWp" (noch offen, braucht Standort/Ausrichtung, Kommentar "mail Urs/Astrid").

## PV-Mapping (Parameter!C17)
`"normiert"` -> Zeitreihen!E * DC Leistung (C16). `"Profil_Zeitreihe"` -> Zeitreihen!F direkt.

## Peak-Shaving — automatisch aus Parameter!C9
Aktiv wenn C9 ≠ 0/NaN, sonst inaktiv (`determine_peakshaving_aktiv()`). Aktueller Wert: C9=13 -> aktiv.

## SRL-Variante — automatisch aus Parameter!C38
`"Ja_residual"` -> `"post_hoc"`, `"Ja_optimiert"` -> `"optimiert"`, `"Nein"` -> `"keine"`. **Achtung:** C38 stand zuletzt noch auf altem Wert `"Ja"` — muss von Beat aktualisiert werden.

## "Maximaler Bezug" (Parameter!C35)
Absolutes Leistungslimit für gesamten Netzbezug = Last-Netzbezug + Batterie-Netzladung ZUSAMMEN.

## Modell-Topologie
Netzbezug-Bus (ein Flow vom Netz, `nominal_value=C35`, zwei Converter-Abgänge zu Last-Bus und Ladebus), Ladebus (dedizierter Zusammenführungspunkt PV-/Netz-Ladeflow, exakte Herkunftsverfolgung), Entladebus (Batterie→Hausbus). Trafo-Wirkungsgrad vereinfachend multiplikativ mit Lade-/Entladewirkungsgrad verrechnet.

## SRL-Modul-Details
- **"post_hoc":** Optimierung ohne SRL, danach Modul-Erträge (Peak-Shaving direkt, Eigenverbrauch/Arbitrage via "gut durchmischter Tank"-Konvention inkl. drittem Anteil "Anfangsbestand"). SRL-Mehrertrag additiv, nur bewertet wenn `soc_min_srl < SOC/Kapazität < soc_max_srl` (Default 10%/90%).
- **"optimiert":** `r_pos[t]`/`r_neg[t]`, nur Leistungsheadroom, Erlösterm direkt in Zielfunktion.

## Skripte (alle an Beat gesendet, aktuellster Stand)
1. `fill_missing_inputs.py` — Input aufbereiten (SwissIX, Tarife, SRL-Preise backen).
2. `battery_optimization.py` — vollständiges oemof.solph-Modell. **In dieser Sandbox nicht end-to-end testbar** (oemof.solph/pyomo/Solver nicht installierbar) — Beat testet lokal.
3. `run_batterieanalyse.py` + `run_batterieanalyse.bat` — Einstiegspunkt für alle Fleco-Nutzer (Dateiauswahl-Dialog, Drag&Drop, Doppelklick).
4. `Vorschlag_Peak-Shaving_und_SRL.md` — Herleitung/Formeln (v3).

## Offene Punkte
- C38 im Excel noch auf altem Wert "Ja" — muss aktualisiert werden.
- `requirements.txt` bewusst zurückgestellt (macht Beat am Schluss des Projekts).
- ENTSO-E-API-Key ist als Fallback-Default im Code hinterlegt — Sicherheitshinweis gegeben, noch nicht behoben (sollte irgendwann auf reine Umgebungsvariable ohne Code-Fallback umgestellt werden).
- PV-Referenzprofil (Standort/Ausrichtung) weiterhin offen.
- Format von `ergebnis.xlsx` ist nur ein Platzhalter — wird später gemeinsam definiert.

## Status
Vollständiger Workflow inkl. benutzerfreundlichem Einstiegspunkt (Doppelklick/Dateidialog) für alle Fleco-Nutzer geschrieben und gesendet. Wartet auf lokalen Testlauf durch Beat (oemof.solph in dieser Sandbox nicht verfügbar).