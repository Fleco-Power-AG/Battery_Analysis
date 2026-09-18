"""
battery_optimization.py
========================

Batterie/PV/Netz-Optimierung mit oemof.solph, aufbauend auf `input_lp.xlsx`
(siehe `input_prep.py`).

WICHTIGER HINWEIS ZUR TESTABDECKUNG
------------------------------------
Dieses Skript wurde OHNE lauffaehige oemof.solph-/Pyomo-/Solver-Installation
geschrieben: die Cloud-Sandbox, in der es entstanden ist, kann nur auf
PyPI-Paket-Registries zugreifen, aber `oemof.solph` und `pyomo` waren darueber
nicht installierbar (kein Solver sowieso nicht). Der Code ist so sorgfaeltig
wie moeglich gegen die offizielle oemof.solph-Dokumentation und den
Quellcode-Ausschnitten von `oemof.solph.constraints.generic_integral_limit`
geschrieben, aber NICHT End-to-End getestet. Bitte in eurer lokalen Umgebung
(mit oemof.solph 0.6.4 + Solver `cbc`, Standard-Setup via Conda -- siehe
SOLVER-Konstante unten und Setup-Hinweis) laufen lassen und mir
Fehlermeldungen zurueckmelden -- dann kann ich gezielt nachbessern.
Kandidaten fuer API-Abweichungen je nach genauer oemof-Version, auf die beim
Testen besonders zu achten ist (Kommentare an den jeweiligen Stellen im Code):
  - `Flow(custom_properties={...})` fuer generic_integral_limit
  - `om.objective.expr = om.objective.expr + ...` zum Erweitern der
    Zielfunktion (siehe main(): NUR NOCH EIN EINZIGES Mal aufgerufen, NICHT
    mehr mehrfach `+=` -- Root-Cause-Fix-VERSUCH eines echten CBC-Absturzes
    "duplicates in objective and matrix" / "termination condition: unknown".
    URSPRUENGLICH (10.9.2026) auf die Peak-Shaving/SRL-"optimiert"-
    Zielfunktions-Erweiterung zurueckgefuehrt; Beats Solver-Log vom 15.9.2026
    (Peak-Shaving UND SRL "optimiert" beide INAKTIV, trotzdem derselbe Fehler
    mit derselben Zahl 105119) zeigt aber, dass die Verdopplung schon in
    OEMOFS EIGENER, unveraenderter Zielfunktion entsteht. Der Quicksum-
    Rebuild-Schritt laeuft deshalb jetzt IMMER, nicht mehr nur, wenn Peak-
    Shaving/SRL tatsaechlich einen Zusatzterm liefern -- siehe ausfuehrlicher
    Kommentar in main(). Noch nicht von Beat bestaetigt/end-to-end getestet.)
  - exakte Attributnamen fuer Ergebnis-Zugriff (`processing.results`)

Modellumfang
------------
- PV (Source, fixes Profil) + Batterie (GenericStorage) + Netz, mit:
  - PV-Profil je nach Parameter!C26: "normiert" -> Zeitreihen!E (CH_PV_normiert_
    1kWp) * DC Leistung (C25); "Profil_Zeitreihe" -> Zeitreihen!F (PV_Profil)
    direkt (siehe `build_pv_profile()`).
  - "Netzbezug-Bus" als gemeinsamer Netzanschlusspunkt: EIN Zufluss vom Netz
    mit hartem Limit `Maximaler Bezug` (Parameter!C44) fuer Last+Batterie
    zusammen (siehe Klaerung im Gespraech).
  - Zusaetzliche, batterie-spezifische Netzlimiten `Maximale Einspeisung
    Batterie` (Parameter!C45) und `Maximaler Bezug Batterie` (Parameter!C46),
    getrennt von den generischen Standortlimiten C43/C44. `Maximale
    Einspeisung Batterie` (kombiniert mit der allgemeinen Batterieleistung
    C30) sitzt NUR auf dem Export-Abgang `link_batt_export`
    (entladebus_zu_export) -- NICHT auf `discharge_flow` selbst, das nur
    noch mit der allgemeinen Batterieleistung C30 gekappt ist (Root-Cause-
    Fix vom 11.9.2026, siehe Bugfix-Log: C45=0 bedeutete vorher "Batterie
    komplett stillgelegt" statt "kein Netz-Export ab Batterie", was nicht
    Beats Absicht entsprach). `Maximaler Bezug Batterie` kappt
    `link_netz_charge` (Netzladung der Batterie) -- siehe
    `build_energysystem()` fuer die genaue Modellierungs-Annahme und deren
    Grenzen.
  - NEU (Beats Erweiterung: PV und Batterie koennen unterschiedlich
    vermarktet werden, Parameter!C14-C22, siehe input_prep.py): der Export-
    Pfad ist in zwei EIGENE, unterschiedlich bepreiste Zufluesse in einen
    gemeinsamen `b_export`-Bus aufgeteilt -- `link_pv_export` (PV direkt,
    Zeitreihen!C-Preis) und `link_batt_export` (aus der Batterie, Zeitreihen!K-
    Preis) -- `b_ac` ist seither NUR NOCH der Last-Bus. `b_export` selbst
    erzwingt nur noch die standortweite Einspeisegrenze (`Maximale
    Einspeisung`, C43), ungepreist.
  - "Ladebus" als dedizierter Zusammenfuehrungspunkt fuer PV- und
    Netz-Ladeflow der Batterie -> exakte Herkunftsverfolgung (PV vs. Netz)
    der Batterieladung, ohne dass GenericStorage mehrere Input-Buse braucht
    (GenericStorage erlaubt nur GENAU einen Input- und einen Output-Bus).
  - Trafo-Wirkungsgrad (C22) wird vereinfachend MULTIPLIKATIV mit
    Lade-/Entladewirkungsgrad (C23/C24) verrechnet (kombinierte Wirkungsgrade
    auf `inflow_conversion_factor`/`outflow_conversion_factor`), statt einen
    eigenen `Link`-Bidirektional-Baustein einzufuehren. Wenn ihr die
    Trafo-Verluste separat ausgewiesen haben wollt, sagt Bescheid, das lässt
    sich nachruesten.
  - Last(t) je nach Parameter!C45: "Ja" -> Zeitreihen!I (Last, in kWh PRO
    15-MIN-INTERVALL -- Umrechnung in kW via `/DT_HOURS`, siehe
    `build_last_profile()`), "Nein"/leer -> Last = 0 ueberall (Modell
    funktioniert auch ganz ohne Verbrauch).
  - Start-/End-SOC = 50% (`initial_storage_level=0.5`, `balanced=True`).
  - Vollzyklen-Durchsatzbegrenzung (C29) ueber
    `oemof.solph.constraints.generic_integral_limit`.
  - NEU (Beats Erweiterung 11.9.2026, Parameter!C27 "Abschaltung bei
    Negativpreisen"): steuert, ob die PV-Anlage abgeregelt/gedrosselt werden
    DARF (z.B. um Einspeisung bei negativen Preisen zu vermeiden). "Ja"
    (Default, auch wenn C27 fehlt -- Rueckwaertskompatibilitaet mit
    aelteren Inputs.xlsx-Dateien) -> bisheriges Verhalten, der Abregelungs-
    Sink `snk_curtailment` bleibt frei nutzbar. "Nein" -> `snk_curtailment`
    wird auf 0 gekappt, die GESAMTE PV-Erzeugung MUSS zwingend ueber
    Eigenverbrauch/Export/Batterieladung abfliessen, auch bei negativen
    Preisen -- siehe `determine_pv_abschaltung_erlaubt()` und die neue
    Vorab-Pruefung in `check_feasibility_preconditions()` (ohne
    Abregelungs-Sicherheitsventil kann ein PV-Ueberschuss ueber die
    kombinierte Kapazitaet von link_pv_ac/link_pv_export/link_pv_charge
    hinaus wieder zu "termination condition: infeasible" fuehren).
- Peak-Shaving als custom Pyomo-Block, AUTOMATISCH aktiv/inaktiv je nach
  Parameter!C9 (Netznutzung Leistung): Wert vorhanden und != 0 -> aktiv,
  0/leer -> inaktiv (siehe `determine_peakshaving_aktiv()`). Hilfsvariable
  `peak[monat] >= Netzbezug-Flow(t)`, Kostenterm in Zielfunktion.
- SRL-Vermarktung, Variante AUTOMATISCH aus Parameter!C40 abgeleitet
  (siehe `determine_srl_variante()`):
  - C40 == "Ja_Residual" -> "post_hoc": NICHT Teil der Optimierung. Nach der
    Loesung wird die freie Lade-/Entladeleistung mit den SRL-Preisen bewertet
    (nur wenn SOC_MIN_SRL < SOC/Kapazitaet < SOC_MAX_SRL, beide waehlbar).
  - C40 == "Ja_Optimiert" -> "optimiert": Teil der Optimierung (custom
    Pyomo-Block, NUR Leistungsheadroom, keine SOC-/Energie-Nachhaltigkeitskopplung).
  - C40 == "Nein" -> "keine": kein SRL-Modul.
- Ertragsaufteilung nach der Optimierung in die Module Eigenverbrauch,
  Arbitrage, Peak-Shaving, SRL -- ueber eine "gut durchmischter Tank"-
  Konvention (PV-/Netz-/Anfangsbestand-Anteil des Speicherinhalts wird pro
  Zeitschritt proportional fortgeschrieben; der Anfangsbestand-Anteil wird
  intern weiter mitgefuehrt, damit die anfaengliche 50%-SOC-Fuellung nicht
  faelschlich als Eigenverbrauch/Arbitrage-Ertrag erscheint -- aber NICHT
  mehr als eigenes Modul ausgewiesen, auf Beats Wunsch).

Benoetigte Pakete: oemof.solph, pyomo, pandas, numpy, openpyxl, ein Solver
(z.B. `pip install pyomo oemof.solph pandas numpy openpyxl` + CBC/GLPK
systemweit installiert).
"""

from __future__ import annotations

import datetime as dt
import os
import sys
import warnings

import numpy as np
import pandas as pd
import pyomo.environ as po
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# NEU (Beat, 18.9.2026: "wie werde ich das UserWarning los?"): openpyxl meldet
# beim Einlesen JEDER .xlsx-Datei mit Dropdown-Validierungen (Data Validation,
# z.B. Beats "Ja"/"Nein"-Felder) diese harmlose Warnung -- die Validierung
# selbst bleibt beim reinen Lesen unberuehrt, nur openpyxls eigene Faehigkeit,
# sie beim SPEICHERN zu erhalten, ist eingeschraenkt (betrifft hier nicht:
# battery_optimization.py speichert keine Excel-Datei). Bewusst NUR diese eine,
# namentlich passende Meldung unterdrueckt (kein pauschales `ignore` aller
# UserWarnings), damit andere, tatsaechlich relevante Warnungen sichtbar bleiben.
warnings.filterwarnings(
    "ignore", message=r".*Data Validation extension is not supported.*", category=UserWarning
)

import oemof.solph as solph

# --------------------------------------------------------------------------
# Konfiguration
# --------------------------------------------------------------------------

# Absoluter Datenordner statt relativer Pfade -- so muss die Excel-Datei NICHT
# im Git-Repo liegen (gleiche Konvention wie in input_prep.py: Default
# ist derselbe Ordner wie dieses Skript, per BATTERIE_DATA_DIR umbiegbar).
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.environ.get("BATTERIE_DATA_DIR", SCRIPT_DIR)

INPUT_PATH = os.path.join(DATA_DIR, "input_lp.xlsx")
SOLVER = "cbc"  # Standard-Setup fuer alle Fleco-Nutzer: Conda-Umgebung mit
# `conda install -c conda-forge coincbc` (installiert cbc.exe inkl. aller
# noetigen DLLs direkt in die Umgebung -- kein separater Windows-Installer,
# kein manueller PATH-Eintrag, kein Reboot noetig, solange die Conda-Umgebung
# aktiviert ist, bevor dieses Skript laeuft). Alternativen: "glpk", "gurobi",
# "cplex" ... oder "appsi_highs" (via `pip install highspy`, funktioniert auch
# ganz ohne Conda in einer normalen venv, falls das mal gebraucht wird).

# 15-Minuten-Zeitschritte -- an mehreren Stellen gebraucht (Last-Umrechnung
# kWh->kW, SRL-Erloesberechnung, Ertragsaufteilung), daher zentral definiert.
DT_HOURS = 0.25

# Peak-Shaving und SRL-Variante werden standardmaessig automatisch aus dem
# Excel abgeleitet (Parameter!C9 bzw. C40 -- siehe determine_peakshaving_aktiv()
# und determine_srl_variante() unten). Diese beiden Konstanten sind nur noch
# Fallback-Werte, falls man `main()` mit einer expliziten Zeitreihe ohne Excel-
# Bezug aufruft; im Normalfall braucht ihr sie nicht anzufassen.
SOC_MIN_SRL = 0.10          # nur fuer SRL-Variante "post_hoc", waehlbar
SOC_MAX_SRL = 0.98          # nur fuer SRL-Variante "post_hoc", waehlbar

# NEU (auf Beats Wunsch, siehe Gespraech vom 2025-08-21 zu "wieso laedt/
# entlaedt die Batterie gleichzeitig"): kleiner Durchsatz-/Degradationskosten-
# Term (CHF/kWh, auf die ENTLADENE Energie), analog zur ueblichen Praxis in
# Batterie-Arbitrage-Modellen (typische Literaturwerte 1-3 Rp/kWh fuer
# Li-Ion-Zyklenalterung). Grund: ohne diesen Term ist die Optimierung
# zwischen "PV deckt Last direkt" und "Batterie entlaedt (aus Netzladung),
# PV wird abgeregelt" bei bestimmten Preiskonstellationen (z.B. stark
# negativer Bezugstarif) wirtschaftlich EXAKT gleichwertig -- der Solver
# waehlt dann bei mehreren gleich guten Loesungen (LP-Entartung) eine davon
# aus, teils die energetisch schlechtere (unnoetiger Round-Trip-Verlust,
# PV wird abgeregelt statt direkt genutzt). Ein kleiner Kostenanteil pro
# entladener kWh macht "unnoetiges Zyklieren" strikt teurer als PV-
# Direktnutzung und loest die Entartung auf, OHNE echte Arbitrage/Peak-
# Shaving/SRL-Faelle zu verhindern (die bleiben weiterhin klar profitabler
# als dieser kleine Kostenanteil). Auf 0.0 setzen, um wie bisher OHNE
# Degradationskosten zu rechnen.
DEGRADATIONSKOSTEN_CHF_PRO_KWH = 0.01  # 1 Rp/kWh entladene Energie (Standardwert)

# Siehe input_prep.py: Excel zaehlt Zeit als fixen UTC+1-Offset,
# keine Sommerzeit-Umstellung.
CET_FIXED = dt.timezone(dt.timedelta(hours=1))

# Spaltenlayout im Sheet "Zeitreihen" (wie in input_prep.py)
COL_ZEIT = 1
COL_SWISSIX = 2
COL_RUECKLIEFER = 3  # PV-Rueckliefertarif
COL_BEZUG = 4
COL_PV_NORMIERT = 5
COL_PV_PROFIL = 6
COL_SRL_NEG = 7
COL_SRL_POS = 8
COL_LAST = 9  # neu: Last/Verbrauch in kWh pro 15-Min-Intervall (nicht kW!)
# NEU (Beats Wunsch: eigener Rueckliefertarif fuer die Batterie, siehe
# input_prep.py::COL_RUECKLIEFER_BATTERIE) -- Spalte K, von input_prep.py
# beim Erzeugen von input_lp.xlsx befuellt.
COL_RUECKLIEFER_BATTERIE = 11


# --------------------------------------------------------------------------
# 1. Input einlesen
# --------------------------------------------------------------------------

def _check_label(ws_param, row: int, expected) -> None:
    """Wie input_prep.py::_check_label() -- prueft Parameter!A<row> gegen den
    erwarteten Feldnamen, damit ein umgebautes Parameter-Sheet (Zeilen
    verschoben/eingefuegt, ist in diesem Projekt schon mehrfach passiert)
    einen klaren Fehler statt stiller Fehlwerte ergibt."""
    actual = ws_param[f"A{row}"].value
    actual_norm = str(actual).strip().lower() if actual is not None else ""
    erwartet = [expected] if isinstance(expected, str) else list(expected)
    if actual_norm not in [e.strip().lower() for e in erwartet]:
        raise RuntimeError(
            f"Parameter!A{row} = {actual!r}, erwartet {erwartet!r}. Das "
            "Parameter-Sheet-Layout hat sich vermutlich veraendert -- bitte "
            "die Zeilennummern in battery_optimization.py::read_inputs() "
            "(und input_prep.py::read_params()) an das neue Layout anpassen."
        )


def read_inputs(path: str = INPUT_PATH):
    wb = load_workbook(path, data_only=True)
    wsP = wb["Parameter"]
    wsZ = wb["Zeitreihen"]

    # NEU (Beats Umbau: Quartals-/Saisonpreise, mehrere HT-Fenster, getrennte
    # PV-/Batterie-Rueckliefertarife) -- Zeilen 6-22 pruefen, da hier am
    # meisten verschoben/eingefuegt wurde. Zeilen ab 25 pruefen wir nur an
    # den Ecken (25/26 und 49/50/54), das genuegt um einen Verrutscher zu
    # erkennen, ohne jede einzelne Zeile hier zu duplizieren.
    for row, label in [
        (5, "Tarifschema"), (6, "HT"), (7, "NT"), (8, "Netznutzung Leistung"),
        (9, "Tage HT"), (10, "Startzeit HT"), (11, "Endzeit HT"),
        (12, "Spot_Abschlag_Lieferung"), (13, "Spot_Aufschlag_Bezug"),
        (14, "Rücklieferung_PV"), (15, "HKN"), (16, "Floor"), (17, "Fixtarif"),
        (19, "Rücklieferung_Batterie"), (20, "HKN"), (21, "Floor"), (22, "Fixtarif"),
        (25, "DC Leistung"), (26, "Produktions-Input"),
        # NEU (Beats Erweiterung 11.9.2026): C27 war bisher eine leere
        # Reserve-Zeile zwischen "Produktions-Input" (C26) und dem
        # "Batterie"-Abschnitt (C28) -- jetzt belegt mit "Abschaltung bei
        # Negativpreisen". Mitgeprueft, damit ein Verrutscher hier ebenso
        # auffaellt wie bei den anderen Eckwerten.
        (27, "Abschaltung bei Negativpreisen"),
        (49, "SRL-Leistung"), (50, "SRL-Leistung-Preise"), (54, "Lastgang"),
    ]:
        _check_label(wsP, row, label)

    params = {
        "beginn": wsP["C2"].value,
        "ende": wsP["C3"].value,
        # NEU (Beats Wunsch nach einer "wichtige Inputs"-Uebersichtsseite im
        # PDF-Report): bisher hat battery_optimization.py nur die Zellen
        # gelesen, die die Optimierung selbst braucht -- Tarifschema (C5) und
        # der SPOT-Auf-/Abschlag (C12/C13) wurden bisher NUR von
        # input_prep.py gelesen (dort schon zur Tarif-Berechnung gebraucht).
        # Da input_prep.py beim Speichern von input_lp.xlsx das GESAMTE
        # Parameter-Sheet uebernimmt (nur Formelzellen werden "gebacken",
        # siehe bake_formula_cells()), sind diese Zellen in der von
        # battery_optimization.py gelesenen Datei trotzdem vorhanden -- hier
        # nur zu reinen Anzeigezwecken (Report) mitgelesen, NICHT fuer die
        # Optimierung selbst gebraucht (die fertig berechneten Tarife stehen
        # bereits in Zeitreihen!C/D/K).
        "tarifschema": wsP["C5"].value,
        # NEU (Beats Wunsch nach Transparenz, ob SwissIX automatisch von der
        # API geladen oder von Hand in Zeitreihen!B eingetragen wurde): wird
        # von input_prep.py in diese bisher ungenutzte Zelle geschrieben,
        # siehe dort. None, falls kein SPOT-Schema bzw. keine SwissIX-Werte.
        "swissix_quelle": wsP["E5"].value,
        "ht_preis": wsP["C6"].value,
        "nt_preis": wsP["C7"].value,
        "netznutzung_leistung": wsP["C8"].value,
        "tage_ht": wsP["C9"].value,
        "start_ht": wsP["C10"].value,
        "end_ht": wsP["C11"].value,
        "spot_abschlag_lieferung": wsP["C12"].value,
        "spot_aufschlag_bezug": wsP["C13"].value,
        # NEU (Beats Erweiterung: PV und Batterie koennen unterschiedlich
        # vermarktet werden -- eigenes Schema je Quelle, siehe input_prep.py::
        # resolve_rueckliefertarif()). Nur zu Anzeigezwecken mitgelesen; die
        # fertig berechneten CHF/kWh-Werte stehen bereits in Zeitreihen!C/K.
        "rueckliefer_pv_schema": wsP["C14"].value,
        "rueckliefer_pv_hkn": wsP["C15"].value,
        "rueckliefer_pv_floor": wsP["C16"].value,
        "rueckliefer_pv_fixtarif": wsP["C17"].value,
        "rueckliefer_batterie_schema": wsP["C19"].value,
        "rueckliefer_batterie_hkn": wsP["C20"].value,
        "rueckliefer_batterie_floor": wsP["C21"].value,
        "rueckliefer_batterie_fixtarif": wsP["C22"].value,
        "dc_leistung": wsP["C25"].value,
        "produktions_input": wsP["C26"].value,
        # NEU (Beats Erweiterung 11.9.2026): "Ja"/"Nein"-Dropdown, ob die
        # PV-Anlage abgeregelt/gedrosselt werden darf (z.B. um Einspeisung
        # bei negativen Preisen zu vermeiden). Rohwert hier nur mitgelesen;
        # die eigentliche Ja/Nein-Auswertung (inkl. Rueckwaertskompatibilitaet
        # fuer aeltere Dateien ohne C27) macht `determine_pv_abschaltung_erlaubt()`.
        "pv_abschaltung_erlaubt": wsP["C27"].value,
        "kapazitaet": wsP["C29"].value,
        "leistung": wsP["C30"].value,
        "trafo_wirkungsgrad": wsP["C31"].value,
        "lade_wirkungsgrad": wsP["C32"].value,
        "entlade_wirkungsgrad": wsP["C33"].value,
        "invest_kosten_kwh": wsP["C34"].value,
        "unterhalt_kosten": wsP["C35"].value,
        "lebenszeit": wsP["C36"].value,
        "vollzyklen": wsP["C37"].value,
        "ladegrenze_soc": wsP["C39"].value / 100.0,
        "entladegrenze_soc": wsP["C40"].value / 100.0,
        "max_einspeisung": wsP["C43"].value,
        "max_bezug": wsP["C44"].value,
        # NEU (Beats zwei zusaetzliche Netz-Parameter): eigene Grenzwerte fuer
        # den Netzaustausch DER BATTERIE (separat von den generischen
        # Standortlimiten C43/C44, die Last+Batterie zusammen betreffen).
        # Siehe build_energysystem() fuer die genaue Verwendung.
        "max_einspeisung_batterie": wsP["C45"].value,  # Batterie -> Netz (Entladung)
        "max_bezug_batterie": wsP["C46"].value,        # Netz -> Batterie (Ladung)
        "srl_teilnahme": wsP["C49"].value,
        # NEU (Beats Wunsch: SRL-Preisschema in der Eingabeparameter-Uebersicht
        # anzeigen, z.B. "Backcast_2025_Kunde") -- bisher nur als Basis fuer
        # die in Zeitreihen!G/H bereits gebackenen SRL-Preise verwendet, aber
        # nie als eigener Wert mitgelesen/angezeigt.
        "srl_preisschema": wsP["C50"].value,
        "last_vorhanden": wsP["C54"].value,  # neu: "Ja" -> Zeitreihen!I, "Nein" -> keine Last
    }

    # Durchsatzbegrenzung/Jahr (C38 ist im Original eine Formel; falls der
    # gecachte Wert beim Excel-Resave verlorengegangen ist, hier neu herleiten)
    durchsatz = wsP["C38"].value
    if not isinstance(durchsatz, (int, float)):
        durchsatz = params["kapazitaet"] * params["vollzyklen"] / params["lebenszeit"]
    params["durchsatz_mwh_jahr"] = durchsatz

    # Unterhaltkosten (C35) ist seit Beats letzter Aenderung eine Formel
    # (=C29*C34*0.01, 1% der Kapitalkosten) statt eines festen Werts -- das
    # hat schon einmal zu einem Absturz gefuehrt (`TypeError: bad operand
    # type for unary -: 'NoneType'` in capital_costs()/main()), weil openpyxl
    # beim Speichern von input_lp.xlsx durch input_prep.py den gecachten
    # Formel-Wert verliert. ROOT-CAUSE-FIX ist in input_prep.py::main()
    # (bake_formula_cells() "backt" ALLE Formel-Zellen im Parameter-Sheet vor
    # dem Speichern in feste Werte) -- diese Pruefung hier ist NUR ein
    # zusaetzliches Sicherheitsnetz mit einer klaren Fehlermeldung, falls
    # trotzdem einmal eine input_lp.xlsx ohne diesen Fix verwendet wird.
    if not isinstance(params["unterhalt_kosten"], (int, float)):
        raise RuntimeError(
            f"Parameter!C35 (Unterhaltkosten) ist kein Zahlenwert "
            f"({params['unterhalt_kosten']!r}) -- vermutlich eine Formel-Zelle, "
            "deren gecachter Wert beim Speichern verlorenging. Bitte "
            "input_prep.py (neueste Version) erneut auf die urspruengliche "
            "Inputs.xlsx anwenden, um input_lp.xlsx neu zu erzeugen."
        )

    n_rows = wsZ.max_row
    zeit_raw = [wsZ.cell(row=r, column=COL_ZEIT).value for r in range(2, n_rows + 1)]

    # NEU: analog zum Fix in input_prep.py::main() (gefunden ueber Beats
    # "Inputs_BKW_Q2.xlsx") -- manche Szenario-Dateien haben nur einen TEIL
    # des 35040-Zeilen-Jahres-Templates mit echten Zeitstempeln befuellt (z.B.
    # nur Q2), waehrend andere Spalten (Last, SRL-Preise) ueber das gesamte
    # Template hinweg Werte haben koennen. Ohne diesen Fix wuerde unten ein
    # Zeitraster ueber die volle Blattlaenge (35040 Zeilen) angenommen und
    # weit ueber die tatsaechlichen Daten hinaus extrapoliert. Fix: nur die
    # FUEHRENDEN, LUECKENLOS befuellten Zeit-Zellen zaehlen. Bei
    # vollstaendig befuellten (Voll-Jahr-)Dateien aendert das nichts.
    n_valid = next((i for i, v in enumerate(zeit_raw) if v is None), len(zeit_raw))
    if n_valid < len(zeit_raw):
        print(
            f"  HINWEIS: Zeitreihen!A ist nur fuer {n_valid} von {len(zeit_raw)} "
            f"Zeilen befuellt -- Szenario wird auf diese Laenge begrenzt."
        )
        zeit_raw = zeit_raw[:n_valid]
        n_rows = n_valid + 1

    # Sicherheitsnetz analog zum Root-Cause-Fix in input_prep.py::main(): auch
    # hier den Zeitindex als exaktes 15-Minuten-Raster ab dem ersten
    # Zeitstempel NEU erzeugen statt die (potenziell Excel-Fliesskomma-
    # drift-behafteten) rohen Zellwerte zu uebernehmen. input_lp.xlsx sollte
    # das dank input_prep.py bereits driftfrei liefern -- diese Zeile
    # schuetzt zusaetzlich den Fall, dass battery_optimization.py direkt auf
    # einer aelteren/anderen Zeitreihen-Datei aufgerufen wird.
    zeit = pd.date_range(start=zeit_raw[0], periods=len(zeit_raw), freq="15min")
    index = pd.DatetimeIndex(zeit).tz_localize(CET_FIXED)

    # NEU (Beat, 18.9.2026, Absturz "ValueError: could not convert string to
    # float: '#N/A'" bei Inputs_BAT_PV_ohneLaden.xlsx): Zeitreihen-Spalten
    # koennen einen rohen Excel-Fehlerwert (z.B. '#N/A' aus einer noch nicht
    # aufloesbaren XLOOKUP-Formel) enthalten, dessen gecachter Wert beim
    # Baken (input_prep.py::bake_formula_cells()) als LITERALER STRING
    # uebernommen wird -- pd.Series(..., dtype=float) stuerzt darauf bisher
    # mit einer kryptischen, spaltenlosen Fehlermeldung ab. Fix: VOR der
    # float-Konvertierung explizit auf bekannte Excel-Fehlerwerte pruefen und
    # eine klare Meldung werfen, die Spalte/Zeile nennt (plus einen
    # spezifischen Hinweis fuer die PV-Normiert-Spalte, siehe
    # fetch_pv_reference_profile() in input_prep.py).
    _EXCEL_FEHLERWERTE = {"#N/A", "#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!"}

    def col(c):
        rohwerte = [wsZ.cell(row=r, column=c).value for r in range(2, n_rows + 1)]
        fehlerhafte = [
            (r, v) for r, v in zip(range(2, n_rows + 1), rohwerte)
            if isinstance(v, str) and v.strip().upper() in _EXCEL_FEHLERWERTE
        ]
        if fehlerhafte:
            spalte_buchstabe = get_column_letter(c)
            erste_zeile, erster_wert = fehlerhafte[0]
            hinweis = ""
            if c == COL_PV_NORMIERT:
                hinweis = (
                    " Das ist vermutlich das PV-Referenzprofil (Spalte E), das "
                    "input_prep.py bewusst NICHT automatisch befuellt, solange "
                    "Standort/Ausrichtung noch nicht final sind (siehe Excel-"
                    "Kommentar 'mail Urs/Astrid'). Sobald bekannt: input_prep."
                    "main(input_path=..., fetch_pv=True, pv_lat=..., pv_lon=...) "
                    "erneut aufrufen -- dann wird Spalte E automatisch mit einem "
                    "echten PVGIS-Profil befuellt."
                )
            raise RuntimeError(
                f"Zeitreihen!{spalte_buchstabe} enthaelt {len(fehlerhafte)} Excel-"
                f"Fehlerwert(e) (z.B. '{erster_wert}' in Zeile {erste_zeile}) statt "
                f"Zahlen -- kann nicht eingelesen werden.{hinweis}"
            )
        return pd.Series(rohwerte, index=index, dtype=float)

    zeitreihen = pd.DataFrame(
        {
            "swissix": col(COL_SWISSIX),
            "rueckliefertarif": col(COL_RUECKLIEFER),  # PV-Rueckliefertarif
            "bezugstarif": col(COL_BEZUG),
            "pv_normiert": col(COL_PV_NORMIERT),
            "pv_profil": col(COL_PV_PROFIL),
            "srl_neg": col(COL_SRL_NEG),
            "srl_pos": col(COL_SRL_POS),
            "last_kwh": col(COL_LAST),  # Achtung: kWh pro 15-Min-Intervall, NICHT kW
            # NEU (Beats Wunsch: eigener Rueckliefertarif fuer die Batterie):
            # von input_prep.py in Spalte K geschrieben (siehe COL_RUECKLIEFER_BATTERIE
            # oben). Faellt auf den PV-Rueckliefertarif zurueck, falls die Spalte in
            # einer aelteren input_lp.xlsx (vor diesem Feature) noch fehlt/leer ist --
            # entspricht dann dem alten Verhalten (ein gemeinsamer Rueckliefertarif).
            "rueckliefertarif_batterie": col(COL_RUECKLIEFER_BATTERIE),
        }
    )
    if zeitreihen["rueckliefertarif_batterie"].isna().all():
        print(
            "  HINWEIS: Zeitreihen!K (Rueckliefertarif_Batterie) ist leer -- "
            "verwende den PV-Rueckliefertarif (Spalte C) auch fuer die Batterie "
            "(altes Verhalten, z.B. bei einer input_lp.xlsx aus einer aelteren "
            "input_prep.py-Version)."
        )
        zeitreihen["rueckliefertarif_batterie"] = zeitreihen["rueckliefertarif"]
    return params, zeitreihen


def build_pv_profile(params: dict, zeitreihen: pd.DataFrame) -> pd.Series:
    """kW-Zeitreihe der PV-Erzeugung, je nach Parameter!C17.

    Gross-/Kleinschreibung und Leerzeichen am Rand werden ignoriert (siehe
    determine_srl_variante() fuer denselben Grund -- Excel-Eintraege variieren
    da leicht)."""
    wert = params["produktions_input"]
    normalisiert = str(wert).strip().lower() if wert is not None else ""

    if normalisiert == "normiert":
        # NEU (gefunden ueber Beats Meldung "C17 steht normiert drin, C16 ist
        # 0"): ist Parameter!C16 (DC-Leistung) 0 (bzw. None/NaN), gibt es in
        # diesem Szenario gar keine PV-Anlage -- das normierte Referenzprofil
        # (Zeitreihen!E) wird dann gar nicht gebraucht, jede Multiplikation
        # damit ergibt ohnehin 0 kW. Ohne diesen Fix loeste ein leeres Profil
        # hier faelschlich einen Absturz aus, obwohl PV im Szenario bewusst
        # abwesend ist (z.B. ein reines Last+Batterie-Szenario, bei dem
        # Parameter!C17 trotzdem auf 'normiert' stehen geblieben ist). Zusaetz-
        # lich wichtig: ein leeres Profil (NaN) mal 0 waere immer noch NaN,
        # nicht 0 -- deshalb hier explizit ein Null-Profil zurueckgeben statt
        # einfach "profile * dc_leistung" auch fuer diesen Fall zu rechnen.
        dc_leistung = params["dc_leistung"]
        dc_leistung_ist_null = dc_leistung is None or (
            isinstance(dc_leistung, (int, float)) and (
                (isinstance(dc_leistung, float) and np.isnan(dc_leistung)) or float(dc_leistung) == 0.0
            )
        )
        if dc_leistung_ist_null:
            return pd.Series(0.0, index=zeitreihen.index)
        profile = zeitreihen["pv_normiert"]
        if profile.isna().all():
            raise RuntimeError(
                "Parameter!C17='normiert' und Parameter!C16 (DC-Leistung) > 0, "
                "aber Zeitreihen!E (CH_PV_normiert_1kWp) ist leer. Bitte "
                "zuerst das PV-Referenzprofil ergaenzen (siehe "
                "input_prep.py::fetch_pv_reference_profile -- braucht "
                "Standort/Ausrichtung)."
            )
        return profile * dc_leistung
    elif normalisiert == "profil_zeitreihe":
        profile = zeitreihen["pv_profil"]
        if profile.isna().all():
            raise RuntimeError(
                "Parameter!C17='Profil_Zeitreihe', aber Zeitreihen!F (PV_Profil) ist "
                "leer. Bitte historische PV-Produktion ergaenzen."
            )
        return profile
    else:
        raise ValueError(
            f"Unbekannter Produktions-Input in Parameter!C17: "
            f"{wert!r} -- erwartet 'normiert' oder 'Profil_Zeitreihe'."
        )


def determine_peakshaving_aktiv(params: dict) -> bool:
    """Peak-Shaving ist aktiv, wenn Parameter!C9 (Netznutzung Leistung) einen
    Wert ungleich 0/NaN enthaelt -- sonst inaktiv. Automatisch aus dem Excel
    abgeleitet (kein manuelles Flag mehr noetig)."""
    wert = params["netznutzung_leistung"]
    if wert is None:
        return False
    try:
        return float(wert) not in (0.0,) and not np.isnan(float(wert))
    except (TypeError, ValueError):
        return False


def determine_srl_variante(params: dict) -> str:
    """Leitet die SRL-Variante aus Parameter!C40 ab (offizielle Werte im
    Hilfen_Listen-Dropdown: 'Ja_Residual', 'Ja_Optimiert', 'Nein'):
    'Ja_Residual' -> 'post_hoc', 'Ja_Optimiert' -> 'optimiert', 'Nein' -> 'keine'.
    Gross-/Kleinschreibung und Leerzeichen am Rand werden zusaetzlich ignoriert
    (falls doch mal jemand manuell 'ja_residual' o.ae. eintippt)."""
    mapping = {"ja_residual": "post_hoc", "ja_optimiert": "optimiert", "nein": "keine"}
    wert = params["srl_teilnahme"]
    normalisiert = str(wert).strip().lower() if wert is not None else ""
    if normalisiert not in mapping:
        raise ValueError(
            f"Parameter!C40 = {wert!r} ist kein bekannter Wert. Erwartet (Gross-/"
            f"Kleinschreibung egal) einen von 'Ja_Residual' (-> post_hoc-Bewertung), "
            f"'Ja_Optimiert' (-> Teil der Optimierung), 'Nein' (-> keine SRL-Vermarktung)."
        )
    return mapping[normalisiert]


def determine_pv_abschaltung_erlaubt(params: dict) -> bool:
    """Leitet aus Parameter!C27 ("Abschaltung bei Negativpreisen") ab, ob die
    PV-Anlage abgeregelt/gedrosselt werden DARF -- z.B. um Einspeisung bei
    negativen Preisen zu vermeiden, oder allgemein als Sicherheitsventil bei
    PV-Ueberschuss (siehe `snk_curtailment` in `build_energysystem()`).

    - "Ja" -> erlaubt (bisheriges Verhalten): `snk_curtailment` bleibt
      unbegrenzt und kostenlos nutzbar, die Optimierung waehlt Abregelung nur
      dann, wenn sie oekonomisch guenstiger ist als Eigenverbrauch/Export/
      Laden (z.B. bei negativem Rueckliefertarif) oder wenn sie als
      Sicherheitsventil noetig ist.
    - "Nein" -> NICHT erlaubt: `snk_curtailment` wird auf 0 gekappt. Die
      GESAMTE PV-Erzeugung MUSS dann ueber Eigenverbrauch/Export/
      Batterieladung abfliessen, auch wenn das wirtschaftlich nachteilig ist
      (z.B. Export bei negativem Preis). Kann das Modell PV-Ueberschuss in
      einem Zeitschritt dadurch nicht mehr loswerden, wird das Modell fuer
      diesen Zeitschritt "infeasible" -- siehe die entsprechende Vorab-
      Pruefung in `check_feasibility_preconditions()`.
    - Zelle leer/fehlt (aeltere Inputs.xlsx-Dateien ohne C27) -> "Ja"
      (Rueckwaertskompatibilitaet: das ist das bisherige, immer-erlaubte
      Verhalten von `snk_curtailment`).
    """
    wert = params.get("pv_abschaltung_erlaubt")
    normalisiert = str(wert).strip().lower() if wert is not None else ""
    if normalisiert == "":
        return True
    if normalisiert == "ja":
        return True
    if normalisiert == "nein":
        return False
    raise ValueError(
        f"Parameter!C27 (Abschaltung bei Negativpreisen) = {wert!r} ist kein "
        f"bekannter Wert. Erwartet 'Ja' oder 'Nein' (Gross-/Kleinschreibung egal)."
    )


def build_last_profile(
    params: dict, zeitreihen: pd.DataFrame, last_series: pd.Series | None = None
) -> pd.Series:
    """kW-Zeitreihe der Last, je nach Parameter!C45.

    - `last_series` explizit uebergeben: nimmt Vorrang (z.B. fuer Tests), muss
      bereits in kW vorliegen.
    - Parameter!C45 == "Ja": Zeitreihen!I (Last) verwenden. WICHTIG: diese
      Spalte ist in kWh PRO 15-MIN-INTERVALL angegeben, nicht in kW -- oemof-
      Flows sind aber Leistungen (kW). Umrechnung: kW = kWh_intervall / DT_HOURS.
    - Parameter!C45 == "Nein" (oder leer): Last = 0 ueberall, das Modell
      funktioniert auch ganz ohne Verbrauch.
    """
    if last_series is not None:
        return last_series.reindex(zeitreihen.index).fillna(0.0)

    last_vorhanden = params.get("last_vorhanden")
    last_vorhanden_normalisiert = (
        str(last_vorhanden).strip().lower() if last_vorhanden is not None else ""
    )
    if last_vorhanden_normalisiert == "ja":
        last_kwh = zeitreihen["last_kwh"]
        if last_kwh.isna().all():
            raise RuntimeError(
                "Parameter!C45='Ja', aber Zeitreihen!I (Last) ist leer. Bitte "
                "Lastreihe ergaenzen oder C45 auf 'Nein' setzen."
            )
        return last_kwh / DT_HOURS  # kWh/15min -> kW

    return pd.Series(0.0, index=zeitreihen.index)


# --------------------------------------------------------------------------
# 2. Energiesystem aufbauen
# --------------------------------------------------------------------------

def build_energysystem(
    params: dict,
    zeitreihen: pd.DataFrame,
    pv_profile: pd.Series,
    last_profile: pd.Series,
):
    # oemof mag einen tz-naiven DatetimeIndex mit klarem `freq` lieber als
    # unseren fixen-Offset-tz-aware-Index -- CET ist hier ohnehin ein reiner
    # Namensgeber (fixer UTC+1-Offset, siehe input_prep.py), daher
    # unbedenklich, die tz-Info fuer oemof zu entfernen.
    naive_index = zeitreihen.index.tz_localize(None)
    solph_index = pd.date_range(
        start=naive_index[0], periods=len(naive_index) + 1, freq="15min"
    )
    # oemof braucht einen Zeitindex mit einem Zeitpunkt MEHR als Zeitschritte
    # (der letzte Punkt markiert das Ende des letzten Intervalls).

    es = solph.EnergySystem(timeindex=solph_index, infer_last_interval=False)

    n = len(naive_index)
    dt_hours = DT_HOURS  # 15-Minuten-Schritte (zentral definiert, siehe Konfiguration)

    # --- Busse -------------------------------------------------------------
    b_pv = solph.Bus(label="b_pv")
    b_netzbezug = solph.Bus(label="b_netzbezug")
    b_ladebus = solph.Bus(label="b_ladebus")   # dediziert: nur PV-/Netz-Ladeflow -> Batterie
    b_entladebus = solph.Bus(label="b_entladebus")  # dediziert: nur Batterie -> Hausbus
    b_ac = solph.Bus(label="b_ac")  # NEU: jetzt NUR NOCH Last-Bus (Export siehe b_export unten)
    # NEU (Beats Wunsch: eigener Rueckliefertarif fuer die Batterie, der vom
    # PV-Rueckliefertarif abweichen kann): b_ac diente bisher als GEMEINSAMER
    # Last-UND-Export-Bus fuer PV UND Batterie zusammen -- mit nur EINEM
    # Rueckliefertarif war das unbedenklich (jede exportierte kWh wurde
    # gleich bewertet, egal woher). Jetzt, wo PV und Batterie UNTERSCHIEDLICH
    # vermarktet werden koennen, muss der Export-Pfad vor dem Bepreisen nach
    # Quelle getrennt sein -- sonst koennte die Optimierung gar nicht auf den
    # Preisunterschied reagieren (und die Ertragsbuchhaltung waere falsch).
    # Deshalb: b_ac ist jetzt NUR der Last-Bus (Bilanz muss exakt last_profile(t)
    # ergeben, siehe snk_last unten); PV- und Batterie-Export laufen je ueber
    # einen EIGENEN, individuell bepreisten Pfad (link_pv_export/
    # link_batt_export) in den gemeinsamen b_export-Bus, der NUR noch die
    # STANDORTWEITE Einspeisegrenze (max_einspeisung) durchsetzt, aber selbst
    # ungepreist ist (der Preis sitzt schon an den beiden Zufluessen).
    b_export = solph.Bus(label="b_export")

    # --- Quellen -------------------------------------------------------------
    src_pv = solph.components.Source(
        label="pv",
        outputs={b_pv: solph.Flow(nominal_capacity=1, fix=pv_profile.values)},
    )

    src_netz = solph.components.Source(
        label="netz",
        outputs={
            b_netzbezug: solph.Flow(
                nominal_capacity=params["max_bezug"],
                variable_costs=zeitreihen["bezugstarif"].values,
            )
        },
    )

    # --- Senken --------------------------------------------------------------
    snk_last = solph.components.Sink(
        label="last",
        inputs={b_ac: solph.Flow(nominal_capacity=1, fix=last_profile.values)},
    )

    # NEU: b_export (nicht mehr b_ac, siehe Kommentar bei b_export oben) --
    # der Preis sitzt jetzt an link_pv_export/link_batt_export, hier nur noch
    # die standortweite Kappung (max_einspeisung), UNGEPREIST (sonst wuerde
    # der Preis doppelt gezaehlt).
    snk_export = solph.components.Sink(
        label="export",
        inputs={
            b_export: solph.Flow(nominal_capacity=params["max_einspeisung"])
        },
    )

    # PV-Abregelung (Curtailment): b_pv hat sonst NUR die drei Abgaenge
    # link_pv_ac (Kappe dc_leistung), link_pv_export (Kappe dc_leistung) und
    # link_pv_charge (Kappe leistung/Batterieladeleistung) -- die PV-Quelle
    # selbst ist aber `fix` (siehe src_pv oben), d.h. ihr gesamter Ertrag MUSS
    # zwingend ueber b_pv abfliessen. Falls PV-Erzeugung in einem Zeitschritt
    # die Summe dieser drei Kapazitaeten uebersteigt (z.B. kurze
    # Ueberschwinger > 1.0 in einem gemessenen/normierten Profil), gibt es
    # sonst KEINEN Abfluss dafuer -> die Busbilanz ist unloesbar
    # ("infeasible"), unabhaengig von allen anderen Parametern. Dieser Sink
    # faengt genau diesen Fall ab -- UND ist der Pfad, ueber den die
    # Optimierung PV bei negativen Rueckliefertarifen freiwillig abregelt
    # (kostenlos ist immer mindestens so gut wie ein negativer Erloes ueber
    # link_pv_export).
    #
    # NEU (Beats Erweiterung 11.9.2026, Parameter!C27 "Abschaltung bei
    # Negativpreisen", siehe `determine_pv_abschaltung_erlaubt()`): bisher
    # war dieser Sink IMMER unbegrenzt und kostenlos nutzbar. Jetzt nur noch,
    # wenn C27="Ja" (oder leer/fehlend, Rueckwaertskompatibilitaet). Bei
    # C27="Nein" wird dieser Sink auf 0 gekappt -- die Anlage darf dann NICHT
    # abgeregelt werden, auch nicht bei negativen Preisen oder PV-
    # Ueberschuss: die gesamte Erzeugung MUSS ueber die drei obigen Pfade
    # abfliessen (siehe Vorab-Pruefung dazu in
    # `check_feasibility_preconditions()`).
    pv_abschaltung_erlaubt = determine_pv_abschaltung_erlaubt(params)
    snk_curtailment = solph.components.Sink(
        label="pv_abregelung",
        inputs={
            b_pv: solph.Flow() if pv_abschaltung_erlaubt else solph.Flow(nominal_capacity=0)
        },
    )

    # --- Direktflüsse (unbegrenzt, ausser durch die Busse/Quellen selbst) ---
    # PV -> Hausbus (direkter Verbrauch/Export) und PV -> Ladebus (Batterieladung)
    # ueber je einen 1:1-Converter, damit b_pv/b_netzbezug sich sauber auf die
    # zwei Abgaenge (Last-Deckung vs. Batterieladung) aufteilen lassen.
    link_pv_ac = solph.components.Converter(
        label="pv_zu_ac",
        inputs={b_pv: solph.Flow()},
        outputs={b_ac: solph.Flow(nominal_capacity=params["dc_leistung"])},
        conversion_factors={b_ac: 1.0},
    )
    # NEU: direkter PV-Export-Pfad (b_pv -> b_export), bepreist mit dem
    # PV-Rueckliefertarif (Parameter!C14-C17, siehe input_prep.py). Bisher
    # lief PV-Export ueber link_pv_ac + den (jetzt entfernten) Preis an
    # snk_export -- siehe Kommentar bei b_export oben.
    link_pv_export = solph.components.Converter(
        label="pv_zu_export",
        inputs={b_pv: solph.Flow()},
        outputs={
            b_export: solph.Flow(
                nominal_capacity=params["dc_leistung"],
                variable_costs=-zeitreihen["rueckliefertarif"].values,
            )
        },
        conversion_factors={b_export: 1.0},
    )
    link_pv_charge = solph.components.Converter(
        label="pv_zu_ladebus",
        inputs={b_pv: solph.Flow()},
        outputs={b_ladebus: solph.Flow(nominal_capacity=params["leistung"])},
        conversion_factors={b_ladebus: 1.0},
    )
    # WICHTIG: Netzbezug->AC ist NICHT einfach mit max_bezug gekappt, sondern
    # zeitschritt-genau auf die aktuelle Last begrenzt (nominal_capacity=1,
    # maximum=last_profile.values, d.h. Fluss(t) <= last_profile(t)). Grund (aus
    # einem echten Testlauf entdeckt): mit einer flachen max_bezug-Kappe kann
    # die Optimierung bei NEGATIVEM Bezugstarif (kommt in den SwissIX-Daten
    # vor, siehe Arbitrage-Diagnose in main()) beliebig viel "gratis" Netz-
    # strom importieren (man wird fuers Beziehen bezahlt) und diesen ueber
    # b_ac wieder exportieren -- das verdraengt echte PV-Erzeugung aus der
    # gemeinsamen Einspeise-Kapazitaet (max_einspeisung) in die Abregelung
    # (snk_curtailment) und blaeht nebenbei die monatliche Netzbezugsspitze
    # (und damit die Peak-Shaving-Kosten) kuenstlich auf. Physikalisch/
    # vertraglich ist dieses "Kaufen-und-gleich-wieder-Verkaufen" ueber denselben
    # Netzanschluss ohnehin nicht sinnvoll -- ein Bezugstarif-Erloes gilt nur
    # fuer TATSAECHLICH bezogene (verbrauchte) Energie, nicht fuer durchgeleitete.
    # Die Batterie kann trotzdem weiterhin frei aus dem Netz laden (siehe
    # link_netz_charge unten, unveraendert) -- nur der direkte Netz->AC-Pfad
    # ist jetzt auf "so viel wie die Last gerade braucht" begrenzt.
    link_netz_ac = solph.components.Converter(
        label="netzbezug_zu_ac",
        inputs={b_netzbezug: solph.Flow()},
        outputs={b_ac: solph.Flow(nominal_capacity=1, maximum=last_profile.values)},
        conversion_factors={b_ac: 1.0},
    )
    # NEU: Netz->Ladebus (Batterie-Ladung AUS DEM NETZ) ist jetzt auf das
    # Minimum aus der allgemeinen Batterieleistung (C21) UND dem neuen,
    # netzseitigen Grenzwert "Maximaler Bezug Batterie" (C37) gekappt, statt
    # nur auf C21 wie bisher. Modellierungs-Annahme (bitte mit Beat pruefen,
    # falls die Absicht eine andere war): C37 begrenzt spezifisch, wie viel
    # Leistung die Batterie ueber ihren Netzanschluss beziehen darf --
    # PV-Ladung (link_pv_charge oben) ist davon NICHT betroffen, da sie einen
    # separaten physischen Pfad nimmt. In Beats aktuellem Beispiel ist C37
    # (500 kW) grosszuegiger als C21 (250 kW), die Kappung ist also aktuell
    # nicht bindend -- wird aber relevant, sobald C37 < C21 gesetzt wird.
    max_bezug_batterie = min(params["leistung"], params["max_bezug_batterie"])
    link_netz_charge = solph.components.Converter(
        label="netzbezug_zu_ladebus",
        inputs={b_netzbezug: solph.Flow()},
        outputs={b_ladebus: solph.Flow(nominal_capacity=max_bezug_batterie)},
        conversion_factors={b_ladebus: 1.0},
    )
    # NEU (Root-Cause-Fix, Beats Rueckmeldung 11.9.2026: "wieso wird die
    # Batterie nicht genutzt um Lastspitzen zu brechen oder den Strombezug
    # von HT zu NT zu verschieben?" -- Ursache war Parameter!C45=0 in seiner
    # Datei "Inputs_BAT_Last_PV_ohneEntladen.xlsx", das bisher NICHT nur die
    # Einspeisung ins Netz, sondern die GESAMTE Entladeleistung auf 0 kappte,
    # also auch Last-Deckung/Peak-Shaving/HT-NT-Verschiebung. Beat hat auf
    # Nachfrage explizit bestaetigt: "Maximale Einspeisung Batterie" (C45)
    # soll NUR die Abgabe ins Netz begrenzen -- die Batterie soll trotzdem
    # frei zur Last hin entladen koennen):
    #
    # Entladebus->AC (Batterie-Entladung Richtung Last) UND Entladebus->
    # Export (eigener Pfad, bepreist mit dem Batterie-Rueckliefertarif,
    # Parameter!C19-C22) sind ZWEI getrennte Abgaenge von b_entladebus. Das
    # Limit "Maximale Einspeisung Batterie" (C45) gilt jetzt NUR NOCH fuer
    # den EXPORT-Abgang (entladebus_zu_export) -- der Last-Abgang
    # (entladebus_zu_ac) ist davon unabhaengig und bleibt nur durch die
    # allgemeine Batterieleistung (C30) gekappt. `discharge_flow` (weiter
    # unten, der EINE Ausgang der Batterie selbst, BEVOR er sich auf die
    # zwei Ziele aufteilt) ist ebenfalls nur noch mit C30 gekappt, NICHT mehr
    # zusaetzlich mit C45 -- C45=0 bedeutet also jetzt "kein Netz-Export ab
    # Batterie", NICHT mehr "Batterie komplett stillgelegt".
    max_einspeisung_batterie = min(params["leistung"], params["max_einspeisung_batterie"])
    link_entlade_ac = solph.components.Converter(
        label="entladebus_zu_ac",
        inputs={b_entladebus: solph.Flow()},
        outputs={b_ac: solph.Flow(nominal_capacity=params["leistung"])},
        conversion_factors={b_ac: 1.0},
    )
    link_batt_export = solph.components.Converter(
        label="entladebus_zu_export",
        inputs={b_entladebus: solph.Flow()},
        outputs={
            b_export: solph.Flow(
                nominal_capacity=max_einspeisung_batterie,
                variable_costs=-zeitreihen["rueckliefertarif_batterie"].values,
            )
        },
        conversion_factors={b_export: 1.0},
    )

    # --- Batterie --------------------------------------------------------------
    # Trafo-Wirkungsgrad (C22) vereinfachend mit Lade-/Entladewirkungsgrad
    # verrechnet (siehe Modulkommentar oben).
    eta_laden = params["lade_wirkungsgrad"] * params["trafo_wirkungsgrad"]
    eta_entladen = params["entlade_wirkungsgrad"] * params["trafo_wirkungsgrad"]

    # custom_properties fuer generic_integral_limit (Vollzyklen-Durchsatz) --
    # Attributname je nach oemof-Version ggf. anzupassen (siehe Hinweis oben).
    # NEU (Beat, 18.9.2026: "wie werde ich das FutureWarning los?"): heisst
    # jetzt `custom_properties` statt `custom_attributes` (der alte Name
    # loeste eine FutureWarning aus -- neuere oemof.solph-Versionen mappen
    # ihn intern ohnehin auf denselben Speicherort, exakt wie bei
    # nominal_capacity/nominal_value unten). generic_integral_limit() weiter
    # unten liest ueber den Keyword-String "vollzyklen_keyword", nicht ueber
    # den Attributnamen selbst -- davon also unberuehrt. Nicht direkt in
    # dieser Sandbox verifizierbar (kein lauffaehiges oemof.solph hier, siehe
    # Modulkopf) -- bitte beim naechsten echten Lauf pruefen, dass die
    # Vollzyklen-Durchsatzbegrenzung weiterhin greift (z.B. am Solver-Log
    # oder an `SOC_kWh` ueber ein extrem hohes C29 hinaus).
    # variable_costs: NEU, der kleine Degradationskosten-Term (siehe
    # DEGRADATIONSKOSTEN_CHF_PRO_KWH oben) -- oemof multipliziert
    # Flow(kW) * variable_costs(CHF/kWh) * Zeitschrittlaenge(h) automatisch
    # in die Zielfunktion, genau wie bei den Bezugs-/Ruecklieferkosten.
    # NEU (Root-Cause-Fix, siehe ausfuehrlicher Kommentar bei
    # link_entlade_ac/link_batt_export oben): nominal_capacity ist jetzt
    # WIEDER nur params["leistung"] (die allgemeine Batterieleistung C30) -- die
    # "Maximale Einspeisung Batterie" (C45) sitzt jetzt AUSSCHLIESSLICH auf
    # dem Export-Abgang (entladebus_zu_export), nicht mehr hier auf dem
    # kombinierten Ausgang. Vorher war C45=0 gleichbedeutend mit "Batterie
    # komplett stillgelegt" (auch fuer Last-Deckung/Peak-Shaving/HT-NT-
    # Verschiebung) -- das entsprach nicht Beats Absicht (siehe Bugfix-Log).
    discharge_flow = solph.Flow(
        nominal_capacity=params["leistung"],
        custom_properties={"vollzyklen_keyword": 1},
        variable_costs=DEGRADATIONSKOSTEN_CHF_PRO_KWH,
    )

    storage = solph.components.GenericStorage(
        label="batterie",
        # NEU: `nominal_capacity` statt `nominal_storage_capacity` -- neuere
        # oemof.solph-Versionen werfen sonst eine FutureWarning ("beide
        # Optionen koennen nicht gleichzeitig gesetzt werden"; die alte
        # Bezeichnung bleibt vorerst nur aus Kompatibilitaetsgruenden
        # erhalten und wird intern ohnehin auf die neue umgemappt).
        nominal_capacity=params["kapazitaet"],
        inputs={b_ladebus: solph.Flow(nominal_capacity=params["leistung"])},
        outputs={b_entladebus: discharge_flow},
        inflow_conversion_factor=eta_laden,
        outflow_conversion_factor=eta_entladen,
        min_storage_level=params["entladegrenze_soc"],
        max_storage_level=params["ladegrenze_soc"],
        initial_storage_level=0.5,
        balanced=True,
        loss_rate=0.0,
    )

    es.add(
        b_pv, b_netzbezug, b_ladebus, b_entladebus, b_ac, b_export,
        src_pv, src_netz, snk_last, snk_export, snk_curtailment,
        link_pv_ac, link_pv_export, link_pv_charge, link_netz_ac, link_netz_charge,
        link_entlade_ac, link_batt_export,
        storage,
    )

    om = solph.Model(es)

    node_refs = {
        "b_pv": b_pv, "b_netzbezug": b_netzbezug, "b_ladebus": b_ladebus,
        "b_entladebus": b_entladebus, "b_ac": b_ac, "b_export": b_export,
        "src_netz": src_netz, "storage": storage, "snk_export": snk_export,
        "snk_curtailment": snk_curtailment,
        "link_pv_charge": link_pv_charge, "link_netz_charge": link_netz_charge,
        "link_entlade_ac": link_entlade_ac, "link_netz_ac": link_netz_ac,
        "link_pv_ac": link_pv_ac, "link_pv_export": link_pv_export,
        "link_batt_export": link_batt_export,
    }
    return es, om, node_refs, dt_hours


# --------------------------------------------------------------------------
# 3. Peak-Shaving (custom Pyomo-Block)
# --------------------------------------------------------------------------

def add_peak_shaving(om, netz_source, netzbezug_bus, leistungspreis: float, timeindex: pd.DatetimeIndex):
    t_month = {t: (timeindex[t].year, timeindex[t].month) for t in om.TIMESTEPS}
    months = sorted(set(t_month.values()))

    block = po.Block()
    om.add_component("PeakShaving", block)
    block.MONATE = po.Set(initialize=months)
    block.peak = po.Var(block.MONATE, within=po.NonNegativeReals)

    def _peak_rule(b, t):
        return b.peak[t_month[t]] >= om.flow[netz_source, netzbezug_bus, t]

    block.peak_constraint = po.Constraint(om.TIMESTEPS, rule=_peak_rule)

    peak_kosten = sum(block.peak[m] * leistungspreis for m in block.MONATE)
    # NEU (Root-Cause-Fix eines Solver-Absturzes, siehe main()): diese
    # Funktion erweitert die Zielfunktion NICHT MEHR selbst -- sie gibt den
    # Kostenterm nur noch zurueck. Grund: siehe ausfuehrlichen Kommentar bei
    # der EINMALIGEN `om.objective.expr = ...`-Zusammenfuehrung in main().
    return block, peak_kosten


# --------------------------------------------------------------------------
# 4. SRL Variante B: Teil der Optimierung (einfache Variante)
# --------------------------------------------------------------------------

def add_srl_optimiert(om, storage, ladebus, entladebus, p_nom: float,
                       srl_pos_price: np.ndarray, srl_neg_price: np.ndarray,
                       dt_hours: float):
    block = po.Block()
    om.add_component("SRLReserve", block)
    block.r_pos = po.Var(om.TIMESTEPS, within=po.NonNegativeReals)
    block.r_neg = po.Var(om.TIMESTEPS, within=po.NonNegativeReals)

    def _p_ist(t):
        return om.flow[storage, entladebus, t] - om.flow[ladebus, storage, t]

    def _pos_rule(b, t):
        return b.r_pos[t] <= p_nom - _p_ist(t)

    def _neg_rule(b, t):
        return b.r_neg[t] <= p_nom + _p_ist(t)

    block.pos_constraint = po.Constraint(om.TIMESTEPS, rule=_pos_rule)
    block.neg_constraint = po.Constraint(om.TIMESTEPS, rule=_neg_rule)

    erloes = sum(
        (block.r_pos[t] * float(srl_pos_price[t]) + block.r_neg[t] * float(srl_neg_price[t])) * dt_hours
        for t in om.TIMESTEPS
    )
    # NEU (Root-Cause-Fix, siehe Kommentar bei add_peak_shaving() und in
    # main()): Zielfunktion nicht mehr direkt hier per `+=` erweitern, nur
    # noch den Kostenterm zurueckgeben (Erloes = negative Kosten).
    return block, -erloes


# --------------------------------------------------------------------------
# 5. Vollzyklen-Durchsatzbegrenzung (nativ ueber oemof.solph.constraints)
# --------------------------------------------------------------------------

def add_vollzyklen_limit(om, storage, entladebus, durchsatz_mwh_jahr: float):
    flows = {(storage, entladebus): storage.outputs[entladebus]}
    solph.constraints.generic_integral_limit(
        om, "vollzyklen_keyword", flows, upper_limit=durchsatz_mwh_jahr * 1000.0
    )


# --------------------------------------------------------------------------
# 5b. Vorab-Pruefung haeufiger Ursachen fuer "termination condition: infeasible"
# --------------------------------------------------------------------------

def check_feasibility_preconditions(
    params: dict, zeitreihen: pd.DataFrame, pv_profile: pd.Series, last_profile: pd.Series
):
    """Grobe Vorab-Pruefung der zwei wahrscheinlichsten Ursachen, wenn der
    Solver mit "termination condition: infeasible" abbricht -- damit ein
    Parameterproblem im EXCEL (statt ein Solver-/Code-Fehler) eine klare,
    verstaendliche Meldung ergibt statt einer kryptischen Pyomo-Meldung ohne
    jeden Kontext. Wird in main() aufgerufen, BEVOR der Solver ueberhaupt
    laeuft.

    Diese Pruefungen sind bewusst KONSERVATIV/UNGEFAEHR (Worst-Case-
    Abschaetzungen, keine exakte LP-Machbarkeitspruefung, insb. Pruefung 2
    ignoriert den aktuellen SOC-Verlauf) -- sie fangen die haeufigsten, leicht
    zu diagnostizierenden Ursachen ab, ersetzen aber nicht den Solver selbst.
    PV-Ueberschuss ueber die Kapazitaet von link_pv_ac/link_pv_export/
    link_pv_charge hinaus wird NUR DANN separat geprueft (Pruefung 3), wenn
    Parameter!C27 ("Abschaltung bei Negativpreisen") = "Nein" ist -- normal-
    erweise (C27="Ja"/leer) hat build_energysystem() dafuer einen
    Abregelungs-Sink (siehe `snk_curtailment`), der diesen Fall strukturell
    abfaengt, aber bei C27="Nein" ist dieser Sink bewusst auf 0 gekappt.

    Wirft RuntimeError mit einer verstaendlichen Sammel-Fehlermeldung, falls
    mind. eine Pruefung anschlaegt.
    """
    fehler = []

    # Immer ausgeben (auch wenn alles i.O. ist) -- damit man auf einen Blick
    # sieht, mit welchen Werten die Pruefung tatsaechlich gerechnet hat. Sehr
    # nuetzlich, um z.B. Einheiten-Missverstaendnisse (kWh/15min vs. kW) oder
    # falsch gelesene Excel-Zellen sofort zu erkennen, statt nur der eigenen
    # Erwartung ("meine Last ist doch max. 200 kW") vertrauen zu muessen.
    print(
        f"  Last-Profil (Zeitreihen!I, umgerechnet in kW): Maximum = "
        f"{last_profile.max():.1f} kW"
    )
    print(
        f"  Kapazitaeten: max. Netzbezug (Parameter!C35) = {params['max_bezug']:.1f} kW, "
        f"Batterie-Leistung (Parameter!C21) = {params['leistung']:.1f} kW"
    )

    # 1. SOC-Grenzen vs. fixer Start-/End-SOC von 50% (initial_storage_level=
    #    0.5, balanced=True in build_energysystem()): liegt Parameter!C30
    #    (Ladegrenze/max. SOC) UNTER 50% oder Parameter!C31 (Entladegrenze/
    #    min. SOC) UEBER 50%, ist bereits der Startzustand bei t=0 ausserhalb
    #    der erlaubten Bandbreite -> GARANTIERT infeasible, unabhaengig von
    #    jeder Zeitreihe.
    if params["ladegrenze_soc"] < 0.5:
        fehler.append(
            f"Parameter!C30 (Ladegrenze/max. SOC) = {params['ladegrenze_soc']:.0%} "
            f"liegt UNTER dem fixen Start-/End-SOC von 50%. Das Modell startet "
            f"(und endet, wegen balanced=True) bei 50% Fuellstand -- das ist "
            f"ausserhalb der erlaubten SOC-Bandbreite und macht das Modell fuer "
            f"JEDEN Zeitschritt unloesbar. Bitte C30 auf mind. 50% setzen."
        )
    if params["entladegrenze_soc"] > 0.5:
        fehler.append(
            f"Parameter!C31 (Entladegrenze/min. SOC) = {params['entladegrenze_soc']:.0%} "
            f"liegt UEBER dem fixen Start-/End-SOC von 50%. Gleiches Problem wie "
            f"bei C30 oben -- bitte C31 auf max. 50% setzen."
        )

    # 2. Grobe Kapazitaets-Abschaetzung: kann die fixe Last (snk_last, siehe
    #    build_energysystem()) ueberhaupt gedeckt werden? Worst-Case pro
    #    Zeitschritt = PV(t) + maximaler Netzbezug (C35) + maximale
    #    Batterie-Entladeleistung (C21) -- optimistisch (ignoriert, ob der SOC
    #    zu diesem Zeitpunkt tatsaechlich genug Energie fuer volle
    #    Entladeleistung haette), um nur die eindeutigen Faelle zu erkennen.
    max_verfuegbar = pv_profile.values + params["max_bezug"] + params["leistung"]
    defizit = last_profile.values - max_verfuegbar
    if defizit.max() > 1e-6:
        i = int(np.argmax(defizit))
        # Roh-Zellwert aus Zeitreihen!I mit ausgeben -- die Umrechnung in kW
        # nimmt an, dass diese Spalte kWh PRO 15-MIN-INTERVALL enthaelt (siehe
        # build_last_profile(): kW = Roh-Wert / DT_HOURS = Roh-Wert * 4).
        # Liegt der Roh-Wert eigentlich schon in kW vor (haeufigster
        # Verwechslungs-Fall), erscheint die Last hier faelschlich 4x zu hoch
        # -- das laesst sich anhand der beiden Zahlen unten direkt pruefen.
        roh_kwh = float(zeitreihen["last_kwh"].values[i]) if "last_kwh" in zeitreihen.columns else None
        roh_hinweis = ""
        if roh_kwh is not None:
            roh_hinweis = (
                f" Roh-Wert in Zeitreihen!I an dieser Stelle: {roh_kwh:.1f} -- "
                f"umgerechnet als kWh/15-Min-Intervall ergibt das {roh_kwh / DT_HOURS:.1f} kW. "
                f"Falls dieser Roh-Wert eigentlich schon direkt in kW gemeint war "
                f"(nicht als Energie pro Intervall), ist DAS sehr wahrscheinlich "
                f"die Ursache -- dann bitte Bescheid geben, dann rechnen wir "
                f"Zeitreihen!I ohne die /DT_HOURS-Umrechnung."
            )
        fehler.append(
            f"Die fixe Last uebersteigt selbst im guenstigsten Fall (PV-Erzeugung "
            f"+ maximaler Netzbezug [Parameter!C35={params['max_bezug']:.0f} kW] + "
            f"maximale Batterie-Entladeleistung [Parameter!C21={params['leistung']:.0f} kW]) "
            f"um bis zu {defizit.max():.1f} kW (z.B. bei Zeitschritt {i}, "
            f"{zeitreihen.index[i]}: Last = {last_profile.values[i]:.1f} kW, "
            f"PV = {pv_profile.values[i]:.1f} kW).{roh_hinweis} Zum Vergleich: "
            f"das Maximum der Last ueber das GANZE Jahr betraegt "
            f"{last_profile.max():.1f} kW (siehe Ausgabe oben)."
        )

    # 3. NEU (Beats Erweiterung 11.9.2026, Parameter!C27 "Abschaltung bei
    #    Negativpreisen"): nur relevant, wenn Abregelung NICHT erlaubt ist --
    #    dann ist `snk_curtailment` in build_energysystem() auf 0 gekappt und
    #    b_pv hat nur noch drei Abgaenge (link_pv_ac + link_pv_export, je
    #    gekappt auf dc_leistung C25, sowie link_pv_charge, gekappt auf
    #    leistung C30). Uebersteigt die PV-Erzeugung in einem Zeitschritt die
    #    Summe dieser drei Kapazitaeten, gibt es KEINEN Abfluss mehr dafuer
    #    -> garantiert infeasible. Bei erlaubter Abregelung (Standardfall)
    #    faengt der Sink das strukturell ab, diese Pruefung greift dann nicht.
    pv_abschaltung_erlaubt = determine_pv_abschaltung_erlaubt(params)
    if not pv_abschaltung_erlaubt:
        max_pv_ohne_abregelung = 2 * params["dc_leistung"] + params["leistung"]
        pv_ueberschuss = pv_profile.values - max_pv_ohne_abregelung
        if pv_ueberschuss.max() > 1e-6:
            i = int(np.argmax(pv_ueberschuss))
            fehler.append(
                f"Parameter!C27 (Abschaltung bei Negativpreisen) = 'Nein' -- die "
                f"PV-Anlage darf nicht abgeregelt werden. Die PV-Erzeugung "
                f"uebersteigt aber an mind. einem Zeitschritt (z.B. Zeitschritt "
                f"{i}, {zeitreihen.index[i]}: PV = {pv_profile.values[i]:.1f} kW) "
                f"die kombinierte Aufnahmekapazitaet von Direktverbrauch + "
                f"PV-Export [je Parameter!C25={params['dc_leistung']:.0f} kW] + "
                f"Batterieladung [Parameter!C30={params['leistung']:.0f} kW] = "
                f"{max_pv_ohne_abregelung:.1f} kW um bis zu "
                f"{pv_ueberschuss.max():.1f} kW. Ohne Abregelungs-Sicherheitsventil "
                f"ist das Modell fuer diesen Zeitschritt unloesbar. Bitte entweder "
                f"C27 auf 'Ja' setzen, oder DC Leistung (C25) / Batterieleistung "
                f"(C30) erhoehen."
            )

    if fehler:
        nummerierte = "\n\n".join(f"  {i + 1}. {f}" for i, f in enumerate(fehler))
        raise RuntimeError(
            "Vorab-Pruefung hat vor dem Solver-Aufruf mind. eine wahrscheinliche "
            "Ursache fuer 'termination condition: infeasible' gefunden:\n\n"
            f"{nummerierte}\n\n"
            "Bitte die genannten Excel-Parameter pruefen/anpassen und erneut "
            "laufen lassen. (Diese Pruefung ist eine grobe Abschaetzung -- "
            "falls sie faelschlich anschlaegt oder eine andere Ursache "
            "vorliegt, bitte mit dieser Fehlermeldung zurueckmelden.)"
        )


# --------------------------------------------------------------------------
# 6. Loesen & Ergebnisse extrahieren
# --------------------------------------------------------------------------

def _ensure_solver_auf_path():
    """NEU (Beat, 18.9.2026: "WARNING: Could not locate the 'cbc' executable"
    trotz `conda install ... coin-or-cbc` -- Verdacht: "ich glaube das ist
    weil VS Code nicht mehr automatisch das Environment oeffnet").

    Root Cause: ein conda-Environment auf Windows legt Solver-Programme wie
    `cbc.exe` NICHT neben `python.exe`, sondern in dessen `Library\\bin`-
    Unterordner ab. Dieser Ordner landet nur dann im PATH der Shell, wenn das
    Environment tatsaechlich AKTIVIERT wurde (z.B. `conda activate ...` oder
    VS Codes automatische Environment-Aktivierung im Terminal). Startet VS
    Code das Skript stattdessen ueber einen direkt ausgewaehlten Python-
    Interpreter OHNE volle Shell-Aktivierung, fehlt `Library\\bin` im PATH --
    Pyomo findet `cbc` dann nicht, OBWOHL es im Environment tatsaechlich
    installiert ist (genau das dortige Environment ist ja an `sys.prefix`
    erkennbar, siehe `site-packages`-Pfad im Traceback).

    Fix: unabhaengig davon, ob die Shell das Environment aktiviert hat, wird
    hier der zum GERADE LAUFENDEN Python-Interpreter (`sys.prefix`) gehoerige
    Solver-Ordner selbst ermittelt und -- falls vorhanden und noch nicht im
    PATH -- vorne eingefuegt, BEVOR der Solver gesucht wird. Betrifft nur den
    Prozess dieses Laufs (kein dauerhafter Eingriff in die Windows-PATH-
    Umgebungsvariable)."""
    kandidaten = [
        os.path.join(sys.prefix, "Library", "bin"),  # conda auf Windows
        os.path.join(sys.prefix, "bin"),              # conda/venv auf Linux/Mac
    ]
    pfad_eintraege = os.environ.get("PATH", "").split(os.pathsep)
    for ordner in kandidaten:
        if os.path.isdir(ordner) and ordner not in pfad_eintraege:
            os.environ["PATH"] = ordner + os.pathsep + os.environ.get("PATH", "")
            print(f"  (Solver-Suche: '{ordner}' zum PATH dieses Laufs hinzugefuegt)")


def solve(om, solver: str = SOLVER, logfile: str | None = None,
          symbolic_solver_labels: bool | None = None):
    """Loest das Modell.

    GEFUNDEN + BEHOBEN (Beat: wiederholter Absturz "termination condition:
    unknown" bei CBC -- Diagnose bisher unmoeglich): `om.solve(solver=solver)`
    ohne `tee`/`logfile` unterdrueckt den KOMPLETTEN Solver-eigenen Output --
    genau die Zeilen, die CBC VOR so einem Fehler ausgibt (Iterationszahl,
    ob CBC selbst "infeasible"/"unbounded" vermutet, Rundungs-/Toleranz-
    Warnungen etc.), waren fuer Beat gar nicht sichtbar, egal wie oft er den
    kompletten Konsolen-Output geschickt haette. Jetzt: `tee=True` gibt den
    vollen CBC-Log live in die Konsole aus, UND (NEU) wird zusaetzlich in
    eine Log-Datei geschrieben (Pfad wird von `main()` gesetzt und vor dem
    Solve-Aufruf ausgegeben) -- damit kann die Datei bei einem erneuten
    Absturz einfach mitgeschickt werden, ohne Copy-Paste-Verluste.

    Robust gegenueber oemof.solph-Versionsunterschieden (siehe die anderen
    FutureWarning-Kommentare in diesem Modul zur API-Instabilitaet): manche
    Versionen erwarten `solve_kwargs={"tee": ..., "logfile": ...}`, andere
    `tee=...`/`logfile=...` direkt als Keyword-Argumente. Erst der modernere
    (Standard-)Weg, bei TypeError automatischer Fallback -- schlaegt auch der
    fehl, wird ganz ohne Logging geloest (mit klarer Konsolen-Warnung), damit
    der Lauf nicht komplett blockiert.

    `symbolic_solver_labels` (NEU, 15.9.2026, Diagnose-Werkzeug fuer den
    "duplicates in objective and matrix"-Absturz, siehe ausfuehrlicher
    Kommentar in main()): Default `None` -- unveraendertes Verhalten (Pyomos
    eigener Default, wir setzen nichts). Falls der quicksum-Rebuild in main()
    den Absturz NICHT behebt, ist dies der naechste Verdaechtige: `True`
    erzwingt volle, lesbare Variablennamen im LP-File statt automatisch
    generierter Kurzformen (oder umgekehrt, `False`) -- ein bekannter
    Fallstrick bei manchen CBC-Versionen mit sehr grossen LP-Dateien
    (Hashing-Kollisionen im LP-Parser bei sehr vielen/sehr langen
    Variablennamen). Zum Testen: `main(..., symbolic_solver_labels=True)`
    bzw. `False` aufrufen und pruefen, ob sich die Fehlermeldung dadurch
    aendert -- OHNE Rueckmeldung aus einem echten Testlauf bleibt unklar,
    welche Richtung (falls ueberhaupt) hilft.
    """
    _ensure_solver_auf_path()
    solve_kwargs = {"tee": True}
    if logfile:
        solve_kwargs["logfile"] = logfile
    if symbolic_solver_labels is not None:
        solve_kwargs["symbolic_solver_labels"] = symbolic_solver_labels
    try:
        om.solve(solver=solver, solve_kwargs=solve_kwargs)
    except TypeError:
        try:
            om.solve(solver=solver, **solve_kwargs)
        except TypeError:
            print(
                "WARNUNG: Konnte tee/logfile nicht an om.solve() uebergeben "
                "(oemof.solph-Versionsunterschied) -- loese ohne Solver-Log. "
                "Bitte diese Warnung mitschicken, falls der Solve danach "
                "fehlschlaegt."
            )
            om.solve(solver=solver)


def remove_custom_blocks(om):
    """Entfernt die custom Pyomo-Bloecke (PeakShaving, SRLReserve) wieder aus
    dem Modell -- MUSS aufgerufen werden, NACHDEM ihre Werte ausgelesen wurden
    (siehe main()), aber BEVOR `solph.processing.results(om)` aufgerufen wird.

    Grund (aus einem echten Testlauf bestaetigt): diese Bloecke enthalten
    Variablen, deren Index NICHT dem ueblichen (Quelle, Ziel, Zeitschritt)-
    Flow-Schema entspricht -- `peak[Monat]` ist nur ueber ein (Jahr, Monat)-
    Tupel indiziert, `r_pos`/`r_neg` nur ueber den Zeitschritt, beide ohne
    Bus-Bezug. `solph.processing.results()` scannt aber pauschal ALLE
    Var-Komponenten des Modells (`model.component_data_objects(Var)`) und
    versucht sie generisch als Flow-Ergebnisse zu interpretieren. Fuer `peak`
    fuehrt das zu einer winzigen Ergebnisgruppe (eine Zeile je Monat statt
    einer je 15-Min-Zeitschritt), was beim Zuweisen des vollen Zeitindex mit
    einem Laengen-Mismatch abstuerzt (`ValueError: Length mismatch ...`,
    gefolgt von `AttributeError: 'int' object has no attribute 'label'` in
    oemof.solph.processing.set_result_index()'s Fehlerbehandlung selbst)."""
    for block_name in ("PeakShaving", "SRLReserve"):
        if hasattr(om, block_name):
            om.del_component(block_name)


def extract_timeseries(results, node_refs, zeitreihen):
    idx = zeitreihen.index
    n = len(idx)

    storage = node_refs["storage"]
    soc = pd.Series(
        results[(storage, None)]["sequences"]["storage_content"].values[:n], index=idx
    )

    # Ladeflüsse: aus den Converter-Ausgangs-Flows in den Ladebus
    ch_pv = pd.Series(
        results[(node_refs["link_pv_charge"], node_refs["b_ladebus"])]["sequences"]["flow"].values[:n],
        index=idx,
    )
    ch_grid = pd.Series(
        results[(node_refs["link_netz_charge"], node_refs["b_ladebus"])]["sequences"]["flow"].values[:n],
        index=idx,
    )
    dis = pd.Series(
        results[(node_refs["storage"], node_refs["b_entladebus"])]["sequences"]["flow"].values[:n],
        index=idx,
    )
    netzbezug_total = pd.Series(
        results[(node_refs["src_netz"], node_refs["b_netzbezug"])]["sequences"]["flow"].values[:n],
        index=idx,
    )
    # NEU: last-spezifischer Netzbezug (nur der Anteil, der ueber
    # link_netz_ac Richtung b_ac/Last fliesst -- per Konstruktion auf
    # last_profile(t) gekappt, siehe build_energysystem(); ist bei Last(t)=0
    # also IMMER exakt 0). Im Unterschied dazu ist netzbezug_total (oben)
    # die AGGREGIERTE Bezugsmenge inkl. Batterie-Netzladung (link_netz_charge),
    # die auch ganz ohne jede Last positiv sein kann (reine Arbitrage-Ladung).
    # Fuer die "haette die Last sonst Netzstrom gekauft"-Bewertung in
    # allocate_modules() (Eigenverbrauch vs. Ruecklieferung) ist NUR dieser
    # last-spezifische Wert korrekt -- siehe Kommentar dort.
    netzbezug_last = pd.Series(
        results[(node_refs["link_netz_ac"], node_refs["b_ac"])]["sequences"]["flow"].values[:n],
        index=idx,
    )
    # NEU: PV-Direktfluss Richtung Hausbus (link_pv_ac) -- noetig in
    # allocate_modules() fuer die Last/Export-Zieldaufteilung der
    # Batterieentladung (siehe dortiger Kommentar zu "Eigenverbrauchsoptimierung"
    # vs. "Einspeiseoptimierung").
    pv_ac = pd.Series(
        results[(node_refs["link_pv_ac"], node_refs["b_ac"])]["sequences"]["flow"].values[:n],
        index=idx,
    )
    # Rueckspeisung (Export ins Netz) -- fuer output_create.py's Monats-Chart
    # (Eigenverbrauch/Rueckspeisung), noetig um "Eigenverbrauch = PV - Export"
    # zu bilden. NEU: liest jetzt von b_export->snk_export (siehe
    # build_energysystem()) statt b_ac->snk_export -- die GESAMTMENGE bleibt
    # dieselbe (Last und Export sind seit dem Bus-Split sauber getrennt),
    # nur der Pfadname hat sich geaendert.
    export = pd.Series(
        results[(node_refs["b_export"], node_refs["snk_export"])]["sequences"]["flow"].values[:n],
        index=idx,
    )
    # NEU (Beats Wunsch: eigener Rueckliefertarif fuer die Batterie): dank des
    # Bus-Splits (siehe build_energysystem()) sind Last- und Export-Anteil der
    # Batterie-Entladung jetzt EXAKTE Optimierungsergebnisse (keine nachtraeglich
    # HERGELEITETE Merit-Order-Aufteilung mehr noetig) -- direkt aus den beiden
    # Abgaengen von b_entladebus gelesen. Ebenso der direkte PV-Export (ohne
    # Umweg ueber die Batterie), fuer die vollstaendige Nachvollziehbarkeit.
    dis_zu_last = pd.Series(
        results[(node_refs["link_entlade_ac"], node_refs["b_ac"])]["sequences"]["flow"].values[:n],
        index=idx,
    )
    dis_zu_export = pd.Series(
        results[(node_refs["link_batt_export"], node_refs["b_export"])]["sequences"]["flow"].values[:n],
        index=idx,
    )
    pv_export_direkt = pd.Series(
        results[(node_refs["link_pv_export"], node_refs["b_export"])]["sequences"]["flow"].values[:n],
        index=idx,
    )
    # PV-Abregelung (Sicherheitsventil, siehe build_energysystem()) -- im
    # Normalfall (Kapazitaeten ausreichend) ueberall 0; nur informativ.
    pv_abregelung = pd.Series(
        results[(node_refs["b_pv"], node_refs["snk_curtailment"])]["sequences"]["flow"].values[:n],
        index=idx,
    )

    out = pd.DataFrame(
        {
            "soc": soc,
            "ch_pv": ch_pv,
            "ch_grid": ch_grid,
            "dis": dis,
            "netzbezug_total": netzbezug_total,
            "netzbezug_last": netzbezug_last,
            "pv_ac": pv_ac,
            "export": export,
            "dis_zu_last": dis_zu_last,
            "dis_zu_export": dis_zu_export,
            "pv_export_direkt": pv_export_direkt,
            "pv_abregelung": pv_abregelung,
        }
    )
    return out


# --------------------------------------------------------------------------
# 7. Modul-Ertragsaufteilung (Tank-Konvention)
# --------------------------------------------------------------------------

def allocate_modules(ts: pd.DataFrame, zeitreihen: pd.DataFrame, params: dict,
                      dt_hours: float, srl_variante: str,
                      soc_min_srl: float = SOC_MIN_SRL, soc_max_srl: float = SOC_MAX_SRL,
                      last_profile: pd.Series | None = None):
    """Modul-Ertragsaufteilung (Tank-Konvention).

    NEU (Beats Wunsch, siehe Chat): "Eigenverbrauchsoptimierung",
    "Einspeiseoptimierung" und "Arbitrage der Batterie" werden jetzt konsequent
    nach dem ZIEL der Batterieentladung UND ihrer HERKUNFT (PV- oder
    netzgeladen) unterschieden -- nicht mehr nur nach der Herkunft allein wie
    zuvor. Beats eigene Definition (vier Konfigurationen):
      - Batterie+PV+Last: EV & Einspeiseoptimierung zusammen, ein wenig
        Arbitrage.
      - Batterie+PV (keine Last): keine EV moeglich (nichts zu vermeiden) --
        heisst konsequent Einspeiseoptimierung, kann mit Arbitrage zusammen
        auftreten.
      - Nur Batterie (kein PV, keine Last): weder EV noch Einspeiseoptimierung
        -> 0 (nur ggf. Arbitrage/SRL).
      - Batterie+Last (kein PV): REINE Eigenverbrauchsoptimierung, keine
        Einspeiseoptimierung.
    Das ergibt folgende Zuordnung (auf Beats Wunsch "kombiniert", d.h. beide
    Dimensionen -- Herkunft UND Ziel -- werden weiterhin einzeln mitgefuehrt,
    siehe `ertrag_eigenverbrauch_pv`/`ertrag_eigenverbrauch_grid` unten -- nur
    die drei Modul-SUMMEN fassen sie wie folgt zusammen):
      - Eigenverbrauchsoptimierung = JEDE Entladung (PV- oder netzgeladen),
        die in diesem Zeitschritt tatsaechlich die Last deckt (zum
        Bezugstarif bewertet, da ein Netzkauf vermieden wird).
      - Einspeiseoptimierung = PV-geladene Entladung, die exportiert wird
        (zum Rueckliefertarif bewertet).
      - Arbitrage der Batterie = NUR netzgeladene Entladung, die exportiert
        wird (reines Kaufen-billig/Verkaufen-teuer ueber den Netzanschluss,
        zum Rueckliefertarif bewertet, Kostenbasis Bezugstarif).
    Das ist die einzige Zuordnung, die zu ALLEN vier Beispielen passt (v.a.
    "Batterie+Last: reine EV" -- ohne PV muss die netzgeladene, lastdeckende
    Entladung ebenfalls als Eigenverbrauchsoptimierung zaehlen, sonst waere
    das Modul dort faelschlich 0).

    `last_profile` (optional, Last in kW) wird fuer die ZIEL-Aufteilung
    (deckt Last vs. wird exportiert) gebraucht -- siehe `dis_to_last`/
    `dis_to_export` unten. Ohne `last_profile` (Rueckwaertskompatibilitaet,
    z.B. alte Aufrufer/Tests) wird `last_profile=0` angenommen, d.h. JEDE
    Entladung gilt als Export (konservativ: keine EV-Zuordnung ohne
    tatsaechliche Last-Information).
    """
    kapazitaet = params["kapazitaet"]
    eta_laden = params["lade_wirkungsgrad"] * params["trafo_wirkungsgrad"]
    eta_entladen = params["entlade_wirkungsgrad"] * params["trafo_wirkungsgrad"]

    n = len(ts)
    pv_anteil = np.zeros(n + 1)
    grid_anteil = np.zeros(n + 1)
    anfang_anteil = np.zeros(n + 1)
    anfang_anteil[0] = 1.0  # Start-SOC (50%) ist zu 100% "Anfangsbestand"
    # NEU: laufende, mengengewichtete DURCHSCHNITTSKOSTEN (CHF/kWh) des
    # PV- bzw. netzgeladenen Tank-Anteils (analog zur "Well-mixed-tank"-
    # Konvention oben, aber fuer die KOSTENBASIS statt die Herkunfts-Anteile).
    # Damit wird die Ladekosten-Verrechnung (die beim Laden anfaellt) korrekt
    # der SPAETEREN Entladung zugeordnet (Ziel-Aufteilung ist ja erst beim
    # Entladen bekannt) -- analog zur "gewichteten Durchschnittskosten"-
    # Lagerbuchhaltung.
    kosten_pv = np.zeros(n + 1)
    kosten_grid = np.zeros(n + 1)

    soc = ts["soc"].values
    soc0 = kapazitaet * 0.5
    soc_full = np.concatenate([[soc0], soc])

    dis_pv = np.zeros(n)
    dis_grid = np.zeros(n)
    dis_anfang = np.zeros(n)

    bezugstarif = zeitreihen["bezugstarif"].values
    rueckliefertarif = zeitreihen["rueckliefertarif"].values
    # NEU (Beats Wunsch: eigener Rueckliefertarif fuer die Batterie): JEDE
    # exportierte Batterie-Entladung (egal ob PV- oder netzgeladen im Tank)
    # verlaesst den Standort physisch ueber den Batterie-Netzanschluss (siehe
    # build_energysystem()::link_batt_export) -- muss also mit dem
    # BATTERIE-Rueckliefertarif bewertet werden, nicht mit dem PV-Tarif (die
    # beiden koennen jetzt voneinander abweichen). `rueckliefertarif` (PV)
    # bleibt weiterhin die richtige Bewertung fuer die PV-LADEKOSTEN-Basis
    # (kosten_pv unten -- das ist der Preis, den PV-Strom beim Laden
    # "geopfert" haette, waere er stattdessen direkt exportiert worden).
    rueckliefertarif_batterie = (
        zeitreihen["rueckliefertarif_batterie"].values
        if "rueckliefertarif_batterie" in zeitreihen.columns
        else rueckliefertarif
    )
    ch_pv_arr = ts["ch_pv"].values
    ch_grid_arr = ts["ch_grid"].values
    dis_arr = ts["dis"].values
    pv_ac_arr = ts["pv_ac"].values if "pv_ac" in ts.columns else np.zeros(n)
    netzbezug_last_arr = ts["netzbezug_last"].values if "netzbezug_last" in ts.columns else np.zeros(n)
    last_arr = last_profile.values if last_profile is not None else np.zeros(n)
    # NEU: dank des Bus-Splits in build_energysystem() (b_ac nur noch Last,
    # b_export nur noch Export) ist die Last-/Export-Zieldaufteilung der
    # Batterieentladung jetzt ein EXAKTES Optimierungsergebnis -- kein
    # nachtraeglich HERGELEITETER Merit-Order-Schaetzwert mehr noetig (siehe
    # extract_timeseries()). Faellt auf die alte Merit-Order-Herleitung
    # zurueck, falls `ts` (z.B. in aelteren Tests) diese Spalten noch nicht
    # enthaelt.
    dis_zu_last_arr = ts["dis_zu_last"].values if "dis_zu_last" in ts.columns else None
    dis_zu_export_arr = ts["dis_zu_export"].values if "dis_zu_export" in ts.columns else None

    ertrag_eigenverbrauch_pv = np.zeros(n)
    ertrag_eigenverbrauch_grid = np.zeros(n)
    ertrag_einspeiseoptimierung = np.zeros(n)
    ertrag_arbitrage = np.zeros(n)

    for i in range(n):
        soc_before = soc_full[i]
        soc_after = soc_full[i + 1]

        # GEFUNDEN + BEHOBEN (Beat: "wie kommt Arbitrage zustande, das passt
        # nicht"): ch_pv_arr/ch_grid_arr/dis_arr sind LEISTUNGEN (kW), die
        # SOC-Pools (pv_anteil*soc etc.) sind aber ENERGIEN (kWh) -- hier
        # fehlte bisher komplett der Faktor dt_hours (0.25 bei 15-Min-
        # Schritten), d.h. jede Viertelstunde wurde ein volles kWh statt nur
        # 0.25 kWh in die Tank-Pools gebucht -- eine systematische 4x-
        # Aufblaehung der pv_anteil/grid_anteil-Pools gegenueber dem echten
        # SOC (verifiziert: pv_anteil+grid_anteil+anfang_anteil driftete bis
        # auf ~3.7 statt konstant 1.0). Das hat NUR "Eigenverbrauchs-
        # optimierung"/"Einspeiseoptimierung"/"Arbitrage der Batterie"
        # verfaelscht (SRL und alle uebrigen Zeitreihen sind unabhaengig
        # davon und korrekt) -- bei Beats Q2-Batterie-Szenario allein hat das
        # die ausgewiesene Arbitrage von ca. 13'600 CHF auf 51'387 CHF
        # (rund 3.8x) hochgetrieben.
        ch_pv_i = ch_pv_arr[i] * eta_laden * dt_hours          # kWh, Tank-Zugang aus PV
        ch_grid_i = ch_grid_arr[i] * eta_laden * dt_hours      # kWh, Tank-Zugang aus Netz
        dis_i = dis_arr[i] * dt_hours  # kWh, AC-seitig (bereits nach Entladewirkungsgrad, oemof-Output)

        # GEFUNDEN + BEHOBEN (2. Teil desselben Bugs): die tatsaechliche
        # SOC-ABNAHME ist wegen der Entladeverluste GROESSER als die AC-
        # seitig abgegebene/verkaufte Energie dis_i (siehe
        # outflow_conversion_factor in build_energysystem()) -- fuer die
        # Pool-Buchhaltung (Tank-Erhaltungssatz: pv_anteil+grid_anteil+
        # anfang_anteil muss IMMER exakt 1 ergeben) muss dis_i wieder auf die
        # tatsaechliche SOC-Abnahme hochgerechnet werden. Fuer die ZIEL-
        # Aufteilung (deckt Last/Export) und den ausgewiesenen Ertrag bleibt
        # dagegen dis_i selbst (die tatsaechlich gemessene/verkaufte Menge)
        # massgeblich -- nur die interne Pool-Abbuchung braucht die groessere
        # Zahl.
        dis_depletion_i = dis_i / eta_entladen

        pv_energy_before = pv_anteil[i] * soc_before
        grid_energy_before = grid_anteil[i] * soc_before
        anfang_energy_before = anfang_anteil[i] * soc_before

        dis_pv[i] = dis_depletion_i * pv_anteil[i]
        dis_grid[i] = dis_depletion_i * grid_anteil[i]
        dis_anfang[i] = dis_depletion_i * anfang_anteil[i]

        # ZIEL-Aufteilung der GESAMTEN Entladung dis_i in "deckt Last" vs.
        # "wird exportiert": NEU bevorzugt aus den exakten Optimierungs-
        # Flows (dis_zu_last_arr/dis_zu_export_arr, siehe oben) gelesen --
        # diese kommen direkt aus den zwei getrennten, unterschiedlich
        # bepreisten Abgaengen von b_entladebus (siehe build_energysystem())
        # und sind deshalb exakt, keine Schaetzung mehr. Fallback (nur falls
        # `ts` diese Spalten nicht enthaelt, z.B. aeltere Tests): alte
        # Merit-Order-Naeherung -- lokale Erzeugung/Entladung deckt zuerst
        # die Last (nach Abzug von PV-Direktverbrauch und last-spezifischem
        # Netzbezug), der Rest gilt als Export.
        if dis_zu_last_arr is not None:
            dis_to_last_i = dis_zu_last_arr[i] * dt_hours
            dis_to_export_i = dis_i - dis_to_last_i
        else:
            residual_last_i = max(0.0, (last_arr[i] - pv_ac_arr[i] - netzbezug_last_arr[i]) * dt_hours)
            dis_to_last_i = min(dis_i, residual_last_i)
            dis_to_export_i = dis_i - dis_to_last_i

        dis_pv_to_last_i = dis_to_last_i * pv_anteil[i]
        dis_pv_to_export_i = dis_to_export_i * pv_anteil[i]
        dis_grid_to_last_i = dis_to_last_i * grid_anteil[i]
        dis_grid_to_export_i = dis_to_export_i * grid_anteil[i]

        # Ertrag bei Entladung = (erzielter Preis - Kostenbasis der
        # entladenen Tank-Energie). Kostenbasis = laufender Durchschnitt der
        # Ladekosten (kosten_pv[i]/kosten_grid[i], CHF/kWh), NICHT die
        # Ladekosten DIESES Zeitschritts -- so wird die Ladekosten-Buchung
        # (beim Laden) korrekt mit der spaeteren Entladung verrechnet
        # (zeitlich "gematcht"), unabhaengig davon, wann genau geladen wurde.
        # NEU: dis_pv_to_export_i/dis_grid_to_export_i verlassen den Standort
        # ueber den BATTERIE-Netzanschluss (link_batt_export) -- deshalb hier
        # rueckliefertarif_batterie statt rueckliefertarif (PV). kosten_pv/
        # kosten_grid (die Ladekosten-Basis) bleiben unveraendert ueber
        # rueckliefertarif (PV) hergeleitet (siehe Kommentar bei
        # pv_kosten_chf_after unten -- das ist der Opportunitaetspreis beim
        # LADEN, nicht der Verkaufspreis beim spaeteren Entladen).
        ertrag_eigenverbrauch_pv[i] = dis_pv_to_last_i * (bezugstarif[i] - kosten_pv[i])
        ertrag_einspeiseoptimierung[i] = dis_pv_to_export_i * (rueckliefertarif_batterie[i] - kosten_pv[i])
        ertrag_eigenverbrauch_grid[i] = dis_grid_to_last_i * (bezugstarif[i] - kosten_grid[i])
        ertrag_arbitrage[i] = dis_grid_to_export_i * (rueckliefertarif_batterie[i] - kosten_grid[i])

        pv_energy_after = pv_energy_before + ch_pv_i - dis_pv[i]
        grid_energy_after = grid_energy_before + ch_grid_i - dis_grid[i]
        anfang_energy_after = anfang_energy_before - dis_anfang[i]

        # Kostenbasis fortschreiben: neue Ladekosten (CHF) kommen dazu, die
        # entladene Menge wird zur AKTUELLEN Durchschnittskosten abgebucht.
        # GEFUNDEN + BEHOBEN (3. Teil desselben Bugs): die tatsaechlich
        # BEZAHLTEN CHF beziehen sich auf die AM ZAEHLER/NETZANSCHLUSS
        # bezogene Energie (ch_pv_arr[i]/ch_grid_arr[i], OHNE eta_laden),
        # nicht auf die (kleinere) im Tank ANGEKOMMENE Energie ch_pv_i/
        # ch_grid_i -- vorher wurde die Kostenbasis dadurch systematisch um
        # den Ladewirkungsgrad zu GUENSTIG ausgewiesen (impliziter "Rabatt"
        # auf die Ladeverluste), was den Arbitrage-Ertrag zusaetzlich
        # ueberschaetzt hat.
        pv_kosten_chf_before = kosten_pv[i] * pv_energy_before
        pv_kosten_chf_after = (
            pv_kosten_chf_before + ch_pv_arr[i] * dt_hours * rueckliefertarif[i] - dis_pv[i] * kosten_pv[i]
        )
        grid_kosten_chf_before = kosten_grid[i] * grid_energy_before
        grid_kosten_chf_after = (
            grid_kosten_chf_before + ch_grid_arr[i] * dt_hours * bezugstarif[i] - dis_grid[i] * kosten_grid[i]
        )

        if soc_after > 1e-9:
            pv_anteil[i + 1] = max(0.0, pv_energy_after / soc_after)
            grid_anteil[i + 1] = max(0.0, grid_energy_after / soc_after)
            anfang_anteil[i + 1] = max(0.0, anfang_energy_after / soc_after)
        else:
            pv_anteil[i + 1] = grid_anteil[i + 1] = anfang_anteil[i + 1] = 0.0

        kosten_pv[i + 1] = pv_kosten_chf_after / pv_energy_after if pv_energy_after > 1e-9 else 0.0
        kosten_grid[i + 1] = grid_kosten_chf_after / grid_energy_after if grid_energy_after > 1e-9 else 0.0

    # dis_anfang/anfang_anteil werden weiterhin oben mitgefuehrt (noetig, damit
    # die anfaengliche 50%-SOC-Fuellung nicht faelschlich als Eigenverbrauch-
    # oder Arbitrage-Ertrag erscheint) -- der daraus resultierende Ertrag
    # ("Anfangsbestand-Verwertung") wird aber auf Beats Wunsch NICHT mehr
    # ausgewiesen (weder als Modul-Zeile noch in irgendeiner Summe).

    ertrag_eigenverbrauch = ertrag_eigenverbrauch_pv + ertrag_eigenverbrauch_grid

    module = {
        "Eigenverbrauchsoptimierung": float(np.sum(ertrag_eigenverbrauch)),
        "Einspeiseoptimierung": float(np.sum(ertrag_einspeiseoptimierung)),
        "Arbitrage der Batterie": float(np.sum(ertrag_arbitrage)),
    }

    # SRL-Leistung/-Ertrag pro Zeitschritt -- Default 0 (Variante "keine" oder
    # "optimiert", bei der die tatsaechlichen r_pos/r_neg-Werte erst nach dem
    # Solve aus dem Pyomo-Block ausgelesen werden koennen, siehe main()).
    srl_pos_kw = np.zeros(n)
    srl_neg_kw = np.zeros(n)
    ertrag_srl = np.zeros(n)

    # --- SRL post-hoc ------------------------------------------------------
    if srl_variante == "post_hoc":
        p_nom = params["leistung"]
        p_ist = ts["dis"].values - (ts["ch_pv"].values + ts["ch_grid"].values)
        soc_frac = soc / kapazitaet
        gate = (soc_frac > soc_min_srl) & (soc_frac < soc_max_srl)

        residual_pos = np.where(gate, p_nom - p_ist, 0.0)
        residual_neg = np.where(gate, p_nom + p_ist, 0.0)
        residual_pos = np.clip(residual_pos, 0, None)
        residual_neg = np.clip(residual_neg, 0, None)

        srl_pos_preis = zeitreihen["srl_pos"].values
        srl_neg_preis = zeitreihen["srl_neg"].values

        srl_pos_kw = residual_pos
        srl_neg_kw = residual_neg
        ertrag_srl = residual_pos * srl_pos_preis * dt_hours + residual_neg * srl_neg_preis * dt_hours

        module["SRL (Post-hoc-Mehrertrag)"] = float(np.sum(ertrag_srl))

    return module, pd.DataFrame(
        {
            "pv_anteil": pv_anteil[1:],
            "grid_anteil": grid_anteil[1:],
            "anfang_anteil": anfang_anteil[1:],
            "ertrag_eigenverbrauch": ertrag_eigenverbrauch,
            # NEU: Herkunfts-Detail hinter "ertrag_eigenverbrauch" (Beats
            # Wunsch, beide Dimensionen -- Herkunft UND Ziel -- verfuegbar zu
            # halten): ertrag_eigenverbrauch_pv + ertrag_eigenverbrauch_grid
            # = ertrag_eigenverbrauch.
            "ertrag_eigenverbrauch_pv": ertrag_eigenverbrauch_pv,
            "ertrag_eigenverbrauch_grid": ertrag_eigenverbrauch_grid,
            "ertrag_einspeiseoptimierung": ertrag_einspeiseoptimierung,
            "ertrag_arbitrage": ertrag_arbitrage,
            "srl_pos_kw": srl_pos_kw,
            "srl_neg_kw": srl_neg_kw,
            "ertrag_srl": ertrag_srl,
            # Peak-Shaving ist inhaerent ein Monatswert (Grundgebuehr auf die
            # Monatsspitze), kein echter Pro-Intervall-Fluss -- wird in main()
            # nachtraeglich befuellt (gleicher Monatswert fuer alle Zeilen
            # dieses Monats), Default 0 falls Peak-Shaving inaktiv ist.
            "peak_shaving_kosten_monat": np.zeros(n),
        },
        index=ts.index,
    )


# --------------------------------------------------------------------------
# 7b. Kapitalkosten (informativ, nicht Teil der LP-Zielfunktion)
# --------------------------------------------------------------------------

def capital_costs(params: dict, wacc: float = 0.03, entladeenergie_kwh: float | None = None) -> dict:
    """Kapitalkosten/Unterhalt/Amortisation -- reine Nachbetrachtung fuer
    Wirtschaftlichkeits-Reports (Excel-Output, PDF-Report), NICHT Teil der
    LP-Zielfunktion. Eigene Funktion (statt Inline-Code in main()), damit
    output_create.py dieselbe Berechnung fuer den PDF-Report wiederverwenden
    kann, ohne main() ein weiteres Mal umzubauen.

    Liefert ein Dict mit:
      capex: Investitionssumme (CHF)
      lebenszeit: Jahre (Parameter!C27)
      wacc: verwendeter Kapitalkostensatz (Default-Annahme, siehe unten)
      annuitaet: jaehrliche Investitions-Annuitaet (Kapitalwiedergewinnung,
          CHF/Jahr, klassische Annuitaetenformel)
      unterhalt: jaehrliche Unterhaltskosten (CHF/Jahr, direkt aus
          Parameter!C26 -- C26 ist bereits CHF/Jahr, KEINE Multiplikation
          mit der Kapazitaet mehr)
      zins_jahr1: kalkulatorischer Zins im ERSTEN Jahr des Annuitaetendarlehens
          (= capex * wacc)
      amortisation_jahr1: Tilgungsanteil im ersten Jahr (= annuitaet -
          zins_jahr1)
      degradation: NEU -- Durchsatz-/Degradationskosten (CHF/Jahr), die schon
          IN der LP-Zielfunktion beruecksichtigt sind (siehe
          DEGRADATIONSKOSTEN_CHF_PRO_KWH/discharge_flow in
          build_energysystem()). Wird hier NUR informativ nachgerechnet
          (entladeenergie_kwh * DEGRADATIONSKOSTEN_CHF_PRO_KWH), damit sie im
          Wirtschaftlichkeits-Report als eigene Ausgaben-Zeile auftaucht --
          sonst würden "Total Einnahmen - Total Ausgaben" im PDF NICHT mit
          dem tatsaechlichen LP-Zielfunktionswert uebereinstimmen (der
          Degradationskosten-Term schmaelert den echten Gewinn bereits,
          taucht aber sonst in keiner der post-hoc Modul-Ertraege auf). Wenn
          `entladeenergie_kwh` nicht uebergeben wird (z.B. alter Aufrufer),
          ist der Wert 0.0 -- rueckwaertskompatibel.

    ANNAHME (bitte mit Beat pruefen/anpassen, falls gewuenscht): WACC=3% als
    Default. Die Aufteilung Zins/Tilgung ist die des ERSTEN Jahres eines
    klassischen Annuitaetendarlehens (Zins auf die volle Restschuld, die zu
    Beginn noch = capex ist). In spaeteren Jahren sinkt der Zinsanteil und der
    Tilgungsanteil steigt -- das wird hier NICHT abgebildet (Vereinfachung,
    konsistent mit der Ein-Jahres-Betrachtung des gesamten Modells: wir
    nehmen an, dass sich das simulierte Jahr ueber die ganze Lebenszeit
    identisch wiederholt).
    """
    capex = params["kapazitaet"] * params["invest_kosten_kwh"]
    lebenszeit = params["lebenszeit"]
    annuitaet = capex * (wacc * (1 + wacc) ** lebenszeit) / ((1 + wacc) ** lebenszeit - 1)
    unterhalt = params["unterhalt_kosten"]  # Parameter!C26 ist bereits CHF/Jahr (nicht mehr CHF/kWh)
    zins_jahr1 = capex * wacc
    amortisation_jahr1 = annuitaet - zins_jahr1
    degradation = (
        entladeenergie_kwh * DEGRADATIONSKOSTEN_CHF_PRO_KWH if entladeenergie_kwh is not None else 0.0
    )
    return {
        "capex": capex,
        "lebenszeit": lebenszeit,
        "wacc": wacc,
        "annuitaet": annuitaet,
        "unterhalt": unterhalt,
        "zins_jahr1": zins_jahr1,
        "amortisation_jahr1": amortisation_jahr1,
        "degradation": degradation,
    }


# --------------------------------------------------------------------------
# 8. Hauptablauf
# --------------------------------------------------------------------------

def main(
    input_path: str = INPUT_PATH,
    peakshaving_aktiv: bool | None = None,
    srl_variante: str | None = None,
    soc_min_srl: float = SOC_MIN_SRL,
    soc_max_srl: float = SOC_MAX_SRL,
    last_series: pd.Series | None = None,
    solver: str = SOLVER,
    symbolic_solver_labels: bool | None = None,
    params_override: dict | None = None,
):
    """
    peakshaving_aktiv=None (Default): automatisch aus Parameter!C9 abgeleitet
        (aktiv, falls C9 != 0/NaN). Explizit True/False uebersteuert das Excel.
    srl_variante=None (Default): automatisch aus Parameter!C40 abgeleitet
        ('Ja_Residual' -> 'post_hoc', 'Ja_Optimiert' -> 'optimiert',
        'Nein' -> 'keine'). Explizite Angabe uebersteuert das Excel.
    symbolic_solver_labels=None (Default): siehe solve()-Docstring --
        Diagnose-Werkzeug fuer den "duplicates in objective and matrix"-
        Absturz, nur bei Bedarf explizit auf True/False setzen.
    params_override=None (Default): NEU (15.9.2026, fuer
        optimale_batteriegroesse.py) -- optionales Dict, dessen Eintraege
        NACH dem Excel-Einlesen in `params` ueberschrieben werden (z.B.
        {"kapazitaet": 400.0, "leistung": 100.0}). Damit kann ein Sweep-Skript
        denselben main()-Ablauf fuer viele Batteriegroessen wiederverwenden,
        ohne fuer jede Kombination eine eigene Excel-Datei anzulegen. Das
        Excel bleibt die Basis/Quelle der Wahrheit fuer alle NICHT
        ueberschriebenen Parameter.
    """
    print(f"Lese {input_path} ...")
    params, zeitreihen = read_inputs(input_path)

    if params_override:
        print(f"Ueberschreibe Parameter (params_override): {params_override}")
        params = dict(params)  # Kopie, damit der Aufrufer sein Original-Dict wiederverwenden kann
        params.update(params_override)

    if peakshaving_aktiv is None:
        peakshaving_aktiv = determine_peakshaving_aktiv(params)
        print(f"Peak-Shaving automatisch aus Parameter!C9 abgeleitet: {peakshaving_aktiv}")

    if srl_variante is None:
        srl_variante = determine_srl_variante(params)
        print(
            f"SRL-Variante automatisch aus Parameter!C40 abgeleitet: "
            f"'{params['srl_teilnahme']}'"
        )

    pv_profile = build_pv_profile(params, zeitreihen)
    last_profile = build_last_profile(params, zeitreihen, last_series)

    print("Pruefe Machbarkeits-Vorbedingungen ...")
    check_feasibility_preconditions(params, zeitreihen, pv_profile, last_profile)

    print("Baue Energiesystem ...")
    es, om, node_refs, dt_hours = build_energysystem(params, zeitreihen, pv_profile, last_profile)

    add_vollzyklen_limit(om, node_refs["storage"], node_refs["b_entladebus"], params["durchsatz_mwh_jahr"])

    # NEU (Root-Cause-Fix eines konkreten Solver-Absturzes, gefunden ueber
    # Beats Solver-Log vom 10.9.2026, Datei "Inputs_BAT_Last_PV_ohneLaden",
    # Peak-Shaving aktiv (Parameter!C8=12), SRL='Ja_Residual' -- also NUR
    # Peak-Shaving haengt sich in die Zielfunktion ein, SRL "optimiert" war in
    # diesem Fall NICHT aktiv): CBC brach beim LP-File-Einlesen mit "### ERROR:
    # 105119 duplicates in objective and matrix" / "Current model not valid"
    # ab, was oemof.solph anschliessend als "termination condition: unknown"
    # weiterreicht (OHNE dass ueberhaupt eine Optimierung stattfand). 105119
    # ist fast exakt 3 * 35040 (die Anzahl Zeitschritte) -- alle DREI
    # zeitreihenbasierten Flow-Kosten (Bezugstarif, PV- und Batterie-
    # Rueckliefertarif) tauchen verdoppelt im LP-File auf.
    #
    # KORREKTUR/NEUER BEFUND (15.9.2026, Beats Solver-Log fuer
    # "Inputs_BAT_standalone.xlsx"): urspruenglich vermutet, die Ursache sei
    # die (damalige) `om.objective.expr += ...`-Erweiterung durch Peak-Shaving/
    # SRL "optimiert" (LinearExpression-Listen werden von Pyomo laut eigener
    # Doku NICHT kopiert -- "the lists that are passed to LinearExpression are
    # not copied", ein `+=` darauf kann daher bestehende Kostenterme
    # verdoppeln statt nur anzuhaengen). DIESE Datei hat aber Peak-Shaving
    # INAKTIV (Parameter!C8=0) UND SRL='Ja_Residual' (post-hoc, NICHT Teil der
    # Optimierung) -- der Code-Pfad unten, der die Zielfunktion ueberhaupt
    # anfasst (`if zusatzkosten_terme: ...`), wurde in diesem Lauf also gar
    # NICHT ausgefuehrt (`zusatzkosten_terme` blieb leer). Trotzdem exakt
    # derselbe Fehler, exakt dieselbe Zahl 105119 (= dieselben drei
    # zeitreihenbasierten Flow-Kosten, jetzt aus oemofs EIGENER, unveraenderter
    # Standard-Zielfunktion). Das falsifiziert die urspruengliche Theorie
    # zumindest als VOLLSTAENDIGE Erklaerung: die Verdopplung entsteht
    # offenbar schon beim Aufbau von oemofs eigener Zielfunktion (viele kleine
    # Kostenterme, eine je Zeitschritt/Flow, zu einer LinearExpression
    # zusammengefuehrt) -- vermutlich derselbe Pyomo-Versions-Fallstrick, nur
    # bereits INNERHALB von oemof.solph selbst statt erst durch unseren
    # Zusatzcode.
    #
    # Angepasster Fix: der `quicksum(..., linear=False)`-"Rebuild"-Schritt
    # laeuft jetzt IMMER (nicht mehr nur, wenn Peak-Shaving/SRL "optimiert"
    # tatsaechlich einen Zusatzterm liefern) -- er zwingt Pyomo, die komplette
    # Zielfunktion (inkl. oemofs eigenem Teil) als generischen, nicht auf
    # LinearExpression optimierten Ausdruck neu aufzubauen, was den vermuteten
    # Listen-Aliasing-Fallstrick auch fuer den reinen oemof-Anteil umgehen
    # sollte. Wirtschaftlich/inhaltlich AEQUIVALENT zum bisherigen Verhalten
    # (dieselben Kostenterme, nur einmal statt potenziell doppelt gezaehlt).
    #
    # BITTE trotzdem: (1) diese Datei erneut mit genau demselben
    # "Inputs_BAT_standalone.xlsx" laufen lassen und den neuen Solver-Log
    # zurueckmelden -- falls der Fehler UNVERAENDERT bleibt (exakt dieselbe
    # Zahl 105119), liegt die Ursache hoechstwahrscheinlich ausserhalb dieses
    # Skripts (Pyomo-/oemof.solph-/CBC-Versionskombination) und wir brauchen
    # `pip show pyomo oemof.solph` sowie die genaue CBC-Version (steht schon
    # im Log: "Version: 2.10.13"), um nach einem bekannten Bug in genau dieser
    # Kombination zu suchen bzw. eine Versions-Aenderung zu empfehlen. (2) Da
    # diese Pyomo-Fallstrick-Klasse in dieser Sandbox mangels Pyomo-
    # Installation nicht nachstellbar ist, bleibt dieser Fix bis zur
    # Rueckmeldung experimentell.
    zusatzkosten_terme = []

    if peakshaving_aktiv:
        print("Aktiviere Peak-Shaving ...")
        _, peak_kosten = add_peak_shaving(
            om, node_refs["src_netz"], node_refs["b_netzbezug"],
            params["netznutzung_leistung"], zeitreihen.index,
        )
        zusatzkosten_terme.append(peak_kosten)

    if srl_variante == "optimiert":
        print("Aktiviere SRL als Teil der Optimierung (einfache Variante) ...")
        _, srl_kosten = add_srl_optimiert(
            om, node_refs["storage"], node_refs["b_ladebus"], node_refs["b_entladebus"],
            params["leistung"], zeitreihen["srl_pos"].values, zeitreihen["srl_neg"].values,
            dt_hours,
        )
        zusatzkosten_terme.append(srl_kosten)

    # NEU: laeuft jetzt IMMER, auch mit leerer `zusatzkosten_terme`-Liste
    # (vorher: nur `if zusatzkosten_terme:`) -- siehe Kommentar oben.
    om.objective.expr = po.quicksum([om.objective.expr] + zusatzkosten_terme, linear=False)

    # NEU (Beat: wiederholter, bisher nicht diagnostizierbarer "termination
    # condition: unknown"-Absturz): Solver-Log-Datei neben der Eingabedatei
    # ablegen, damit der volle CBC-Output bei einem erneuten Absturz einfach
    # mitgeschickt werden kann, siehe solve()-Docstring.
    logfile = os.path.join(
        os.path.dirname(os.path.abspath(input_path)),
        f"solver_log_{os.path.splitext(os.path.basename(input_path))[0]}_"
        f"{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.log",
    )
    print(f"Loese mit Solver '{solver}' ... (Solver-Log: {logfile})")
    solve(om, solver, logfile=logfile, symbolic_solver_labels=symbolic_solver_labels)

    # WICHTIG: Custom-Block-Werte (Peak-Shaving, SRL "optimiert") MUESSEN vor
    # solph.processing.results() ausgelesen werden, siehe Kommentar in
    # remove_custom_blocks() -- sonst stuerzt processing.results() an ihnen ab.
    peak_kosten_je_monat = None
    if peakshaving_aktiv:
        peak_block = om.PeakShaving
        peak_kosten_je_monat = {
            m: po.value(peak_block.peak[m]) * params["netznutzung_leistung"] for m in peak_block.MONATE
        }

    srl_pos_kw_opt = srl_neg_kw_opt = None
    if srl_variante == "optimiert":
        srl_block = om.SRLReserve
        srl_pos_kw_opt = np.array([po.value(srl_block.r_pos[t]) for t in om.TIMESTEPS])
        srl_neg_kw_opt = np.array([po.value(srl_block.r_neg[t]) for t in om.TIMESTEPS])

    remove_custom_blocks(om)

    print("Werte Ergebnisse aus ...")
    results = solph.processing.results(om)
    ts = extract_timeseries(results, node_refs, zeitreihen)
    module, anteile = allocate_modules(
        ts, zeitreihen, params, dt_hours, srl_variante, soc_min_srl, soc_max_srl,
        last_profile=last_profile,
    )

    pv_abregelung_kwh = float(ts["pv_abregelung"].sum() * dt_hours)
    if pv_abregelung_kwh > 1e-6:
        print(
            f"HINWEIS: {pv_abregelung_kwh:,.0f} kWh/Jahr PV-Erzeugung mussten "
            f"abgeregelt werden (siehe snk_curtailment in build_energysystem()) "
            f"-- PV-Erzeugung uebersteigt in manchen Zeitschritten die Summe aus "
            f"dc_leistung (Parameter!C16) und Batterie-Ladeleistung (Parameter!C21)."
        )

    if peakshaving_aktiv:
        # Monat je Zeitschritt -- gleiche Zuordnung wie in add_peak_shaving().
        t_month = {t: (zeitreihen.index[t].year, zeitreihen.index[t].month) for t in om.TIMESTEPS}
        peak_kosten = sum(peak_kosten_je_monat.values())

        # Beat's Wunsch: das Peak-Shaving-Modul soll nicht die rohen Kosten
        # ausweisen, sondern die ERSPARNIS ggue. dem Referenzfall OHNE
        # Batterie -- das ist der eigentliche "Gewinn" des Peak-Shaving.
        # Referenzfall: ohne Batterie waere der Netzbezug schlicht die
        # ungedeckte Last (Last - PV), nach unten bei 0 gekappt (ueberschuess-
        # ige PV wuerde exportiert, nicht den Bezug negativ machen).
        netzbezug_baseline = np.maximum(last_profile.values - pv_profile.values, 0.0)
        baseline_series = pd.Series(netzbezug_baseline, index=zeitreihen.index)
        baseline_peak_je_monat = baseline_series.groupby(
            lambda z: (z.year, z.month)
        ).max()
        peak_kosten_ohne_batterie = float(
            sum(baseline_peak_je_monat[m] * params["netznutzung_leistung"]
                for m in baseline_peak_je_monat.index)
        )
        peak_ersparnis = peak_kosten_ohne_batterie - peak_kosten

        module["Peak-Shaving"] = peak_ersparnis  # Ersparnis = Modul-Ertrag

        print("\n=== Peak-Shaving: Ersparnis ggue. Referenzfall (ohne Batterie) ===")
        print(f"  {'Bezugsspitzen-Kosten OHNE Batterie (Referenz)':50s} {peak_kosten_ohne_batterie:>12,.0f} CHF/Jahr")
        print(f"  {'Bezugsspitzen-Kosten MIT Batterie (optimiert)':50s} {peak_kosten:>12,.0f} CHF/Jahr")
        print(f"  {'--> Ersparnis durch Peak-Shaving (= Modul-Ertrag)':50s} {peak_ersparnis:>12,.0f} CHF/Jahr")

        # Fuer den Output/Excel-Detailspalte: gleicher Monatswert auf allen
        # Zeitschritten des jeweiligen Monats -- bleibt bewusst die ECHTEN,
        # tatsaechlich anfallenden Kosten (negativ), NICHT die Ersparnis, da
        # diese Spalte zeigt, was effektiv bezahlt wird. Die Ersparnis ist
        # ausschliesslich der Modul-Ertragswert oben.
        anteile["peak_shaving_kosten_monat"] = [
            -peak_kosten_je_monat[t_month[t]] for t in om.TIMESTEPS
        ]

    if srl_variante == "optimiert":
        ertrag_srl = (
            srl_pos_kw_opt * zeitreihen["srl_pos"].values + srl_neg_kw_opt * zeitreihen["srl_neg"].values
        ) * dt_hours
        anteile["srl_pos_kw"] = srl_pos_kw_opt
        anteile["srl_neg_kw"] = srl_neg_kw_opt
        anteile["ertrag_srl"] = ertrag_srl
        module["SRL (Teil der Optimierung)"] = float(np.sum(ertrag_srl))

    total = sum(module.values())
    module["Total (operativ)"] = total

    print("\n=== Ertrag pro Modul (CHF/Jahr) ===")
    for k, v in module.items():
        print(f"  {k:45s} {v:>12,.0f}")

    # --- Arbitrage-Diagnose ---------------------------------------------------
    # Hilft zu unterscheiden, ob "Arbitrage der Batterie" == 0 ein Bug ist oder
    # eine legitime Optimierungs-Entscheidung. FRUEHERE VERSION dieser Diagnose
    # nannte zwei Gruende (Tarifspread < Rundwirkungsgrad-Verlust; Vollzyklen-
    # Limit durch PV ausgeschoepft) -- beide wurden an Beats echten Zahlen
    # WIDERLEGT (Tarifspread war riesig [-0.24 bis +0.30 CHF/kWh, weit ueber
    # dem ~24% Rundwirkungsgrad-Verlust]; Vollzyklen-Limit lag bei >100 Mio.
    # kWh/Jahr, wo real nur ~39'000 kWh/Jahr Entladeenergie anfielen -- also
    # bei weitem nicht ausgeschoepft). Beats eigene Vermutung ("bei tiefen/
    # negativen Preisen war Solar da -> Batterie schon durch PV geladen, kein
    # Platz/Bedarf fuer zusaetzliches Netz-Laden") ist plausibler -- die beiden
    # Pruefungen unten testen das DIREKT an den Daten, statt weiter zu raten.
    eta_laden = params["lade_wirkungsgrad"] * params["trafo_wirkungsgrad"]
    eta_entladen = params["entlade_wirkungsgrad"] * params["trafo_wirkungsgrad"]
    netz_ladeenergie_kwh = float(ts["ch_grid"].sum() * dt_hours)
    pv_ladeenergie_kwh = float(ts["ch_pv"].sum() * dt_hours)
    entladeenergie_kwh = float(ts["dis"].sum() * dt_hours)
    vollzyklen_limit_kwh = params["durchsatz_mwh_jahr"] * 1000.0
    rundwirkungsgrad = eta_laden * eta_entladen

    print("\n=== Arbitrage-Diagnose ===")
    print(f"  {'Netz-Ladeenergie (Batterie aus Netz geladen)':45s} {netz_ladeenergie_kwh:>12,.0f} kWh/Jahr")
    print(f"  {'PV-Ladeenergie (Batterie aus PV geladen)':45s} {pv_ladeenergie_kwh:>12,.0f} kWh/Jahr")
    print(f"  {'Entladeenergie total':45s} {entladeenergie_kwh:>12,.0f} kWh/Jahr")
    print(f"  {'Vollzyklen-Limit (Durchsatz/Jahr)':45s} {vollzyklen_limit_kwh:>12,.0f} kWh/Jahr "
          f"({'NICHT ausgeschoepft' if entladeenergie_kwh < 0.9 * vollzyklen_limit_kwh else 'FAST ausgeschoepft'})")
    print(f"  {'Rundwirkungsgrad Batterie (Laden*Entladen)':45s} {rundwirkungsgrad:>12.1%}")
    print(f"  {'Bezugstarif Minimum':45s} {zeitreihen['bezugstarif'].min():>12.3f} CHF/kWh")
    print(f"  {'Bezugstarif Maximum':45s} {zeitreihen['bezugstarif'].max():>12.3f} CHF/kWh")

    if netz_ladeenergie_kwh < 1.0:
        # Direkter Datencheck statt weiterer Spekulation: SOC-Fuellstand und
        # PV-Erzeugung genau WAEHREND der Zeitschritte mit negativem
        # Bezugstarif (den attraktivsten Netz-Arbitrage-Momenten).
        kapazitaet = params["kapazitaet"]
        soc_pct = (ts["soc"].values / kapazitaet) if kapazitaet else np.zeros(len(ts))
        neg_preis_mask = zeitreihen["bezugstarif"].values < 0
        anz_neg_preis = int(neg_preis_mask.sum())

        print(
            "\n  -> Die Batterie wurde im optimalen Ergebnis (praktisch) NIE aus dem "
            "Netz geladen -- daher Arbitrage-Ertrag ~0. Direkter Datencheck fuer "
            "die Zeitschritte mit negativem Bezugstarif (die attraktivsten "
            "Arbitrage-Momente):"
        )
        if anz_neg_preis == 0:
            print("     Keine Zeitschritte mit negativem Bezugstarif in diesem Lauf gefunden.")
        else:
            soc_bei_neg = soc_pct[neg_preis_mask]
            pv_bei_neg = pv_profile.values[neg_preis_mask]
            anteil_voll = float(np.mean(soc_bei_neg >= params["ladegrenze_soc"] - 0.01))
            print(f"     Zeitschritte mit negativem Bezugstarif: {anz_neg_preis} von {len(zeitreihen)}")
            print(f"     Durchschnittlicher SOC dabei: {np.mean(soc_bei_neg):.0%} (Ladegrenze/C30: {params['ladegrenze_soc']:.0%})")
            print(f"     Anteil davon mit SOC bereits an/nahe der Ladegrenze (kein Platz mehr): {anteil_voll:.0%}")
            print(f"     Durchschnittliche PV-Erzeugung dabei: {np.mean(pv_bei_neg):.1f} kW (Jahresdurchschnitt: {pv_profile.values.mean():.1f} kW)")
            print(
                "     Hoher Anteil 'SOC nahe Ladegrenze' und/oder deutlich "
                "ueberdurchschnittliche PV-Erzeugung hier stuetzen Beats Vermutung: "
                "die Batterie war zu genau diesen Zeitpunkten durch PV bereits "
                "(fast) voll -- kein physischer Platz fuer zusaetzliches "
                "Netz-Laden, unabhaengig vom Preis."
            )

    # --- Kapitalkosten/Unterhalt (informativ, nicht Teil der LP-Zielfunktion) --
    # entladeenergie_kwh (siehe Arbitrage-Diagnose oben) wird hier mitgegeben,
    # damit die Degradationskosten (siehe DEGRADATIONSKOSTEN_CHF_PRO_KWH),
    # die schon IN der LP-Zielfunktion stecken, auch in dieser Uebersicht und
    # im PDF-Report als eigene Zeile sichtbar werden.
    kk = capital_costs(params, entladeenergie_kwh=entladeenergie_kwh)

    print(f"\n=== Kapitalkosten/Unterhalt (informativ, Annahme WACC={kk['wacc']:.0%}) ===")
    print(f"  {'Investitions-Annuitaet':45s} {-kk['annuitaet']:>12,.0f}")
    print(f"  {'Unterhaltkosten':45s} {-kk['unterhalt']:>12,.0f}")
    print(f"  {'Degradationskosten (' + str(DEGRADATIONSKOSTEN_CHF_PRO_KWH) + ' CHF/kWh entladen)':45s} {-kk['degradation']:>12,.0f}")
    print(f"  {'Netto-Ergebnis (operativ - Kapital/Unterhalt/Degradation)':45s} "
          f"{total - kk['annuitaet'] - kk['unterhalt'] - kk['degradation']:>12,.0f}")

    # pv_profile/last_profile/zeitreihen/params zusaetzlich zurueckgeben --
    # output_create.py braucht sie fuer die Output-Kontroll-Zeitreihe
    # (Energiepreis, PV-Erzeugung, Last), ohne alles selbst neu berechnen
    # zu muessen.
    return module, ts, anteile, zeitreihen, params, pv_profile, last_profile


if __name__ == "__main__":
    main()