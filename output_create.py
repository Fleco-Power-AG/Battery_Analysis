"""
output_create.py
=================

Baut die "Output-Kontrolle" aus den Ergebnissen von
`battery_optimization.main()`: eine Zeitreihen-Tabelle (eine Zeile pro
15-Min-Intervall) plus eine Modul-Ertrags-Zusammenfassung (CHF/Jahr),
geschrieben nach einer vom Input-Dateinamen abgeleiteten Ergebnis-Datei (z.B.
`Inputs_default.xlsx` -> `Ergebnis_default.xlsx`, siehe
`derive_result_filename()`). Wird von `run_battery_analysis.py` aufgerufen --
kann aber auch direkt mit den Rueckgabewerten von `battery_optimization.main()`
genutzt werden.

Spalten der Zeitreihen-Tabelle (siehe `build_output_table()`):
  - Index "Zeit" (15-Min-Zeitstempel)
  - Energiepreis_Bezug_CHF_kWh / Energiepreis_Rueckliefer_CHF_kWh
  - PV_Erzeugung_kW, Last_kW, Netzbezug_Total_kW (zur Kontrolle der
    Energiebilanz -- nicht explizit angefragt, aber praktisch fuer die
    Verifikation)
  - PV_Abregelung_kW: Sicherheitsventil in battery_optimization.py (siehe
    snk_curtailment) fuer PV-Erzeugung, die in einem Zeitschritt die Summe aus
    dc_leistung und Batterie-Ladeleistung uebersteigt -- im Normalfall
    ueberall 0; falls nicht, lohnt sich ein Blick auf Parameter!C16/C21.
  - Batterieleistung -- alle Lade-/Entlade-Fluesse einzeln:
    Batterie_Laden_PV_kW, Batterie_Laden_Netz_kW, Batterie_Entladen_kW,
    plus Batterie_Leistung_Netto_kW (Entladen minus beide Ladefluesse)
  - SOC_kWh und SOC_Prozent
  - SRL_Leistung_Pos_kW / SRL_Leistung_Neg_kW (freie Regelleistung; bei
    SRL-Variante "keine" ueberall 0)
  - Ertragsspalten je Zeitschritt: Ertrag_Eigenverbrauch_CHF,
    Ertrag_Arbitrage_CHF, Ertrag_SRL_CHF (Anfangsbestand-Verwertung wird auf
    Beats Wunsch NIRGENDS mehr ausgewiesen -- weder hier noch im
    Modul-Ertraege-Sheet noch im PDF-Report)
  - Ertrag_PeakShaving_Monat_CHF: ACHTUNG, das ist eine monatliche
    Grundgebuehr (Kosten der Monatsspitze), kein echter Pro-Intervall-Fluss --
    der gleiche Monatswert steht auf allen Zeilen des jeweiligen Monats,
    damit die Spalte trotzdem in der Zeitreihen-Tabelle auftaucht (0, falls
    Peak-Shaving inaktiv ist).

Die Modul-Ertrags-Zusammenfassung (`build_module_summary_table()`) fasst
dieselben Groessen auf CHF/Jahr zusammen (das bisherige "Modul-Ertraege"-Sheet).

Zusaetzlich baut `build_pdf_report()` (siehe unten) daraus einen mehrseitigen
PDF-Report: Wirtschaftlichkeit (Einnahmen/Ausgaben/Gewinn-Verlust),
Renditebetrachtung (LCOS/Amortisationsdauer/Kapitalverzinsung), sowie drei
Charts (Eigenverbrauch/Rueckspeisung pro Monat, Nettolastprofil je Quartal,
Lastspitze pro Monat). Wird ebenfalls von run_battery_analysis.py aufgerufen.

STANDALONE-AUFRUF (`python output_create.py [Pfad-zu-input_lp.xlsx]`):
Dieses Skript kann auch alleine laufen, OHNE ueber run_battery_analysis.py zu
gehen -- praktisch zum schnellen Iterieren an Optimierung/Output, ohne jedes
Mal input_prep.py (und damit die ENTSO-E-Abfrage) neu laufen zu lassen. Dafuer
braucht es aber ein bereits vorhandenes `input_lp.xlsx` (das Ergebnis von
input_prep.py) -- ohne das schlaegt es fehl, weil `battery_optimization.main()`
genau diese Datei einliest. `ergebnis.xlsx` landet dann im selben Ordner wie
das verwendete `input_lp.xlsx`. Ohne Pfad-Argument oeffnet sich ein
Datei-Auswahl-Dialog (wie bei run_battery_analysis.py).
"""

from __future__ import annotations

import argparse
import datetime as dt
import os
import re

import numpy as np
import pandas as pd


# Erkennt ein fuehrendes 'Input'/'Inputs' ODER bereits 'input_lp' (Gross-/
# Kleinschreibung egal, optional gefolgt von '_'/'-') am Anfang des
# Dateinamens (ohne Endung). Gemeinsame Basis fuer alle drei abgeleiteten
# Dateinamen unten (input_lp/Ergebnis/Bericht) -- so liefert ein Lauf ueber
# eine bereits abgeleitete 'input_lp_default.xlsx' dieselbe Zuordnung
# ('default') wie ein Lauf ueber die urspruengliche 'Inputs_default.xlsx'.
_INPUT_PREFIX_RE = re.compile(r"^input(?:s|_lp)?[_-]?(.*)$", re.IGNORECASE)


def _strip_input_prefix(input_path: str) -> str:
    base = os.path.splitext(os.path.basename(input_path))[0]
    match = _INPUT_PREFIX_RE.match(base)
    return (match.group(1) if match else base).strip("_- ")


def derive_input_lp_filename(input_path: str) -> str:
    """Leitet den Namen der aufbereiteten Zwischen-Datei (Ergebnis von
    input_prep.py, Basis fuer battery_optimization.py) vom Namen der
    Input-Datei ab, damit sie sich eindeutig dem Szenario zuordnen laesst und
    man bei mehreren Szenarien im selben Ordner nicht jedes Mal dieselbe
    'input_lp.xlsx' ueberschreibt:

        Inputs_default.xlsx  -> input_lp_default.xlsx
        batterie_4.xlsx      -> input_lp_batterie_4.xlsx
        Inputs.xlsx          -> input_lp.xlsx  (kein Rest nach Praefix-Entfernen)
    """
    rest = _strip_input_prefix(input_path)
    return f"input_lp_{rest}.xlsx" if rest else "input_lp.xlsx"


def derive_result_filename(input_path: str) -> str:
    """Leitet den Namen der Ergebnis-Datei vom Namen der Input-Datei ab, damit
    man bei mehreren Szenarien nicht jedes Mal 'ergebnis.xlsx' ueberschreibt:

        Inputs_default.xlsx  -> Ergebnis_default.xlsx
        batterie_4.xlsx      -> Ergebnis_batterie_4.xlsx
        Inputs.xlsx          -> Ergebnis.xlsx  (kein Rest nach Praefix-Entfernen)

    Ein fuehrendes 'Input'/'Inputs' (Gross-/Kleinschreibung egal, optional
    gefolgt von '_' oder '-') wird entfernt; der Rest des Dateinamens wird an
    'Ergebnis_' angehaengt. Bleibt kein Rest (z.B. bei genau 'Inputs.xlsx'),
    heisst die Datei schlicht 'Ergebnis.xlsx'.
    """
    rest = _strip_input_prefix(input_path)
    return f"Ergebnis_{rest}.xlsx" if rest else "Ergebnis.xlsx"


def derive_pdf_filename(input_path: str) -> str:
    """Analog zu derive_result_filename(), aber fuer den PDF-Report:

        Inputs_default.xlsx  -> Bericht_default.pdf
        batterie_4.xlsx      -> Bericht_batterie_4.pdf
        Inputs.xlsx          -> Bericht.pdf
    """
    rest = _strip_input_prefix(input_path)
    return f"Bericht_{rest}.pdf" if rest else "Bericht.pdf"


# --------------------------------------------------------------------------
# Anzeige-Zeitzone: Excel/PDF sollen ECHTE Zuercher Ortszeit (mit Sommerzeit-
# Umstellung) zeigen, siehe _to_zurich_local() unten.
# --------------------------------------------------------------------------
_ZURICH_TZ = "Europe/Zurich"


def _to_zurich_local(*objs):
    """Konvertiert eine oder mehrere zeitindizierte DataFrames/Series (ts,
    anteile, zeitreihen, pv_profile, last_profile) von der intern verwendeten
    FESTEN CET-Zeitzone (UTC+1 ohne Sommerzeit-Wechsel -- siehe CET_FIXED in
    input_prep.py/battery_optimization.py, entspricht 1:1 der rohen
    Excel-Spalte "Zeit") auf ECHTE Zuercher Ortszeit (Europe/Zurich, MIT
    Sommerzeit-Umstellung) fuer die ANZEIGE in Excel-Export und PDF-Charts.

    Wichtig: das ist eine reine Umbenennung des Zeitstempel-LABELS, keine
    Verschiebung von Werten -- `tz_convert()` aendert den zugrunde liegenden
    Zeitpunkt (Instant) nicht, nur seine Darstellung. Reihenfolge und Anzahl
    der 15-Min-Zeitschritte bleiben exakt gleich; im Sommer verschiebt sich
    nur das angezeigte Label um +1h (CEST=UTC+2) gegenueber der bisherigen
    festen CET-Anzeige (UTC+1), im Winter bleibt es unveraendert (CET=UTC+1
    entspricht in diesem Halbjahr ohnehin der echten Ortszeit). Betrifft NUR
    die Ausgabe (Excel-Zeitspalte, Monats-/Quartals-/Tagesprofil-Gruppierung
    in den Charts) -- die LP-Optimierung selbst (battery_optimization.py)
    bleibt davon komplett unberuehrt, dort zaehlt nur die Reihenfolge der
    Zeitschritte, nicht ihr Kalenderlabel.

    Gibt fuer jedes uebergebene Objekt eine KOPIE mit umgestelltem Index
    zurueck (Originale bleiben unveraendert), in derselben Reihenfolge wie
    uebergeben.
    """
    out = []
    for obj in objs:
        obj2 = obj.copy()
        if obj2.index.tz is not None:
            obj2.index = obj2.index.tz_convert(_ZURICH_TZ)
        out.append(obj2)
    return out


def _valid_calendar_groups(group_keys, min_fraction: float = 0.05, min_absolute: int = 24) -> set:
    """Bestimmt, welche Monats-/Quartals-Gruppen genug Zeitschritte haben, um
    sinnvoll in einem Monats-/Quartals-Chart dargestellt zu werden (Beats
    Frage "wieso hat der Report Q3-Werte, obwohl das Ganze nur bis 30.6.
    23:45 geht"):

    Die Umstellung auf echte Zuercher Ortszeit MIT Sommerzeit (siehe
    _to_zurich_local()) verschiebt Zeitstempel in der Sommerzeit-Haelfte des
    Jahres um +1h. Endet ein Szenario NICHT exakt am Jahresende (z.B. Beats
    Q2-Szenario, das am 30.6. endet -- mitten in der Sommerzeit), rutschen
    dadurch die letzten paar (max. 4, bei 15-Min-Intervallen bis zu einer
    vollen Stunde) Zeitschritte formal in den naechsten Kalendermonat/
    -quartal (30.6. 23:00-23:45 CET-fix -> 1.7. 00:00-00:45 echte
    Zuercher Ortszeit). Ohne Filter erzeugt das einen fast leeren "Geister"-
    Balken/Linie mit nur 1-4 Datenpunkten -- z.B. einen wilden Ausschlag im
    Tagesprofil-Chart, weil der "Durchschnitt" aus nur 4 Werten statt aus
    Tausenden berechnet wird.

    Eine Gruppe gilt als "echt" (wird behalten), wenn sie mindestens
    `min_absolute` Zeitschritte ODER mindestens `min_fraction` der groessten
    Gruppe in derselben Tabelle hat. Gibt es nur EINE Gruppe insgesamt (z.B.
    ein bewusst kurzes Test-Szenario von nur ein paar Tagen), wird NICHTS
    gefiltert -- der Vergleichsmassstab (relative Groesse) macht dort keinen
    Sinn, und ein Sicherheitsnetz verhindert, dass jemals ALLE Gruppen
    herausgefiltert werden."""
    counts = pd.Series(group_keys).value_counts()
    if len(counts) <= 1:
        return set(counts.index)
    schwelle = max(min_absolute, min_fraction * counts.max())
    keep = set(counts[counts >= schwelle].index)
    return keep if keep else set(counts.index)


def build_output_table(
    ts: pd.DataFrame,
    anteile: pd.DataFrame,
    zeitreihen: pd.DataFrame,
    params: dict,
    pv_profile: pd.Series,
    last_profile: pd.Series,
    dt_hours: float = 0.25,
) -> pd.DataFrame:
    """Baut die Zeitreihen-Output-Tabelle (eine Zeile pro 15-Min-Intervall).

    ts, anteile: Rueckgabewerte von battery_optimization.main() (bzw. von
        extract_timeseries()/allocate_modules() direkt).
    zeitreihen, params, pv_profile, last_profile: ebenfalls von
        battery_optimization.main() zurueckgegeben (Tarife, Parameter,
        gebautes PV-/Lastprofil).
    """
    # Excel/openpyxl kann keine tz-aware Zeitstempel schreiben ("Excel does
    # not support datetimes with timezones"). ts.index ist zu diesem Zeitpunkt
    # bereits per _to_zurich_local() (siehe save_output_excel()) auf ECHTE
    # Zuercher Ortszeit MIT Sommerzeit-Umstellung umgestellt -- hier wird nur
    # noch die tz-Information selbst entfernt (die Uhrzeiten/Werte sind
    # dadurch bereits korrekt DST-verschoben, das faellt beim Entfernen der
    # tz-Markierung nicht mehr weg).
    #
    # NEU (Beats Wunsch): da die "Zeit"-Spalte selbst keine tz-Info tragen
    # kann, wird VOR dem Entfernen der tz-Markierung eine eigene Spalte
    # "Zeitzone" gebaut (z.B. "CET (UTC+01:00)" im Winter, "CEST
    # (UTC+02:00)" im Sommer) -- damit ist pro Zeile eindeutig ersichtlich,
    # welche Zeitzone/welcher Offset fuer den jeweiligen Zeitstempel galt
    # (v.a. an den beiden Umstellungstagen im Maerz/Oktober relevant).
    idx = ts.index
    if idx.tz is not None:
        zeitzone_abbr = idx.strftime("%Z")
        zeitzone_offset = idx.strftime("%z")  # z.B. "+0100" (ohne Doppelpunkt)
        zeitzone_spalte = [
            f"{abbr} (UTC{off[:3]}:{off[3:]})" for abbr, off in zip(zeitzone_abbr, zeitzone_offset)
        ]
        idx = idx.tz_localize(None)
    else:
        zeitzone_spalte = [""] * len(idx)
    kapazitaet = params["kapazitaet"]

    soc_kwh = ts["soc"]
    soc_pct = (soc_kwh / kapazitaet * 100.0) if kapazitaet else soc_kwh * 0.0

    batterie_laden_pv = ts["ch_pv"]
    batterie_laden_netz = ts["ch_grid"]
    batterie_entladen = ts["dis"]
    batterie_netto = batterie_entladen - batterie_laden_pv - batterie_laden_netz

    out = pd.DataFrame(
        {
            "Zeitzone": zeitzone_spalte,
            "Energiepreis_Bezug_CHF_kWh": zeitreihen["bezugstarif"].values,
            "Energiepreis_Rueckliefer_CHF_kWh": zeitreihen["rueckliefertarif"].values,
            "PV_Erzeugung_kW": pv_profile.values,
            "Last_kW": last_profile.values,
            "Netzbezug_Total_kW": ts["netzbezug_total"].values,
            # NEU: last-spezifischer Anteil des Netzbezugs (nur was ueber
            # link_netz_ac Richtung Last fliesst, siehe battery_optimization.py
            # extract_timeseries()) -- Differenz zu Netzbezug_Total_kW ist die
            # reine Batterie-Netzladung (Arbitrage), die mit der Last selbst
            # nichts zu tun hat. ".get"-Fallback fuer aeltere ts-DataFrames.
            "Netzbezug_Last_kW": ts["netzbezug_last"].values if "netzbezug_last" in ts.columns else np.zeros(len(idx)),
            "Rueckspeisung_kW": ts["export"].values,
            # PV-Abregelung (Sicherheitsventil in battery_optimization.py, siehe
            # snk_curtailment) -- im Normalfall ueberall 0, ".get" als Fallback
            # fuer aeltere ts-DataFrames ohne diese Spalte.
            "PV_Abregelung_kW": ts["pv_abregelung"].values if "pv_abregelung" in ts.columns else np.zeros(len(idx)),
            "Batterie_Laden_PV_kW": batterie_laden_pv.values,
            "Batterie_Laden_Netz_kW": batterie_laden_netz.values,
            "Batterie_Entladen_kW": batterie_entladen.values,
            "Batterie_Leistung_Netto_kW": batterie_netto.values,
            "SOC_kWh": soc_kwh.values,
            "SOC_Prozent": soc_pct.values,
            "SRL_Leistung_Pos_kW": anteile["srl_pos_kw"].values,
            "SRL_Leistung_Neg_kW": anteile["srl_neg_kw"].values,
            "Ertrag_Eigenverbrauch_CHF": anteile["ertrag_eigenverbrauch"].values,
            # NEU (Beats Wunsch, EV/Einspeise/Arbitrage nach Ziel+Herkunft):
            # Herkunfts-Detail hinter Ertrag_Eigenverbrauch_CHF (Summe der
            # beiden = Ertrag_Eigenverbrauch_CHF), plus die neue Einspeise-
            # optimierung-Spalte. ".get"-Fallback fuer aeltere anteile-
            # DataFrames ohne diese Spalten.
            "Ertrag_Eigenverbrauch_PV_CHF": (
                anteile["ertrag_eigenverbrauch_pv"].values
                if "ertrag_eigenverbrauch_pv" in anteile.columns else np.zeros(len(idx))
            ),
            "Ertrag_Eigenverbrauch_Netz_CHF": (
                anteile["ertrag_eigenverbrauch_grid"].values
                if "ertrag_eigenverbrauch_grid" in anteile.columns else np.zeros(len(idx))
            ),
            "Ertrag_Einspeiseoptimierung_CHF": (
                anteile["ertrag_einspeiseoptimierung"].values
                if "ertrag_einspeiseoptimierung" in anteile.columns else np.zeros(len(idx))
            ),
            "Ertrag_Arbitrage_CHF": anteile["ertrag_arbitrage"].values,
            "Ertrag_SRL_CHF": anteile["ertrag_srl"].values,
            "Ertrag_PeakShaving_Monat_CHF": anteile["peak_shaving_kosten_monat"].values,
        },
        index=idx,
    )
    out.index.name = "Zeit"
    return out


# Feste, lesbare Reihenfolge fuer die Modul-Zusammenfassung. Module, die fuer
# die aktuelle SRL-Variante nicht zutreffen (z.B. "SRL (Post-hoc-Mehrertrag)",
# wenn tatsaechlich Variante "optimiert" aktiv war), stehen im `module`-Dict
# ohnehin nicht drin und werden hier automatisch uebersprungen.
MODUL_REIHENFOLGE = [
    "Eigenverbrauchsoptimierung",
    "Einspeiseoptimierung",
    "Arbitrage der Batterie",
    "SRL (Post-hoc-Mehrertrag)",
    "SRL (Teil der Optimierung)",
    "Peak-Shaving",
    "Total (operativ)",
]


def build_module_summary_table(module: dict, spalten_label: str = "CHF_pro_Jahr") -> pd.DataFrame:
    """Modul-Ertraege (Ist-Summe ueber den simulierten Zeitraum) in fester,
    lesbarer Reihenfolge.

    `spalten_label` (NEU): die Spalte hiess bisher IMMER "CHF_pro_Jahr",
    obwohl `module[...]` tatsaechlich die reine IST-SUMME ueber den
    simulierten Zeitraum ist (nicht auf ein Jahr hochgerechnet) -- bei einem
    Teil-Jahr-Szenario (z.B. nur Q2) war das irrefuehrend beschriftet.
    save_output_excel() setzt hier bei einem Nicht-Jahres-Zeitraum ein
    passendes Label (z.B. "CHF_pro_Zeitraum_91_Tage").
    """
    reihenfolge = [k for k in MODUL_REIHENFOLGE if k in module]
    # Falls kuenftig ein neuer/unerwarteter Modul-Schluessel dazukommt, trotzdem
    # anzeigen statt stillschweigend zu verschlucken.
    reihenfolge += [k for k in module if k not in reihenfolge]
    return pd.DataFrame(
        {spalten_label: [module[k] for k in reihenfolge]}, index=reihenfolge
    )


# --------------------------------------------------------------------------
# Wirtschaftlichkeit / Rendite (fuer den PDF-Report, Teil 1+2)
# --------------------------------------------------------------------------

# Interne Modul-Schluessel (aus battery_optimization.allocate_modules()) auf
# die im PDF-Report gewuenschten Anzeige-Labels gemappt. "Total (operativ)"
# ist ABSICHTLICH nicht drin (wird hier separat berechnet). Anfangsbestand-
# Verwertung existiert seit Beats Wunsch gar nicht mehr als eigenes Modul --
# weder hier noch in ergebnis.xlsx noch im PDF-Report (siehe
# battery_optimization.allocate_modules()).
EINNAHMEN_LABELS = {
    "Eigenverbrauchsoptimierung": "Eigenverbrauchsoptimierung",
    "Einspeiseoptimierung": "Einspeiseoptimierung",
    "Peak-Shaving": "Peak-Shaving",
    "Arbitrage der Batterie": "Arbitrage",
    "SRL (Post-hoc-Mehrertrag)": "SRL",
    "SRL (Teil der Optimierung)": "SRL",
}


def _scenario_dauer(start: pd.Timestamp, n: int, dt_hours: float = 0.25) -> tuple[float, str]:
    """Ermittelt die TATSAECHLICH simulierte Dauer (Bruchteil eines Jahres,
    365.25 Tage) sowie einen passenden Titel fuer den Wirtschaftlichkeits-
    Abschnitt des PDF-Reports.

    Hintergrund (Beats Frage "wieso sagt der Report eine Jahresbetrachtung,
    obwohl es nur 3 Monate sind"): bisher stand im PDF-Report IMMER
    "Wirtschaftlichkeit — Jahresbetrachtung", unabhaengig davon, wie viele
    Zeitschritte tatsaechlich simuliert wurden -- fuer ein Teil-Jahr-Szenario
    (z.B. "..._Q2.xlsx", nur April-Juni) war das schlicht falsch, UND (siehe
    build_income_expense_summary()/build_rendite_kennzahlen()) die
    zugrunde liegende Rechnung war entsprechend inkonsistent (volle
    Jahres-Ausgaben gegen nur 3 Monate Einnahmen).

    `start`/`n` MUESSEN der urspruengliche (fest-CET, VOR _to_zurich_local())
    Start-Zeitstempel und die Zeitschritt-Anzahl des Szenarios sein, nicht der
    bereits auf echte Zuercher Ortszeit umgerechnete Index -- sonst wuerde der
    Sommerzeit-Versatz (+1h, siehe _to_zurich_local()) das angezeigte
    Enddatum systematisch um einen Tag nach vorne verschieben (z.B. "endet am
    01.07." statt korrekt "30.06.", obwohl das Szenario laut Eingabedatei
    exakt am 30.06. 23:45 endet). Fuer den reinen Kalender-Zeitraum im Titel
    ist die (DST-freie) Eingabedatei-Zeitangabe die richtige Referenz -- die
    Zuercher-Lokalzeit-Umrechnung ist nur fuer die Chart-Achsen/-Gruppierung
    gedacht (siehe _to_zurich_local()-Docstring), nicht fuer diese Anzeige.

    Dauer wird aus der Anzahl Zeitschritte * Intervall-Laenge berechnet
    (robust gegen Schaltjahre/Sommerzeit-Rundung). Liegt die Dauer sehr nah an
    einem vollen Jahr (± 3 Tage Toleranz, z.B. wegen Schaltjahr), bleibt es
    beim gewohnten Titel "Jahresbetrachtung"; sonst zeigt der Titel den
    TATSAECHLICHEN Kalenderzeitraum (erster bis letzter simulierter Tag) plus
    Anzahl Tage, z.B. "Wirtschaftlichkeit — 01.04.–30.06.2026 (91 Tage)".
    """
    dauer_tage = n * dt_hours / 24.0
    dauer_jahre = dauer_tage / 365.25
    if abs(dauer_tage - 365.25) <= 3 or abs(dauer_tage - 366) <= 3:
        titel = "Wirtschaftlichkeit — Jahresbetrachtung"
    else:
        ende_ts = start + pd.Timedelta(hours=(n - 1) * dt_hours)
        ende = ende_ts.strftime("%d.%m.%Y")
        # Kurzes Format (kein Jahr doppelt, wenn Start/Ende im selben Jahr
        # liegen) -- verhindert, dass der Titel im PDF am rechten Rand
        # abgeschnitten wird.
        if start.year == ende_ts.year:
            start_kurz = start.strftime("%d.%m.")
            titel = f"Wirtschaftlichkeit — {start_kurz}–{ende} ({dauer_tage:.0f} Tage)"
        else:
            start_str = start.strftime("%d.%m.%Y")
            titel = f"Wirtschaftlichkeit — {start_str}–{ende} ({dauer_tage:.0f} Tage)"
    return dauer_jahre, titel


def build_income_expense_summary(module: dict, kapitalkosten: dict, dauer_jahre: float = 1.0) -> dict:
    """Wirtschaftlichkeits-Uebersicht fuer den PDF-Report, Teil 1:
      - Einnahmen/Ersparnisse: NUR die tatsaechlich im Excel aktivierten
        Module (Eigenverbrauch ist immer dabei; Peak-Shaving/SRL nur, wenn
        der jeweilige Schluessel im `module`-Dict vorkommt -- das ist bereits
        durch battery_optimization.py sichergestellt: inaktive Module tauchen
        dort gar nicht erst auf). Anfangsbestand-Verwertung ist seit Beats
        Wunsch kein Modul-Schluessel mehr und kann daher gar nicht erst
        auftauchen (siehe battery_optimization.allocate_modules()). Diese
        Summe ist immer die TATSAECHLICHE Summe ueber den simulierten
        Zeitraum (egal ob 1 Jahr oder kuerzer).
      - Ausgaben: Amortisation, Kapitalkosten (Zins), Unterhalt -- aus
        `kapitalkosten` (Rueckgabewert von battery_optimization.capital_costs()),
        die dort IMMER als volle JAHRES-Groessen berechnet werden (unabhaengig
        von der tatsaechlich simulierten Dauer). Damit diese Ausgaben mit den
        (nur ueber den simulierten Zeitraum summierten) Einnahmen vergleichbar
        bleiben, werden sie hier mit `dauer_jahre` (Bruchteil eines Jahres,
        siehe _scenario_dauer()) SKALIERT -- bei einem 3-Monats-Szenario also
        z.B. nur 1/4 der Jahres-Amortisation/-Kapitalkosten/-Unterhalt.
        Batterie-Degradation wird NICHT skaliert, da sie bereits aus der
        tatsaechlichen (nicht-annualisierten) Entladeenergie des simulierten
        Zeitraums berechnet ist (siehe battery_optimization.capital_costs()).
      - Gewinn/Verlust: Total Einnahmen - Total Ausgaben (beide jetzt auf
        denselben, tatsaechlich simulierten Zeitraum bezogen).

    `kapitalkosten` als Parameter (statt hier neu zu berechnen), damit diese
    Funktion selbst kein battery_optimization/oemof braucht -- der Aufrufer
    (build_pdf_report()/save_output_excel()) importiert battery_optimization
    ohnehin schon lazy und reicht das fertige Dict durch.

    NEU (Beats Frage "wieso sagt der Report eine Jahresbetrachtung, obwohl es
    nur 3 Monate sind"): `dauer_jahre` (Default 1.0 = Vollzeitraum, wie
    bisher) kommt von _scenario_dauer() und macht diese Tabelle bei
    Teil-Jahr-Szenarien korrekt (vorher wurden IMMER volle Jahres-Ausgaben
    gegen nur den simulierten Einnahmen-Zeitraum gerechnet -- bei einem
    3-Monats-Szenario waeren die Ausgaben also 4x zu hoch ausgewiesen worden).
    """
    einnahmen: dict = {}
    for key, label in EINNAHMEN_LABELS.items():
        if key in module:
            einnahmen[label] = einnahmen.get(label, 0.0) + module[key]
    total_einnahmen = sum(einnahmen.values())

    wacc_pct = kapitalkosten["wacc"] * 100.0
    # Label zeigt die WACC-Annahme direkt mit an (Beats Wunsch), z.B.
    # "Kapitalkosten (Annahme 3%)" -- damit ist im Report sofort klar, dass
    # der Zins-Anteil nicht willkuerlich ist, sondern auf einer expliziten
    # (aktuell Default-)Annahme beruht, die mit Beat noch geprueft werden kann.
    if wacc_pct == int(wacc_pct):
        wacc_label = f"{int(wacc_pct)}%"
    else:
        wacc_label = f"{wacc_pct:.1f}%"
    ausgaben = {
        "Amortisation": kapitalkosten["amortisation_jahr1"] * dauer_jahre,
        f"Kapitalkosten (Annahme {wacc_label})": kapitalkosten["zins_jahr1"] * dauer_jahre,
        "Unterhalt": kapitalkosten["unterhalt"] * dauer_jahre,
    }
    # NEU: Degradationskosten (siehe battery_optimization.
    # DEGRADATIONSKOSTEN_CHF_PRO_KWH) -- steckt schon IN der LP-Zielfunktion,
    # wird hier nur sichtbar gemacht, damit "Total Einnahmen - Total Ausgaben"
    # mit dem tatsaechlichen Optimierungsergebnis uebereinstimmt. `.get(...)`
    # mit Default 0.0, falls ein aelterer Aufrufer `kapitalkosten` ohne diesen
    # Schluessel uebergibt (z.B. alter capital_costs()-Aufruf ohne
    # entladeenergie_kwh) -- dann bleibt das Verhalten wie bisher.
    degradation = kapitalkosten.get("degradation", 0.0)
    if degradation:
        ausgaben["Batterie-Degradation"] = degradation
    total_ausgaben = sum(ausgaben.values())

    return {
        "einnahmen": einnahmen,
        "total_einnahmen": total_einnahmen,
        "ausgaben": ausgaben,
        "total_ausgaben": total_ausgaben,
        "gewinn_verlust": total_einnahmen - total_ausgaben,
    }


def build_rendite_kennzahlen(
    module: dict, kapitalkosten: dict, ts: pd.DataFrame, dt_hours: float = 0.25, dauer_jahre: float = 1.0
) -> dict:
    """Renditebetrachtung fuer den PDF-Report, Teil 2:
      - lcos_rp_kwh: Levelized Cost of Storage in Rp/kWh =
        (Annuitaet + Unterhalt) / jaehrliche Entladeenergie [kWh] * 100.
      - amortisationsdauer_jahre: capex / Netto-Cashflow (Jahre). Netto-
        Cashflow = Einnahmen (auf ein volles Jahr hochgerechnet, siehe unten)
        - Unterhalt. Amortisation/Kapitalkosten selbst werden NICHT
        abgezogen -- sonst waere die Frage "wie lange dauert die
        Amortisation" zirkulaer.
      - kapitalverzinsung_pct: derselbe Netto-Cashflow als Prozentsatz des
        Capex -- eine VEREINFACHTE durchschnittliche Rendite (keine echte
        IRR/NPV-Rechnung mit Wiederanlage/Restwert; falls gewuenscht, spaeter
        nachruestbar).

    NEU (Beats Frage "wieso sagt der Report eine Jahresbetrachtung, obwohl es
    nur 3 Monate sind"): diese drei Kennzahlen sind ihrer Natur nach JAHRES-/
    LEBENSZEIT-Groessen (LCOS = Kosten pro kWh ueber die Lebensdauer,
    Amortisationsdauer/Kapitalverzinsung vergleichen JAEHRLICHEN Cashflow
    gegen die einmalige Investition). Deckt das Szenario NICHT ein volles
    Jahr ab (dauer_jahre != 1.0, siehe _scenario_dauer()), werden die
    BEOBACHTETEN Einnahmen/Entladeenergie des simulierten Zeitraums auf ein
    volles Jahr HOCHGERECHNET (geteilt durch dauer_jahre), bevor diese drei
    Kennzahlen berechnet werden -- sonst waeren sie massiv verzerrt (z.B.
    Amortisationsdauer bei einem 3-Monats-Szenario ohne Hochrechnung um den
    Faktor 4 zu hoch, weil nur 3 Monate Einnahmen gegen die volle
    Jahres-Investition gerechnet wuerden). WICHTIG: das ist eine Hochrechnung/
    Annahme (die uebrigen Monate verhalten sich wie die simulierten) -- kein
    Ersatz fuer einen echten Volljahres-Lauf. Der PDF-Report weist bei
    dauer_jahre != 1 im Titel explizit auf den simulierten Zeitraum hin
    (siehe build_pdf_report()/_scenario_dauer()), damit diese Hochrechnung
    nicht mit einer echten Jahresmessung verwechselt wird.

    Bei Netto-Cashflow <= 0 sind Amortisationsdauer/Kapitalverzinsung nicht
    sinnvoll definierbar -- dann stehen dort `None` (der PDF-Report zeigt dann
    "nicht amortisierbar" statt einer Zahl).
    """
    # Einnahmen/Degradation des TATSAECHLICH simulierten Zeitraums (nicht die
    # bereits auf `dauer_jahre` SKALIERTEN Ausgaben aus
    # build_income_expense_summary() -- hier wird stattdessen in die andere
    # Richtung gerechnet: die IST-Einnahmen werden HOCHGERECHNET).
    einnahmen_ist = sum(module[key] for key in EINNAHMEN_LABELS if key in module)
    einnahmen_annualisiert = einnahmen_ist / dauer_jahre
    degradation_ist = kapitalkosten.get("degradation", 0.0)
    degradation_annualisiert = degradation_ist / dauer_jahre
    netto_cashflow = einnahmen_annualisiert - kapitalkosten["unterhalt"] - degradation_annualisiert

    entlade_energie_kwh_ist = float(ts["dis"].sum() * dt_hours)
    entlade_energie_kwh = entlade_energie_kwh_ist / dauer_jahre
    if entlade_energie_kwh > 1e-9:
        lcos_rp_kwh = (
            (kapitalkosten["annuitaet"] + kapitalkosten["unterhalt"] + degradation_annualisiert)
            / entlade_energie_kwh * 100.0
        )
    else:
        lcos_rp_kwh = None

    if netto_cashflow > 1e-9:
        amortisationsdauer_jahre = kapitalkosten["capex"] / netto_cashflow
        kapitalverzinsung_pct = netto_cashflow / kapitalkosten["capex"] * 100.0
    else:
        amortisationsdauer_jahre = None
        kapitalverzinsung_pct = None

    return {
        "lcos_rp_kwh": lcos_rp_kwh,
        "amortisationsdauer_jahre": amortisationsdauer_jahre,
        "kapitalverzinsung_pct": kapitalverzinsung_pct,
        "netto_cashflow": netto_cashflow,
        "entlade_energie_kwh": entlade_energie_kwh,  # annualisiert, siehe oben
    }


# --------------------------------------------------------------------------
# Monats-/Quartals-Aggregationen (fuer den PDF-Report, Teil 3)
# --------------------------------------------------------------------------

def build_monthly_energy_table(
    ts: pd.DataFrame, pv_profile: pd.Series, dt_hours: float = 0.25
) -> pd.DataFrame:
    """Monatliche Energie-Kennzahlen (PDF-Report Teil 3, erster Chart):
    Eigenverbrauch_kWh, Rueckspeisung_kWh, Netzbezug_kWh,
    Eigenverbrauchsquote_Prozent.

    Definitionen (Standard-PV-Begriffe):
      - Rueckspeisung = Export-Flow ins Netz (ts["export"]).
      - Eigenverbrauch = PV-Erzeugung - Rueckspeisung, d.h. die gesamte
        PV-Energie, die vor Ort bleibt -- egal ob direkt verbraucht oder ueber
        die Batterie zwischengespeichert.
      - Netzbezug = gesamter Netzbezug (ts["netzbezug_total"], inkl. der
        Batterieladung aus dem Netz) -- "wie viel wurde in diesem Monat
        effektiv vom Netz gekauft".
      - Eigenverbrauchsquote = Eigenverbrauch / PV-Erzeugung.

    Monatsgruppierung ueber den Kalendermonat (1-12) -- setzt wie im ganzen
    Projekt voraus, dass die Daten genau ein Kalenderjahr abdecken.
    """
    idx = ts.index
    monat = idx.month
    gueltige_monate = _valid_calendar_groups(monat)

    pv_kwh = pd.Series(pv_profile.values * dt_hours, index=idx).groupby(monat).sum()
    export_kwh = pd.Series(ts["export"].values * dt_hours, index=idx).groupby(monat).sum()
    netzbezug_kwh = pd.Series(ts["netzbezug_total"].values * dt_hours, index=idx).groupby(monat).sum()
    # NEU: Monate mit nur einer Handvoll Zeitschritten (DST-Randfragment,
    # siehe _valid_calendar_groups()) werden verworfen, bevor daraus Summen/
    # Quoten berechnet werden.
    pv_kwh = pv_kwh[pv_kwh.index.isin(gueltige_monate)]
    export_kwh = export_kwh[export_kwh.index.isin(gueltige_monate)]
    netzbezug_kwh = netzbezug_kwh[netzbezug_kwh.index.isin(gueltige_monate)]

    eigenverbrauch_kwh = pv_kwh - export_kwh
    # np.errstate: bei einem Bat_Last-System (kein PV, siehe Beats Wunsch
    # "reine Eigenverbrauchsoptimierung") ist pv_kwh ueberall 0 -- np.where
    # wertet trotz des 0.0-Fallbacks BEIDE Zweige aus, was sonst eine
    # harmlose, aber verwirrende "divide by zero"-Konsolenwarnung ausloest.
    with np.errstate(divide="ignore", invalid="ignore"):
        eigenverbrauchsquote_pct = np.where(
            pv_kwh.values > 1e-9, eigenverbrauch_kwh.values / pv_kwh.values * 100.0, 0.0
        )

    tabelle = pd.DataFrame(
        {
            "Eigenverbrauch_kWh": eigenverbrauch_kwh,
            "Rueckspeisung_kWh": export_kwh,
            "Netzbezug_kWh": netzbezug_kwh,
            "Eigenverbrauchsquote_Prozent": eigenverbrauchsquote_pct,
        }
    )
    tabelle.index.name = "Monat"
    return tabelle


def _build_quarterly_avg_profile_generic(idx: pd.DatetimeIndex, werte: np.ndarray, label: str = "") -> pd.DataFrame:
    """Kern-Logik hinter dem Quartals-Tagesprofil-Chart, verallgemeinert auf
    eine beliebige kW-Zeitreihe (statt nur Nettolast) -- Basis fuer die
    Ohne/Mit-Batterie-Vergleichscharts weiter unten (Eigenverbrauch,
    Netzbezug-Tagesprofil): fuer jede Viertelstunde des Tages wird der
    Durchschnitt ueber alle Tage des jeweiligen Quartals gebildet -> 4 Spalten
    (Q1..Q4) ueber eine 24h-Achse (Index = "HH:MM").

    `label` ist nur fuer die Diagnose-Ausgabe (siehe unten) -- damit im
    Konsolen-Log klar ist, welche Vergleichs-Zeitreihe (z.B. "Nettolast",
    "Netzbezug ohne Batterie") betroffen ist."""
    tag = f" [{label}]" if label else ""

    # Diagnose (Beat meldete: Q2/Q3/Q4 fehlten im PDF-Chart, mit synthetischen
    # Testdaten nicht reproduzierbar) -- Laengen-Check zuerst, da `werte` hier
    # POSITIONELL (nicht ueber den Index) mit `idx` kombiniert wird: bei
    # abweichender Laenge/Reihenfolge wuerden Werte den falschen Zeitstempeln
    # zugeordnet, ohne dass ein Fehler auftritt.
    if len(werte) != len(idx):
        print(
            f"  WARNUNG (Quartalsprofil{tag}): Laengen-Mismatch -- "
            f"idx={len(idx)}, werte={len(werte)}. Werte koennten falsch "
            f"zugeordnet sein."
        )

    quartal = idx.quarter
    zeit_im_tag = idx.strftime("%H:%M")  # zero-padded -> alphabetisch = chronologisch sortierbar

    # NEU (Beats Frage "wieso hat der Report Q3-Werte, obwohl das Ganze nur
    # bis 30.6. 23:45 geht"): ein Quartal mit nur einer Handvoll Zeitschritten
    # -- typischerweise ein DST-Randfragment durch die Sommerzeit-Umstellung,
    # siehe _valid_calendar_groups() -- wird verworfen, BEVOR der Mittelwert
    # pro Uhrzeit-Bin berechnet wird (sonst waere der "Tagesdurchschnitt"
    # dieses Quartals aus nur 1-4 Werten statt aus Tausenden berechnet, was
    # zu einem irrefuehrenden Ausschlag im Chart fuehrt).
    gueltige_quartale = _valid_calendar_groups(quartal)
    verworfene_quartale = sorted(set(quartal) - gueltige_quartale)
    if verworfene_quartale:
        n_verworfen = int(pd.Series(quartal).isin(verworfene_quartale).sum())
        print(
            f"  HINWEIS (Quartalsprofil{tag}): Quartal(e) "
            f"{[f'Q{q}' for q in verworfene_quartale]} mit nur {n_verworfen} "
            "Zeitschritt(en) verworfen (DST-Randfragment durch die Sommerzeit-"
            "Umstellung -- siehe _valid_calendar_groups())."
        )
        maske = pd.Series(quartal).isin(gueltige_quartale).values
        quartal = quartal[maske]
        zeit_im_tag = zeit_im_tag[maske]
        werte = np.asarray(werte)[maske]

    df = pd.DataFrame({"wert": werte, "quartal": quartal, "zeit": zeit_im_tag})
    pivot = df.pivot_table(index="zeit", columns="quartal", values="wert", aggfunc="mean")
    pivot = pivot.sort_index()
    pivot.columns = [f"Q{q}" for q in pivot.columns]
    pivot.index.name = "Uhrzeit"

    # Diagnose-Ausgabe: Zeitschritte je Quartal + NaN-Check je Spalte. Falls
    # eine Quartals-Spalte KOMPLETT NaN ist, zeichnet matplotlib nichts (die
    # Legende zeigt die Linie aber trotzdem an, siehe ax.plot(..., label=col)
    # in build_pdf_report) -- das erklaert genau Beats urspruengliches Symptom
    # "Legende zeigt Q1-Q4, aber nur Q1 ist sichtbar".
    quartal_counts = pd.Series(quartal).value_counts().sort_index()
    print(f"  Diagnose Quartalsprofil{tag} -- Zeitschritte je Quartal:",
          {f"Q{q}": int(c) for q, c in quartal_counts.items()})
    for col in pivot.columns:
        n_nan = int(pivot[col].isna().sum())
        if n_nan > 0:
            print(
                f"  WARNUNG (Quartalsprofil{tag}): Spalte {col} hat "
                f"{n_nan} von {len(pivot)} NaN-Werten"
                + (" (KOMPLETT NaN -> Linie im PDF unsichtbar!)" if n_nan == len(pivot) else "") + "."
            )
    if idx.duplicated().any():
        print(
            f"  WARNUNG (Quartalsprofil{tag}): idx enthaelt "
            f"{int(idx.duplicated().sum())} doppelte Zeitstempel."
        )

    return pivot


def build_quarterly_avg_profile(ts: pd.DataFrame, pv_profile: pd.Series, last_profile: pd.Series) -> pd.DataFrame:
    """Durchschnittliches Tagesprofil des Nettolastprofils (Last - PV, kW) je
    Quartal: fuer jede Viertelstunde des Tages wird der Durchschnitt ueber
    alle Tage des jeweiligen Quartals gebildet -> 4 Spalten (Q1..Q4) ueber
    eine 24h-Achse (Index = "HH:MM"). Zeigt die typische saisonale Tagesform
    des Nettolastprofils (nutzt intern denselben Helper wie die
    Ohne/Mit-Batterie-Netzbezug-Charts, siehe _build_quarterly_avg_profile_generic()).
    Nicht mehr direkt im PDF-Report verwendet (siehe
    build_netzbezug_quarterly_profiles() fuer den Ohne/Mit-Vergleich auf
    Beats Wunsch), bleibt aber als eigenstaendige Funktion/Kennzahl erhalten."""
    netto = last_profile.values - pv_profile.values
    return _build_quarterly_avg_profile_generic(ts.index, netto, label="Nettolast")


# --------------------------------------------------------------------------
# Referenzfall OHNE Batterie (fuer die Ohne/Mit-Batterie-Vergleichscharts im
# PDF-Report, auf Beats Wunsch: "Kosten ... vergleichen zu den 'initialen'
# Kosten, also vor der Batterie"). Rein aus pv_profile/last_profile ableitbar
# -- kein battery_optimization/oemof-Zugriff noetig.
# --------------------------------------------------------------------------

def build_baseline_energy_series(pv_profile: pd.Series, last_profile: pd.Series):
    """Referenzfall OHNE Batterie, pro Zeitschritt (kW): ohne Speicher gibt es
    nur DIREKTEN PV-Eigenverbrauch (keine Zwischenspeicherung fuer spaeteren
    Verbrauch):
      - Eigenverbrauch_ohne(t) = min(PV(t), Last(t))
      - Rueckspeisung_ohne(t)  = PV(t) - Eigenverbrauch_ohne(t)   (>= 0)
      - Netzbezug_ohne(t)      = Last(t) - Eigenverbrauch_ohne(t) (>= 0)
    Gibt die drei kW-Arrays zurueck (gleiche Laenge/Reihenfolge wie
    pv_profile/last_profile)."""
    pv = pv_profile.values
    last = last_profile.values
    eigenverbrauch_ohne = np.minimum(pv, last)
    rueckspeisung_ohne = pv - eigenverbrauch_ohne
    netzbezug_ohne = last - eigenverbrauch_ohne
    return eigenverbrauch_ohne, rueckspeisung_ohne, netzbezug_ohne


def build_baseline_monthly_energy_table(
    pv_profile: pd.Series, last_profile: pd.Series, dt_hours: float = 0.25
) -> pd.DataFrame:
    """Wie build_monthly_energy_table(), aber fuer den Referenzfall OHNE
    Batterie -- fuer das "Ohne Batterie"-Panel neben dem bestehenden
    Eigenverbrauch/Rueckspeisung-Chart."""
    idx = pv_profile.index
    monat = idx.month
    gueltige_monate = _valid_calendar_groups(monat)
    _, rueckspeisung_ohne, netzbezug_ohne = build_baseline_energy_series(pv_profile, last_profile)

    pv_kwh = pd.Series(pv_profile.values * dt_hours, index=idx).groupby(monat).sum()
    rueckspeisung_kwh = pd.Series(rueckspeisung_ohne * dt_hours, index=idx).groupby(monat).sum()
    netzbezug_kwh = pd.Series(netzbezug_ohne * dt_hours, index=idx).groupby(monat).sum()
    # NEU: DST-Randfragment-Monate verwerfen, siehe _valid_calendar_groups().
    pv_kwh = pv_kwh[pv_kwh.index.isin(gueltige_monate)]
    rueckspeisung_kwh = rueckspeisung_kwh[rueckspeisung_kwh.index.isin(gueltige_monate)]
    netzbezug_kwh = netzbezug_kwh[netzbezug_kwh.index.isin(gueltige_monate)]
    eigenverbrauch_kwh = pv_kwh - rueckspeisung_kwh
    with np.errstate(divide="ignore", invalid="ignore"):
        eigenverbrauchsquote_pct = np.where(
            pv_kwh.values > 1e-9, eigenverbrauch_kwh.values / pv_kwh.values * 100.0, 0.0
        )

    tabelle = pd.DataFrame(
        {
            "Eigenverbrauch_kWh": eigenverbrauch_kwh,
            "Rueckspeisung_kWh": rueckspeisung_kwh,
            "Netzbezug_kWh": netzbezug_kwh,
            "Eigenverbrauchsquote_Prozent": eigenverbrauchsquote_pct,
        }
    )
    tabelle.index.name = "Monat"
    return tabelle


def build_baseline_monthly_peak_table(pv_profile: pd.Series, last_profile: pd.Series) -> pd.Series:
    """Wie build_monthly_peak_table(), aber fuer den Referenzfall OHNE
    Batterie (dieselbe Groesse, auf der auch die Peak-Shaving-Ersparnis in
    battery_optimization.main() basiert)."""
    idx = pv_profile.index
    monat = idx.month
    gueltige_monate = _valid_calendar_groups(monat)
    _, _, netzbezug_ohne = build_baseline_energy_series(pv_profile, last_profile)
    peak = pd.Series(netzbezug_ohne, index=idx).groupby(monat).max()
    peak = peak[peak.index.isin(gueltige_monate)]  # DST-Randfragment verwerfen
    peak.index.name = "Monat"
    peak.name = "Lastspitze_kW"
    return peak


def build_netzbezug_quarterly_profiles(
    ts: pd.DataFrame, pv_profile: pd.Series, last_profile: pd.Series
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Durchschnittliches Tagesprofil des NETZBEZUGS (kW, >= 0) je Quartal --
    einmal fuer den Referenzfall OHNE Batterie (max(Last-PV, 0)) und einmal
    MIT Batterie (ts["netzbezug_total"], das tatsaechliche Optimierungs-
    ergebnis). Basis fuer den Ohne/Mit-Vergleichschart "Bezug Energie ab Netz
    -- Tagesprofil je Quartal" im PDF-Report (ersetzt den vorherigen reinen
    Nettolast-Chart, siehe build_quarterly_avg_profile())."""
    _, _, netzbezug_ohne = build_baseline_energy_series(pv_profile, last_profile)
    profil_ohne = _build_quarterly_avg_profile_generic(
        pv_profile.index, netzbezug_ohne, label="Netzbezug ohne Batterie"
    )
    profil_mit = _build_quarterly_avg_profile_generic(
        ts.index, ts["netzbezug_total"].values, label="Netzbezug mit Batterie"
    )
    return profil_ohne, profil_mit


def build_einspeisung_quarterly_profiles(
    ts: pd.DataFrame, pv_profile: pd.Series, last_profile: pd.Series
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Wie build_netzbezug_quarterly_profiles(), aber fuer die NETZEINSPEISUNG
    (Rueckspeisung/Export ins Netz, kW, >= 0) statt des Netzbezugs -- Beats
    Wunsch: "kannst du auf Seite 2 noch den Quartalsplot für Netzeinspeisung
    hinzufügen, können ja die gleichen Farben sein jedoch gestrichelte
    Linien" (im selben Chart wie Netzbezug ueberlagert, siehe
    _plot_netzbezug() in build_pdf_report()).

    Referenzfall OHNE Batterie: direkter PV-Ueberschuss (PV - Eigenverbrauch,
    siehe build_baseline_energy_series()). MIT Batterie: ts["export"], der
    tatsaechliche Optimierungs-Export (kann auch groesser sein als der reine
    PV-Ueberschuss, falls die Batterie zeitweise ins Netz entlaedt)."""
    _, rueckspeisung_ohne, _ = build_baseline_energy_series(pv_profile, last_profile)
    profil_ohne = _build_quarterly_avg_profile_generic(
        pv_profile.index, rueckspeisung_ohne, label="Einspeisung ohne Batterie"
    )
    profil_mit = _build_quarterly_avg_profile_generic(
        ts.index, ts["export"].values, label="Einspeisung mit Batterie"
    )
    return profil_ohne, profil_mit


def build_monthly_charge_discharge_table(ts: pd.DataFrame, dt_hours: float = 0.25) -> pd.DataFrame:
    """Monatliche Lade-/Entladeenergie (kWh) -- Ersatz fuer den Eigenverbrauch/
    Rueckspeisung-Chart bei einem REINEN Batterie-System (kein PV, keine Last,
    siehe `nur_batterie`-Weiche in build_pdf_report()). In diesem Fall sind
    "Eigenverbrauch"/"Rueckspeisung" begrifflich nicht sinnvoll (es gibt weder
    PV-Erzeugung noch Verbrauch, dessen Deckung man "Eigenverbrauch" nennen
    koennte) -- der Chart zeigt statt dessen schlicht, wie viel die Batterie
    pro Monat ge- bzw. entladen hat (unabhaengig von PV-/Netz-Quelle, d.h.
    ch_pv+ch_grid zusammen als "Laden")."""
    idx = ts.index
    monat = idx.month
    gueltige_monate = _valid_calendar_groups(monat)
    laden_kwh = pd.Series(
        (ts["ch_pv"].values + ts["ch_grid"].values) * dt_hours, index=idx
    ).groupby(monat).sum()
    entladen_kwh = pd.Series(ts["dis"].values * dt_hours, index=idx).groupby(monat).sum()
    # NEU: DST-Randfragment-Monate verwerfen, siehe _valid_calendar_groups()
    # (Beats Frage "wieso hat der Report Q3-Werte" -- dieselbe Ursache zeigte
    # sich hier als ein fast leerer "Jul"-Balken).
    laden_kwh = laden_kwh[laden_kwh.index.isin(gueltige_monate)]
    entladen_kwh = entladen_kwh[entladen_kwh.index.isin(gueltige_monate)]
    tabelle = pd.DataFrame({"Laden_kWh": laden_kwh, "Entladen_kWh": entladen_kwh})
    tabelle.index.name = "Monat"
    return tabelle


def build_monthly_peak_table(ts: pd.DataFrame) -> pd.Series:
    """Monatliche Lastspitze (Maximum von Netzbezug_Total, kW; PDF-Report Teil
    3, dritter Chart) -- dieselbe Groesse, auf der auch die Peak-Shaving-
    Kosten basieren (Parameter!C9)."""
    monat = ts.index.month
    gueltige_monate = _valid_calendar_groups(monat)
    peak = ts["netzbezug_total"].groupby(monat).max()
    peak = peak[peak.index.isin(gueltige_monate)]  # DST-Randfragment verwerfen
    peak.index.name = "Monat"
    peak.name = "Lastspitze_kW"
    return peak


def save_output_excel(
    output_path: str,
    module: dict,
    ts: pd.DataFrame,
    anteile: pd.DataFrame,
    zeitreihen: pd.DataFrame,
    params: dict,
    pv_profile: pd.Series,
    last_profile: pd.Series,
    dt_hours: float = 0.25,
):
    """Schreibt die Output-Kontrolle als Excel mit zwei Sheets:
    "Zeitreihen" (eine Zeile pro 15-Min-Intervall, siehe build_output_table())
    und "Modul-Ertraege" (Jahres-Zusammenfassung, siehe
    build_module_summary_table()). Wird von run_battery_analysis.py
    aufgerufen; gibt beide Tabellen zusaetzlich zurueck, falls man sie direkt
    weiterverwenden will (z.B. spaeter fuer build_pdf_report()).
    """
    # Spalten-Label fuer "Modul-Ertraege" (NEU, siehe build_module_summary_
    # table()-Docstring): vor der Zuercher-Lokalzeit-Umrechnung (unten)
    # ermitteln, aus demselben Grund wie in build_pdf_report() (DST-Versatz
    # wuerde das Enddatum sonst verfaelschen) -- hier reicht die reine
    # Dauer, kein Titel noetig.
    _dauer_jahre_excel, _ = _scenario_dauer(ts.index[0], len(ts.index), dt_hours)
    if abs(_dauer_jahre_excel * 365.25 - 365.25) <= 3 or abs(_dauer_jahre_excel * 365.25 - 366) <= 3:
        _modul_spalten_label = "CHF_pro_Jahr"
    else:
        _modul_spalten_label = f"CHF_pro_Zeitraum_{_dauer_jahre_excel * 365.25:.0f}_Tage"

    # NEU (Beats Wunsch): Excel-Zeitspalte in ECHTER Zuercher Ortszeit MIT
    # Sommerzeit-Umstellung zeigen, statt der intern verwendeten festen
    # UTC+1-"CET" ohne DST -- siehe _to_zurich_local().
    ts, anteile, zeitreihen, pv_profile, last_profile = _to_zurich_local(
        ts, anteile, zeitreihen, pv_profile, last_profile
    )
    zeitreihen_tabelle = build_output_table(
        ts, anteile, zeitreihen, params, pv_profile, last_profile, dt_hours
    )
    modul_tabelle = build_module_summary_table(module, spalten_label=_modul_spalten_label)

    with pd.ExcelWriter(output_path) as writer:
        zeitreihen_tabelle.to_excel(writer, sheet_name="Zeitreihen")
        modul_tabelle.to_excel(writer, sheet_name="Modul-Ertraege")

    return zeitreihen_tabelle, modul_tabelle


MONATSNAMEN = ["Jan", "Feb", "Mar", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"]

# --------------------------------------------------------------------------
# Fleco-Power-Corporate-Design fuer den PDF-Report. Auf Beats Wunsch an sein
# eigenes Fleco-Dokument ("20260127 Ertragspotenzial alle Produkte.docx")
# angenaehert: helles, klares Letterhead-Layout mit Fleco-Logo/-Gruen statt
# des vorherigen dunklen Dashboard-Looks. Farben 1:1 aus dem Word-Theme
# dieses Referenzdokuments uebernommen (word/theme/theme1.xml: accent3 =
# Dunkelgruen [Ueberschriften/dunkle Chart-Linie], accent6 = Hellgruen
# [helle Chart-Linie], Grautoene aus den Tabellen/Beschriftungen). "Ohne
# Batterie" wird bewusst neutral-grau dargestellt (Referenzzustand), "Mit
# Batterie" in Fleco-Gruen (das eigentliche Ergebnis/Produkt) -- Fleco selbst
# hat keine zweite Akzentfarbe fuer diesen Zweck, Grau/Gruen ist eine
# uebliche und dezente Konvention fuer "Vorher/Nachher".
#
# EINSCHRAENKUNG: Die Fleco-Hausschrift laut Referenzdokument ist "Museo Sans
# 500" -- liess sich in dieser Sandbox nicht installieren (Font-Registry-
# Zugriff aktuell blockiert), daher DejaVu Sans (matplotlib-Standard) als
# naechstliegender Ersatz.
_BG = "#FFFFFF"
_TEXT = "#1A1A1A"
_TEXT_DIM = "#595959"
_LINE = "#BFBFBF"
_GRID = "#E8E8E8"
_ACCENT = "#196B24"       # Fleco Dunkelgruen (Headings, "Mit Batterie")
_ACCENT2 = "#4EA72E"      # Fleco Hellgruen (Sekundaerserie "Mit Batterie")
_GRAU = "#8C8C8C"         # Referenzzustand "Ohne Batterie"
_GRAU2 = "#C7C7C7"        # Sekundaerserie "Ohne Batterie"
_POS = _ACCENT
_NEG = "#B23B3B"
# Kategorial-Palette fuer die Quartals-Linien (Q1..Q4).
# NEU (Beat: "kannst du die Farben von Q1 und Q2 im Netzbezug Tagesprofil
# noch deutlicher unterscheiden?"): das alte Q2-Blau "#156082" hatte zu
# wenig Farbsaettigung (Chroma) -- dadurch wirkte es im Chart fast so
# gedeckt/dunkel wie das Q1-Gruen "#196B24" und die beiden Linien waren auf
# den ersten Blick schwer zu unterscheiden, obwohl der Farbton (Gruen vs.
# Blau) technisch verschieden war. Ersetzt durch ein kraeftigeres,
# CVD-sicheres Blau (Okabe-Ito "#0072B2") mit deutlich hoeherer Chroma --
# mit dem `dataviz`-Skill-Validator geprueft: normal-vision ΔE Q1<->Q2 jetzt
# ~15.9 (vorher, mit dem alten Blau, war das schwaechste Paar im Vergleich
# insgesamt nur knapp ueber der Schwelle) und CVD-Trennung (deutan/tritan)
# ueber allen Paaren im gruenen Bereich -- keine der anderen drei Farben
# (Q1/Q3/Q4) wurde angetastet.
_QCOLORS = ["#196B24", "#0072B2", "#4EA72E", "#8C8C8C"]

# Fleco-Logo (aus Beats Referenzdokument extrahiert) -- muss im selben Ordner
# wie dieses Skript liegen. Fehlt die Datei (z.B. weil nur die .py-Dateien
# kopiert wurden), wird der Report einfach ohne Logo gebaut statt abzubrechen
# (siehe _load_fleco_logo()).
_LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fleco_logo.png")


def _load_fleco_logo():
    try:
        import matplotlib.image as mpimg
        return mpimg.imread(_LOGO_PATH)
    except Exception as exc:
        print(f"  HINWEIS: Fleco-Logo nicht gefunden/lesbar ({_LOGO_PATH}): {exc} "
              f"-- Report wird ohne Logo erstellt.")
        return None


def _pdf_style_axes(ax):
    """Gemeinsames Achsen-Styling im hellen Fleco-Look: dezente, helle
    Gridlines, keine Umrandung oben/rechts (siehe dataviz-Skill: recessive
    grid/axes, thin marks)."""
    ax.set_facecolor(_BG)
    ax.grid(axis="y", color=_GRID, linewidth=0.8, zorder=0)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color(_LINE)
    ax.spines["bottom"].set_color(_LINE)
    ax.tick_params(colors=_TEXT_DIM)


def _add_logo(fig, ax, logo_img, x_in, y_in, size_in):
    """Platziert das Fleco-Logo an einer festen Position (in Zoll, gemessen
    von unten-links) mit fester Hoehe `size_in` -- Breite wird aus dem
    Bild-Seitenverhaeltnis abgeleitet, damit das Logo nicht verzerrt wird."""
    if logo_img is None:
        return
    fig_w, fig_h = fig.get_size_inches()
    h_px, w_px = logo_img.shape[0], logo_img.shape[1]
    aspect = w_px / h_px
    width_in = size_in * aspect
    x0, y0 = x_in / fig_w, y_in / fig_h
    # WICHTIG: aspect="auto" ist hier zwingend -- imshow() setzt sonst per
    # Default (rcParams["image.aspect"]="equal") ax.set_aspect("equal") auf
    # der GESAMTEN Overlay-Axes. Bei einer nicht-quadratischen Figure (z.B.
    # A4 8.27x11.69in) schrumpft/zentriert matplotlib dann die effektive
    # Zeichenflaeche der Axes, um das Seitenverhaeltnis einzuhalten -- das
    # verschiebt ALLE anderen per transAxes platzierten Elemente (Titel,
    # Tabellen, Fusszeile) sichtbar von den Seitenraendern weg nach innen.
    # Erst durch aspect="auto" bleibt die Axes-Box exakt bei [0,0,1,1].
    ax.imshow(logo_img, extent=(x0, x0 + width_in / fig_w, y0, y0 + size_in / fig_h),
              zorder=10, interpolation="bilinear", aspect="auto")


def _fleco_page(fig_size, page_title, logo_img, page_no=None):
    """Baut eine neue Seite im Fleco-Letterhead-Stil: Logo oben, darunter ein
    durchgehender gruener Titel-Banner (voller Breite, weisser fetter Text),
    Kontakt-Fusszeile unten (Adresse aus Beats Referenzdokument). Gibt (fig,
    ax, fig_w, fig_h, content_top_in, content_bottom_in) zurueck -- ax ist
    eine Vollbild-Overlay-Axes mit Koordinaten 0..1 (transAxes),
    content_top_in/content_bottom_in sind die fuer den eigentlichen
    Seiteninhalt verfuegbaren Grenzen IN ZOLL (von unten gemessen, wie alle
    anderen Positionsangaben in dieser Funktion).

    NEU (Beats Wunsch "Report formeller machen, ein wenig im Style von
    [Referenzdokument]"): das Referenzdokument nutzt auf JEDER Seite einen
    vollen gruenen Banner als Titelzeile statt einer duennen Trennlinie unter
    dem Titel -- hier uebernommen (durchgaengiges Redesign, betrifft ALLE
    Seiten, da alle ueber diese eine Funktion laufen). Das Logo (farbiges
    Wordmark-PNG) bleibt bewusst OBERHALB des Banners auf weissem Grund --
    direkt AUF dem gruenen Banner waere es kaum lesbar (kein Weiss-
    ausgespartes Logo verfuegbar). `page_no` (optional, z.B. "0", "1", "2"):
    rechts in der Fusszeile als "Seite N" angezeigt, analog zur "1/4"-Seiten-
    zahl im Referenzdokument -- Beats Wunsch, dass die bisherige
    Eingabeparameter-Seite als "Seite 0" gefuehrt wird (die neue System-
    Zusammenfassung wird dadurch "Seite 1", usw.).
    """
    import matplotlib.pyplot as plt
    from matplotlib.patches import Rectangle
    fig_w, fig_h = fig_size
    fig = plt.figure(figsize=fig_size)
    fig.patch.set_facecolor(_BG)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.axis("off")
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)

    margin_in = 0.5
    logo_h_in = 0.34
    top_in = fig_h - margin_in
    _add_logo(fig, ax, logo_img, x_in=margin_in, y_in=top_in - logo_h_in, size_in=logo_h_in)

    banner_top_in = top_in - logo_h_in - 0.14
    banner_h_in = 0.42
    banner_bottom_in = banner_top_in - banner_h_in
    ax.add_patch(Rectangle(
        (margin_in / fig_w, banner_bottom_in / fig_h),
        (fig_w - 2 * margin_in) / fig_w, banner_h_in / fig_h,
        transform=ax.transAxes, facecolor=_ACCENT, edgecolor="none", zorder=4, clip_on=False,
    ))
    ax.text((margin_in + 0.16) / fig_w, (banner_bottom_in + banner_h_in / 2) / fig_h, page_title,
            transform=ax.transAxes, fontsize=15, fontweight="bold", color="#FFFFFF",
            va="center", ha="left", zorder=5)

    footer_rule_in = 0.46
    ax.plot([margin_in / fig_w, 1 - margin_in / fig_w], [footer_rule_in / fig_h] * 2,
            color=_LINE, linewidth=0.8, transform=ax.transAxes, zorder=5)
    ax.text(margin_in / fig_w, (footer_rule_in - 0.17) / fig_h,
            "Fleco Power AG  |  Technoparkstrasse 2  |  8406 Winterthur  |  052 209 04 04  |  "
            "info@flecopower.ch  |  www.flecopower.ch",
            transform=ax.transAxes, fontsize=6.8, color=_TEXT_DIM, va="center", ha="left")
    if page_no is not None:
        ax.text(1 - margin_in / fig_w, (footer_rule_in - 0.17) / fig_h, f"Seite {page_no}",
                transform=ax.transAxes, fontsize=6.8, color=_TEXT_DIM, va="center", ha="right")

    content_top_in = banner_bottom_in - 0.22
    content_bottom_in = footer_rule_in + 0.12
    return fig, ax, fig_w, fig_h, content_top_in, content_bottom_in


def _draw_bordered_table(ax, fig_w, fig_h, x_in, top_in, width_in, rows, row_h_in=0.32):
    """Zeichnet eine einfache Tabelle im Fleco-Stil (duenner grauer Aussen-
    rahmen + duenne Trennlinien zwischen den Zeilen, mirrors die schlichten
    "Table Grid"-Tabellen in Beats Referenzdokument -- KEINE farbige
    Kopfzeile, nur fetter/gruener Text fuer Zwischenueberschriften und
    Totalzeilen). `rows`: Liste von Dicts mit "label", optional "value",
    "style" ("header"/"normal"/"total"/"result") und optional "color".
    Gibt die neue obere Position (in Zoll) NACH der Tabelle zurueck."""
    from matplotlib.patches import Rectangle
    height_in = len(rows) * row_h_in
    bottom_in = top_in - height_in
    ax.add_patch(Rectangle(
        (x_in / fig_w, bottom_in / fig_h), width_in / fig_w, height_in / fig_h,
        transform=ax.transAxes, facecolor="none", edgecolor=_LINE, linewidth=1.0,
        zorder=2, clip_on=False,
    ))
    pad_in = 0.14
    y_in = top_in
    for i, row in enumerate(rows):
        y_in -= row_h_in
        if i > 0:
            ax.plot([x_in / fig_w, (x_in + width_in) / fig_w], [(y_in + row_h_in) / fig_h] * 2,
                    color=_LINE, linewidth=0.6, transform=ax.transAxes, zorder=2)
        style = row.get("style", "normal")
        weight = "bold" if style in ("header", "total", "result") else "normal"
        size = 13 if style == "result" else 10.5
        color = row.get("color") or (_ACCENT if style == "header" else _TEXT)
        ax.text((x_in + pad_in) / fig_w, (y_in + row_h_in / 2) / fig_h, row["label"],
                transform=ax.transAxes, fontsize=size, fontweight=weight, color=color,
                va="center", ha="left")
        if row.get("value") is not None:
            ax.text((x_in + width_in - pad_in) / fig_w, (y_in + row_h_in / 2) / fig_h, row["value"],
                    transform=ax.transAxes, fontsize=size, fontweight=weight, color=color,
                    va="center", ha="right", family="monospace")
    return bottom_in


def _section_label(ax, fig_w, fig_h, x_in, y_in, text, size=13):
    """Abschnitts-Ueberschrift im Fleco-Stil (fett, dunkelgruen -- mirrors
    die "Heading 2"-Formatvorlage in Beats Referenzdokument: Farbe #196B24,
    fett, keine Dekoration)."""
    ax.text(x_in / fig_w, y_in / fig_h, text, transform=ax.transAxes, fontsize=size,
            fontweight="bold", color=_ACCENT, va="center", ha="left")


def _prose_box(ax, fig_w, fig_h, x_in, top_in, width_in, text, fontsize=9.8,
                line_h_in=0.20, pad_in=0.16):
    """Hellgraue Info-Box mit Fliesstext (NEU, Beats Wunsch "Report
    formeller machen ... im Style von [Referenzdokument]"): mirrors die
    hellgrauen Textboxen im Referenzdokument (z.B. "Ziele der durchgeführten
    Analysen"). Manueller Zeilenumbruch per `textwrap` (wie schon bei
    `_kpi_tile()` -- matplotlibs eingebautes `wrap=True` haelt sich nicht
    zuverlaessig an eine feste Breite, siehe dortiger Kommentar), mit einer
    an `width_in` orientierten Zeichenbreite. Gibt die neue obere Position
    (in Zoll) NACH der Box zurueck, fuer nachfolgenden Inhalt."""
    import textwrap
    from matplotlib.patches import Rectangle
    # Empirischer Zeichen-zu-Zoll-Faktor fuer die DejaVu-Sans-Standardschrift
    # bei `fontsize` -- grosszuegig genug bemessen, um ein Ueberlaufen der
    # Box (wie beim urspruenglichen KPI-Kachel-Bug) zu vermeiden.
    zeichen_pro_zoll = 12.5 * (9.8 / fontsize)
    breite_zeichen = max(20, int((width_in - 2 * pad_in) * zeichen_pro_zoll))
    zeilen = []
    for absatz in text.split("\n\n"):
        zeilen += textwrap.wrap(absatz, width=breite_zeichen) or [""]
        zeilen.append("")  # Leerzeile zwischen Absaetzen
    if zeilen and zeilen[-1] == "":
        zeilen.pop()
    height_in = len(zeilen) * line_h_in + 2 * pad_in
    bottom_in = top_in - height_in
    ax.add_patch(Rectangle(
        (x_in / fig_w, bottom_in / fig_h), width_in / fig_w, height_in / fig_h,
        transform=ax.transAxes, facecolor="#F2F2F0", edgecolor="none", zorder=2, clip_on=False,
    ))
    y_in = top_in - pad_in - line_h_in * 0.75
    for zeile in zeilen:
        ax.text((x_in + pad_in) / fig_w, y_in / fig_h, zeile,
                transform=ax.transAxes, fontsize=fontsize, color=_TEXT, va="center", ha="left")
        y_in -= line_h_in
    return bottom_in - 0.18


# NEU (Beats Rueckmeldung zu Seite 2: "bei den Modulen der Tabelle
# Beschreibung ein Modul pro Zeile und das Modul 'fett' machen, nicht
# groesser"): Modulname + Kurzbeschreibung als Datentupel, damit
# _modul_glossar_box() jedes Modul auf einer eigenen Zeile mit fett
# gesetztem Namen zeichnen kann (statt eines einzigen durchlaufenden
# Fliesstext-Absatzes wie zuvor).
_MODUL_GLOSSAR = [
    ("Eigenverbrauchsoptimierung", "Batterie deckt Last statt Netzbezug."),
    ("Einspeiseoptimierung", "PV-Energie wird zeitversetzt eingespeist (bei tiefem "
                             "Rückliefertarif geladen, bei höherem exportiert)."),
    ("Arbitrage", "Netzenergie wird günstig geladen und teurer exportiert."),
    ("SRL", "separat vermarktete Sekundärregelleistung."),
    ("Peak-Shaving", "reduziert die monatliche Bezugsspitze."),
]


def _modul_glossar_box(fig, ax, fig_w, fig_h, x_in, top_in, width_in, module_liste,
                        fontsize=7.6, line_h_in=None):
    """Zeichnet je Eintrag aus `module_liste` (Modulname, Beschreibung) EINE
    eigene Zeile "**Name** = Beschreibung" -- der Name fett (`fontweight=
    "bold"`), bei GLEICHER Schriftgroesse wie die Beschreibung (Beats
    Wunsch: "fett machen, nicht groesser"). Eine lange Beschreibung wird
    (wie bei `_prose_box()`) manuell per `textwrap` umgebrochen; Folgezeilen
    stehen ohne Einzug unter dem Namen. Da matplotlib keine gemischte Fett-/
    Normal-Formatierung INNERHALB eines einzigen `ax.text()`-Aufrufs
    unterstuetzt, wird der Name als eigener Text-Artist gezeichnet und der
    Rest der ersten Zeile direkt danach positioniert.

    NEU (Fix waehrend des Testrenderns gefunden): die x-Distanz dafuer NICHT
    ueber den pauschalen Zeichen-zu-Zoll-Faktor annaehern (der ist auf
    NORMALGEWICHTIGEN Text fuer den Zeilenumbruch kalibriert) -- FETTE
    Glyphen sind spuerbar breiter, das fuehrte bei laengeren Namen
    ("Eigenverbrauchsoptimierung") wie auch bei kurzen ("SRL") zu einer
    unterschaetzten Namensbreite, wodurch der Beschreibungstext (inkl. dem
    fuehrenden Leerzeichen vor "=") in den fett gezeichneten Namen
    hineinlief und wie "SRL= separat..." statt "SRL = separat..." aussah.
    Stattdessen wird die TATSAECHLICH gerenderte Breite des Namens per
    `Text.get_window_extent()` gemessen (Pixel -> Zoll ueber `fig.dpi`) und
    der Rest der Zeile exakt daran anschliessend platziert."""
    import textwrap
    line_h_in = line_h_in or (fontsize * 1.5 / 72.0)
    zeichen_pro_zoll = 12.5 * (9.8 / fontsize)
    breite_zeichen = max(20, int(width_in * zeichen_pro_zoll))
    y_in = top_in
    renderer = fig.canvas.get_renderer()
    for name, beschreibung in module_liste:
        voller_text = f"{name} = {beschreibung}"
        zeilen = textwrap.wrap(voller_text, width=breite_zeichen) or [""]
        erste_zeile = zeilen[0]
        if erste_zeile.startswith(name):
            t_name = ax.text(x_in / fig_w, y_in / fig_h, name, transform=ax.transAxes, fontsize=fontsize,
                              fontweight="bold", color=_TEXT_DIM, va="top", ha="left")
            name_w_in = t_name.get_window_extent(renderer=renderer).width / fig.dpi
            ax.text((x_in + name_w_in) / fig_w, y_in / fig_h, erste_zeile[len(name):],
                    transform=ax.transAxes, fontsize=fontsize, color=_TEXT_DIM, va="top", ha="left")
        else:
            ax.text(x_in / fig_w, y_in / fig_h, erste_zeile, transform=ax.transAxes,
                    fontsize=fontsize, color=_TEXT_DIM, va="top", ha="left")
        y_in -= line_h_in
        for folgezeile in zeilen[1:]:
            ax.text(x_in / fig_w, y_in / fig_h, folgezeile, transform=ax.transAxes,
                    fontsize=fontsize, color=_TEXT_DIM, va="top", ha="left")
            y_in -= line_h_in
    return y_in


def _kpi_tile(ax, fig_w, fig_h, x_in, top_in, width_in, height_in, label, value):
    """Eine einzelne Kennzahlen-Kachel (helle Box, grauer Rahmen, Label oben
    dezent, grosse gruene Zahl darunter) -- fuer die Renditebetrachtung.

    NEU (Beats Wunsch: "Schrift in Felder, eventuell alles einmitteln"):
    Label und Wert sind jetzt HORIZONTAL ZENTRIERT in der Kachel statt
    linksbuendig -- wirkt bei einer einzelnen, isolierten Kachel (statt einer
    fortlaufenden Tabellenzeile) aufgeraeumter."""
    import textwrap
    from matplotlib.patches import Rectangle
    bottom_in = top_in - height_in
    center_x_in = x_in + width_in / 2
    ax.add_patch(Rectangle(
        (x_in / fig_w, bottom_in / fig_h), width_in / fig_w, height_in / fig_h,
        transform=ax.transAxes, facecolor="#FAFAFA", edgecolor=_LINE, linewidth=1.0,
        zorder=2, clip_on=False,
    ))
    # NEU: Matplotlibs `wrap=True` haelt sich NICHT zuverlaessig an die
    # Kachelbreite (bekannter Fallstrick, siehe Kommentar bei der
    # Teil-Jahr-Hinweiszeile weiter unten in build_pdf_report()) -- ein langes
    # Label wie "Durchschnittliche Kapitalverzinsung" lief dadurch ueber den
    # Kachelrand hinaus. Deshalb hier manuell umbrechen (textwrap), mit einer
    # Zeichenbreite, die sich an der tatsaechlichen Kachelbreite orientiert.
    zeichen_pro_zeile = max(8, int((width_in - 0.3) / 0.062))
    label_zeilen = textwrap.wrap(label, width=zeichen_pro_zeile)
    ax.text(center_x_in / fig_w, (top_in - 0.30) / fig_h, "\n".join(label_zeilen), transform=ax.transAxes,
            fontsize=9.3, color=_TEXT_DIM, va="center", ha="center", linespacing=1.3)
    # NEU: Fallback-Werte wie "nicht amortisierbar" (z.B. wenn eine Anlage
    # rechnerisch nie amortisiert) sind deutlich laenger als die normalen
    # Zahlenwerte ("13.2 Jahre") -- bei fester fontsize=19 liefen sie ueber
    # die Kachelbreite hinaus und ueberlappten die naechste Kachel (gefunden
    # beim Testen mit einem Szenario, dessen Kapitalverzinsung negativ war).
    # Schriftgroesse deshalb an die Textlaenge anpassen, aehnlich zum
    # Wirtschaftlichkeit-Titel weiter oben.
    value_fontsize = 19 if len(value) <= 10 else (15 if len(value) <= 14 else 12)
    ax.text(center_x_in / fig_w, (bottom_in + height_in * 0.36) / fig_h, value,
            transform=ax.transAxes, fontsize=value_fontsize, fontweight="bold", color=_ACCENT,
            va="center", ha="center", family="monospace")


def _mini_chart_cell(fig, ax, fig_w, fig_h, x_in, top_in, width_in, height_in,
                      title, title_color, plot_fn, caption=None,
                      legend_bottom_h_in=0.0, legend_right_w_in=0.0):
    """Eine einzelne Mini-Chart-Zelle im 3x2-Raster von Seite 2: kleiner
    Titel oben (Topic + Ohne/Mit Batterie), darunter eine echte Matplotlib-
    Axes (per `plot_fn(cax)` befuellt), optional eine einzeilige Kennzahl-
    Caption darunter.

    NEU (Beats Rueckmeldung "Netzbezug-Legende ueberlagert den Chart, bitte
    fix unterhalb"/"Lastspitze-Legende ueberlagert die Balken, bitte fix
    rechts"): `legend_bottom_h_in`/`legend_right_w_in` reservieren
    zusaetzlichen Platz UNTERHALB bzw. RECHTS der eigentlichen Chart-Achse,
    in den die jeweilige `plot_fn` ihre Legende ausserhalb der Achsen-Bbox
    (per `bbox_to_anchor`) zeichnen kann, statt sie -- wie bisher via
    `loc="upper right"/"upper center"` -- ÜBER die Daten zu legen."""
    title_h_in = 0.20
    # tick_pad_in reserviert Platz UNTERHALB der Chart-Axes-Bbox fuer deren
    # eigene x-Achsen-Beschriftung (Monatsnamen bzw. rotierte "HH:MM"-Labels
    # bei den Netzbezug-Quartalscharts) -- matplotlib zeichnet Tick-Labels
    # AUSSERHALB der Axes-Bbox, ohne diesen Platz wuerden sie mit der
    # Caption-Zeile darunter kollidieren.
    tick_pad_in = 0.30
    caption_h_in = 0.16 if caption else 0.0
    ax.text((x_in + 0.02) / fig_w, (top_in - 0.09) / fig_h, title, transform=ax.transAxes,
            fontsize=9.3, fontweight="bold", color=title_color, va="center", ha="left")
    chart_top_in = top_in - title_h_in
    chart_bottom_in = top_in - height_in + caption_h_in + tick_pad_in + legend_bottom_h_in
    chart_width_in = width_in - legend_right_w_in
    cax = fig.add_axes([
        x_in / fig_w, chart_bottom_in / fig_h,
        chart_width_in / fig_w, (chart_top_in - chart_bottom_in) / fig_h,
    ])
    plot_fn(cax)
    if caption:
        ax.text((x_in + 0.02) / fig_w, (top_in - height_in + caption_h_in * 0.5) / fig_h, caption,
                transform=ax.transAxes, fontsize=7.6, color=_TEXT_DIM, va="center",
                ha="left", family="monospace")


def _fmt_param(value, suffix: str = "", decimals: int = 1, na: str = "--") -> str:
    """Formatiert einen Parameter-Wert fuer die Eingabeparameter-Uebersicht --
    einheitliche Behandlung von None/NaN (-> `na`) und Zahlen (mit Tausender-
    trennzeichen, `decimals` Nachkommastellen, optionalem Einheiten-Suffix)."""
    if value is None:
        return na
    if isinstance(value, float) and np.isnan(value):
        return na
    if isinstance(value, (int, float)):
        return f"{value:,.{decimals}f}".replace(",", "'") + (f" {suffix}" if suffix else "")
    return str(value) + (f" {suffix}" if suffix else "")


# NEU (Beats Rueckmeldung zur Eingabeparameter-Uebersicht): rohe Komma-Strings
# wie "0.2159,0.1965, 0.2159,0.1965" waren auf der PDF-Seite kaum lesbar --
# hier je Periode benannt darstellen ("Winter=.. / Sommer=.." bzw.
# "Q1=.. / Q2=.. / Q3=.. / Q4=.."), analog zur Saison-/Quartalslogik aus
# input_prep.py::parse_saisonal_werte()/saisonal_werte_zu_serie(). Bewusst ein
# eigener, fehlertoleranter Parser (gibt bei Problemen None zurueck statt zu
# werfen) statt eines Imports aus input_prep.py -- output_create.py soll wie
# bisher unabhaengig von den anderen Modulen bleiben (siehe Modulkopf-Kommentar
# zu den schlanken Abhaengigkeiten) und hier wird nur ANGEZEIGT, nicht gerechnet.
_SAISON_LABELS = ["Winter (Okt-Mär)", "Sommer (Apr-Sep)"]
_QUARTAL_LABELS = ["Q1", "Q2", "Q3", "Q4"]


def _parse_saisonal_anzeige(raw) -> list[float] | None:
    if raw is None:
        return None
    if isinstance(raw, (int, float)):
        if isinstance(raw, float) and np.isnan(raw):
            return None
        return [float(raw)]
    try:
        werte = [float(p.strip()) for p in str(raw).split(",")]
    except ValueError:
        return None
    if len(werte) not in (1, 2, 4):
        return None
    return werte


def _fmt_saisonal_werte(raw, suffix: str = "", decimals: int = 3, na: str = "--") -> str:
    """Formatiert 1/2/4 komma-getrennte Werte kompakt mit Perioden-Label
    (Beats Wunsch: statt eines unbenannten Komma-Strings "Q1=.., Q2=.." bzw.
    "Winter=.., Sommer=.." anzeigen). Faellt bei nicht parsbaren/unerwarteten
    Werten auf die normale _fmt_param()-Darstellung zurueck."""
    werte = _parse_saisonal_anzeige(raw)
    if werte is None:
        return _fmt_param(raw, suffix, decimals, na)
    if len(werte) == 1:
        return _fmt_param(werte[0], suffix, decimals, na)
    labels = _SAISON_LABELS if len(werte) == 2 else _QUARTAL_LABELS
    teile = [f"{lbl}={v:,.{decimals}f}".replace(",", "'") for lbl, v in zip(labels, werte)]
    return " / ".join(teile) + (f" {suffix}" if suffix else "")


def _parse_zeitfenster_anzeige(raw) -> list[str] | None:
    """Rein fuer Anzeigezwecke: parst eine Startzeit/Endzeit-HT-Zelle (ein
    oder mehrere komma-getrennte HH:MM-Werte, siehe
    input_prep.py::parse_zeitfenster_liste()) zu einer Liste von
    'HH:MM'-Strings. None bei Problemen (keine Exception, da hier nicht
    gerechnet wird)."""
    if raw is None:
        return None
    if isinstance(raw, dt.datetime):
        return [raw.strftime("%H:%M")]
    if isinstance(raw, dt.time):
        return [raw.strftime("%H:%M")]
    ergebnis = []
    for teil in str(raw).split(","):
        teil = teil.strip()
        if len(teil) < 4 or ":" not in teil:
            return None
        ergebnis.append(teil[:5])
    return ergebnis or None


def _fmt_ht_zeitfenster(start_raw, end_raw) -> str:
    starts = _parse_zeitfenster_anzeige(start_raw)
    ends = _parse_zeitfenster_anzeige(end_raw)
    if not starts or not ends or len(starts) != len(ends):
        return "--"
    return ", ".join(f"{s}–{e}" for s, e in zip(starts, ends))


_WOCHENTAGE = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]


def _fmt_ht_tage(tage_ht) -> str:
    """tage_ht = Anzahl Tage/Woche mit HT, ab Montag gezaehlt (siehe
    input_prep.py::build_ht_nt_bezugstarif()), z.B. 5 -> 'Mo–Fr'."""
    try:
        n = int(tage_ht)
    except (TypeError, ValueError):
        return _fmt_param(tage_ht, na="unbekannt")
    if n <= 0:
        return "keine (nur NT)"
    if n >= 7:
        return "Mo–So"
    return f"Mo–{_WOCHENTAGE[n - 1]}"


def _ist_wert_null(value) -> bool:
    """True, wenn value None/NaN oder (nach float-Konvertierung) 0 ist --
    z.B. um zu pruefen, ob ein Netz-Grenzwert de facto 'kein Netzbezug'
    bedeutet."""
    if value is None:
        return True
    try:
        v = float(value)
    except (TypeError, ValueError):
        return False
    return v == 0.0 or np.isnan(v)


SRL_MODUS_LABELS = {
    "ja_residual": "Post-hoc (Residual, SOC-gebandet)",
    "ja_optimiert": "Teil der Optimierung",
    "nein": "keine SRL-Vermarktung",
}

# NEU (Beats Rueckmeldung zu Seite 1 "System-Zusammenfassung": "Bei SRL bitte
# nur Residual (oder optimiert) angeben"): kuerzere Variante der obigen
# Labels fuer die knappe Zusammenfassungs-Tabelle -- Seite 0 (Eingabe-
# parameter) bleibt bei den ausfuehrlichen SRL_MODUS_LABELS.
_SRL_KURZ_LABELS = {
    "ja_residual": "Residual",
    "ja_optimiert": "Optimiert",
    "nein": "inaktiv",
}


def _ruecklieferung_kurz(params: dict, schema_key: str, fixtarif_key: str, floor_key: str | None = None) -> str:
    """Kompakte Einzeiler-Darstellung von Vergütungsschema + Tarif fuer die
    Seite-1-Zusammenfassung (Beats Wunsch: "Vergütungsschema und Tarife der
    Batterie und Solar angeben") -- eine schlankere Variante der
    ausfuehrlicheren Zeilen aus build_input_summary_rows()/Seite 0, die dort
    HKN/Markup als eigene Zeilen aufschluesseln; hier reicht Schema + der
    fuer dieses Schema massgebliche Tarifwert in EINER Zeile."""
    schema_wert = params.get(schema_key)
    schema_norm = str(schema_wert).strip().lower() if schema_wert is not None else ""
    schema_anzeige = _fmt_param(schema_wert, na="unbekannt")
    if schema_norm == "fixtarif":
        return f"{schema_anzeige} ({_fmt_param(params.get(fixtarif_key), 'CHF/kWh', 3)})"
    if schema_norm == "rmp_floor" and floor_key:
        return f"{schema_anzeige} (Floor {_fmt_param(params.get(floor_key), 'CHF/kWh', 3)})"
    return schema_anzeige


def build_input_summary_rows(
    params: dict,
    zeitreihen: pd.DataFrame,
    scenario_start: pd.Timestamp,
    n_steps: int,
    dt_hours: float,
    peakshaving_hat_tarif: bool,
) -> list[dict]:
    """Baut die Zeilen fuer die NEUE Eingabeparameter-Uebersichtsseite (Beats
    Wunsch: "können wir irgendwo im Output noch die 'wichtigen' Inputs
    festhalten? Zeitperiode, Tarifschema, PV-Leistung, Batterie-Leistung und
    -Kapazität, Netz- und Batterie-Grenzen, SRL-Modus und -Preisschema sowie
    Energie-Markup").

    Bewusst als eigene, unabhaengige Funktion (analog zu den anderen
    build_..._table()-Funktionen) -- braucht nur `params`/`zeitreihen`, kein
    battery_optimization-Import (siehe Modulkopf-Kommentar zu den schlanken
    Abhaengigkeiten von output_create.py). Der SRL-Modus-Mapping-Dict oben
    dupliziert bewusst (klein, 3 Eintraege) die Kern-Logik von
    battery_optimization.determine_srl_variante() rein fuer die Anzeige.
    """
    _, zeitraum_titel = _scenario_dauer(scenario_start, n_steps, dt_hours)
    zeitraum = zeitraum_titel.replace("Wirtschaftlichkeit — ", "")
    if zeitraum == "Jahresbetrachtung":
        ende = scenario_start + pd.Timedelta(hours=(n_steps - 1) * dt_hours)
        zeitraum = (
            f"{scenario_start.strftime('%d.%m.%Y')}–{ende.strftime('%d.%m.%Y')} "
            f"({n_steps * dt_hours / 24:.0f} Tage, volles Jahr)"
        )

    rows = [
        {"label": "Zeitperiode", "style": "header"},
        {"label": "Simulierter Zeitraum", "value": zeitraum},
    ]

    tarifschema = params.get("tarifschema")
    rows += [
        {"label": "Tarifschema (Bezug)", "style": "header"},
        {"label": "Schema", "value": _fmt_param(tarifschema, na="unbekannt")},
    ]
    schema_normalisiert = str(tarifschema).strip().upper() if tarifschema is not None else ""
    if schema_normalisiert == "SPOT":
        # NEU: Label und Wert je Zeile bewusst kurz gehalten (Label links,
        # Wert rechts, KEIN Umbruch in _draw_bordered_table()) -- eine
        # kombinierte "Aufschlag Bezug X% / Abschlag Lieferung Y%"-Zeile mit
        # zusaetzlichem Label-Praefix ueberlappte sich in der Bildmitte, da
        # Label+Wert zusammen breiter als die Tabellenspalte wurden.
        rows += [
            {"label": "Energie-Markup: Aufschlag Bezug", "value": _fmt_param(params.get("spot_aufschlag_bezug"), "%")},
        ]
        # NEU (Beats Hinweis zu Inputs_BAT_standalone.xlsx): zeigt, ob die
        # SwissIX-Preise automatisch von der ENTSO-E-API kamen oder von Hand
        # in Zeitreihen!B eingetragen wurden (siehe input_prep.py) -- nur
        # sichtbar, falls dieser Hinweis beim Erstellen von input_lp.xlsx
        # gesetzt wurde.
        if params.get("swissix_quelle"):
            rows.append({"label": "SwissIX-Preisquelle", "value": str(params["swissix_quelle"])})
    elif schema_normalisiert == "HT_NT":
        # NEU (Beats Rueckmeldung): HT und NT sauber getrennt (eigene Zeilen
        # statt einer zusammengequetschten "HT-Preis / NT-Preis"-Zeile), je
        # Periode benannt (Q1=../Winter=..) statt als rohem Komma-String, plus
        # Zeitfenster und Wochentage als eigene Zeilen.
        rows += [
            {"label": "HT-Preis", "value": _fmt_saisonal_werte(params.get("ht_preis"), "CHF/kWh", 3)},
            {"label": "NT-Preis", "value": _fmt_saisonal_werte(params.get("nt_preis"), "CHF/kWh", 3)},
            {"label": "HT-Zeitfenster", "value": _fmt_ht_zeitfenster(params.get("start_ht"), params.get("end_ht"))},
            {"label": "HT-Tage", "value": _fmt_ht_tage(params.get("tage_ht"))},
        ]
    else:
        rows.append({"label": "Hinweis", "value": "historischer Bezugstarif (Zeitreihen)"})

    # NEU (Beats Erweiterung: PV und Batterie koennen unterschiedlich
    # vermarktet werden -- unabhaengig vom obigen Bezugsschema, siehe
    # input_prep.py::resolve_rueckliefertarif()). PV zeigt weiterhin nur die
    # fuer das gewaehlte Schema tatsaechlich relevanten Werte (HKN jetzt mit
    # Perioden-Label statt rohem Komma-String, Beats Rueckmeldung).
    def _ruecklieferung_pv_zeilen(schema_key: str, hkn_key: str, floor_key: str, fixtarif_key: str):
        schema_wert = params.get(schema_key)
        zeilen = [{"label": "PV-Schema", "value": _fmt_param(schema_wert, na="unbekannt")}]
        schema_norm = str(schema_wert).strip().lower() if schema_wert is not None else ""
        if schema_norm == "spot":
            zeilen.append({"label": "Energie-Markup: Abschlag Lieferung", "value": _fmt_param(params.get("spot_abschlag_lieferung"), "%")})
        if schema_norm == "fixtarif":
            zeilen.append({"label": "PV-Fixtarif", "value": _fmt_param(params.get(fixtarif_key), "CHF/kWh", 3)})
        if schema_norm == "rmp_floor":
            zeilen.append({"label": "PV-Floor", "value": _fmt_param(params.get(floor_key), "CHF/kWh", 3)})
        if schema_norm in ("fixtarif", "rmp", "rmp_floor"):
            zeilen.append({"label": "PV-HKN", "value": _fmt_saisonal_werte(params.get(hkn_key), "CHF/kWh", 3)})
        return zeilen

    # NEU (Beats Rueckmeldung): fuer die Batterie reicht Schema + Fixtarif --
    # HKN nur anzeigen, wenn er auch tatsaechlich sauber zuordenbar ist: kein
    # Netzbezug fuer die Batterie zulaessig (max_bezug_batterie = 0, d.h. sie
    # kann NUR aus PV geladen werden) UND eine PV-Anlage vorhanden ist (sonst
    # gibt es gar keine PV-Energie, der die HKN zugeordnet werden koennte).
    def _ruecklieferung_batterie_zeilen(schema_key: str, hkn_key: str, fixtarif_key: str):
        schema_wert = params.get(schema_key)
        zeilen = [
            {"label": "Batterie-Schema", "value": _fmt_param(schema_wert, na="unbekannt")},
            {"label": "Batterie-Fixtarif", "value": _fmt_param(params.get(fixtarif_key), "CHF/kWh", 3)},
        ]
        kein_netzbezug_batterie = _ist_wert_null(params.get("max_bezug_batterie"))
        pv_vorhanden = not _ist_wert_null(params.get("dc_leistung"))
        if kein_netzbezug_batterie and pv_vorhanden:
            zeilen.append({"label": "Batterie-HKN", "value": _fmt_saisonal_werte(params.get(hkn_key), "CHF/kWh", 3)})
        return zeilen

    rows += [{"label": "Rücklieferung", "style": "header"}]
    rows += _ruecklieferung_pv_zeilen(
        "rueckliefer_pv_schema", "rueckliefer_pv_hkn", "rueckliefer_pv_floor", "rueckliefer_pv_fixtarif"
    )
    rows += _ruecklieferung_batterie_zeilen(
        "rueckliefer_batterie_schema", "rueckliefer_batterie_hkn", "rueckliefer_batterie_fixtarif",
    )

    rows += [
        {"label": "PV-Anlage", "style": "header"},
        {"label": "DC-Leistung", "value": _fmt_param(params.get("dc_leistung"), "kWp", 1, na="keine PV-Anlage")},
    ]
    # NEU (Beats Erweiterung 11.9.2026, Parameter!C27): ob die PV-Anlage
    # abgeregelt/gedrosselt werden darf (z.B. um Einspeisung bei negativen
    # Preisen zu vermeiden). Nur anzeigen, wenn ueberhaupt eine PV-Anlage
    # vorhanden ist. Fehlt die Zelle (aeltere Inputs.xlsx-Dateien), gilt
    # "Ja" (siehe battery_optimization.py::determine_pv_abschaltung_erlaubt) --
    # das wird hier ebenso dargestellt.
    if not _ist_wert_null(params.get("dc_leistung")):
        abschaltung_raw = params.get("pv_abschaltung_erlaubt")
        abschaltung_normalisiert = str(abschaltung_raw).strip().lower() if abschaltung_raw is not None else ""
        abschaltung_anzeige = "Nein" if abschaltung_normalisiert == "nein" else "Ja"
        rows.append({"label": "Abschaltung bei Negativpreisen", "value": abschaltung_anzeige})

    eta_laden = None
    eta_entladen = None
    if params.get("lade_wirkungsgrad") is not None and params.get("trafo_wirkungsgrad") is not None:
        eta_laden = params["lade_wirkungsgrad"] * params["trafo_wirkungsgrad"]
    if params.get("entlade_wirkungsgrad") is not None and params.get("trafo_wirkungsgrad") is not None:
        eta_entladen = params["entlade_wirkungsgrad"] * params["trafo_wirkungsgrad"]

    rows += [
        {"label": "Batterie", "style": "header"},
        {"label": "Leistung", "value": _fmt_param(params.get("leistung"), "kW")},
        {"label": "Kapazität", "value": _fmt_param(params.get("kapazitaet"), "kWh")},
        {"label": "Rundwirkungsgrad (Laden × Entladen)", "value": (
            _fmt_param(eta_laden * eta_entladen * 100 if eta_laden and eta_entladen else None, "%", 1)
        )},
        {"label": "SOC-Betriebsgrenzen", "value": (
            f"{_fmt_param((params.get('entladegrenze_soc') or 0) * 100, '%', 0)} – "
            f"{_fmt_param((params.get('ladegrenze_soc') or 0) * 100, '%', 0)}"
        )},
        {"label": "Netz-Grenzen Batterie (Einspeisung / Bezug)", "value": (
            f"{_fmt_param(params.get('max_einspeisung_batterie'), 'kW')} / "
            f"{_fmt_param(params.get('max_bezug_batterie'), 'kW')}"
        )},
    ]

    rows += [
        {"label": "Netz (Standort)", "style": "header"},
        {"label": "Netz-Grenzen Standort (Einspeisung / Bezug)", "value": (
            f"{_fmt_param(params.get('max_einspeisung'), 'kW')} / "
            f"{_fmt_param(params.get('max_bezug'), 'kW')}"
        )},
        {"label": "Peak-Shaving", "value": (
            f"aktiv, {_fmt_param(params.get('netznutzung_leistung'), 'CHF/kW/Monat', 2)}"
            if peakshaving_hat_tarif else "inaktiv"
        )},
    ]

    srl_teilnahme = params.get("srl_teilnahme")
    srl_normalisiert = str(srl_teilnahme).strip().lower() if srl_teilnahme is not None else ""
    srl_modus_label = SRL_MODUS_LABELS.get(srl_normalisiert, _fmt_param(srl_teilnahme, na="unbekannt"))
    rows += [
        {"label": "SRL (Sekundärregelleistung)", "style": "header"},
        {"label": "Modus", "value": srl_modus_label},
    ]
    if srl_normalisiert in ("ja_residual", "ja_optimiert"):
        # Preisschema-NAME (z.B. "Backcast_2025_Kunde", Parameter!C50)
        # anzeigen. NEU (Beats Rueckmeldung): keine Ø-Preise mehr -- nur
        # Modus + Preisschema, die Ø-Preis-Zeile wurde entfernt.
        rows.append({"label": "Preisschema", "value": _fmt_param(params.get("srl_preisschema"), na="unbekannt")})

    return rows


# Kategorial-Palette fuer die fuenf Ertragsmodule im kumulierten Flaechenplot
# (Seite 4, Beats Wunsch) -- eigene Palette statt _QCOLORS (Quartals-Linien),
# da hier FUENF statt vier Kategorien unterschieden werden muessen. Mit dem
# dataviz-Skill-Validator geprueft (`validate_palette.js ... --pairs all`):
# alle Normalsicht-Paare bestehen klar (>= 15.6), einzig Vermillion<->Gruen
# liegt unter Protanopie im 6-8-Bereich (WARN, nicht FAIL) -- dafuer traegt
# der Chart eine sichtbare Legende (Farbe nie alleinige Kodierung).
_ERTRAG_COLORS = {
    "Arbitrage": "#196B24",
    "SRL": "#0072B2",
    "Peak-Shaving": "#E69F00",
    "Eigenverbrauchsoptimierung": "#CC79A7",
    "Einspeiseoptimierung": "#D55E00",
}


def build_cumulative_ertrag_table(
    anteile: pd.DataFrame,
    params: dict,
    pv_profile: pd.Series,
    last_profile: pd.Series,
    peakshaving_hat_tarif: bool,
) -> pd.DataFrame:
    """Baut die Tabelle fuer die NEUE Seite 4 (Beats Wunsch: "eine weitere
    Seite im Report ... X-Achse ist die Datum, als Werte hätte ich gerne die
    kumulierten Erträge der einzelnen Bereiche (Arbitrage, SRL, Peakshaving,
    EV-Optimierung, Einspeiseoptimierung) in additiven Flächenplots"): pro
    Kalendertag summierte Ertraege je Modul, anschliessend ueber die Zeit
    kumulativ aufsummiert -- direkt als Input fuer matplotlib.stackplot()
    geeignet.

    Erwartet `anteile`/`pv_profile`/`last_profile` bereits in ECHTER
    Zuercher Lokalzeit (also NACH _to_zurich_local(), wie sie in
    build_pdf_report() ab dort ueberall verwendet werden) -- die
    Tagesgruppierung soll echte Kalendertage zeigen, keine DST-verschobenen.

    Vier der fuenf Module (Arbitrage, SRL, Eigenverbrauchsoptimierung,
    Einspeiseoptimierung) liegen in `anteile` bereits als ECHTE
    Pro-Zeitschritt-CHF-Fluesse vor (siehe allocate_modules()/main() in
    battery_optimization.py, "ertrag_srl" ist dort fuer BEIDE SRL-Varianten
    -- post_hoc UND optimiert -- einheitlich befuellt) und werden direkt
    uebernommen.

    Peak-Shaving ist die Ausnahme: `anteile["peak_shaving_kosten_monat"]`
    enthaelt bewusst NICHT die Ersparnis, sondern die tatsaechlich
    anfallenden (negativen) Kosten MIT Batterie, je Monat wiederholt (siehe
    Kommentar in battery_optimization.py::main() -- diese Spalte soll
    zeigen, was effektiv bezahlt wird). Die Ersparnis ggue. dem Referenzfall
    OHNE Batterie (= der eigentliche Modul-"Ertrag", identisch zur
    Definition von module["Peak-Shaving"] dort) wird hier NACHTRAEGLICH aus
    `pv_profile`/`last_profile` zurueckgerechnet und fuer die Darstellung
    als kumulierte Flaeche GLEICHMAESSIG auf die Zeitschritte des
    jeweiligen Monats verteilt (die Ersparnis wird de facto erst am
    Monatsende final, ein Sprung der kumulierten Kurve am Monatsende waere
    aber im Vergleich zu den anderen, echt kontinuierlich anfallenden
    Modulen visuell irrefuehrend).
    """
    idx = anteile.index

    if peakshaving_hat_tarif and "peak_shaving_kosten_monat" in anteile.columns:
        netznutzung = float(params["netznutzung_leistung"])
        baseline = np.maximum(last_profile.values - pv_profile.values, 0.0)
        baseline_series = pd.Series(baseline, index=idx)
        monat_key = pd.Series([(t.year, t.month) for t in idx], index=idx)
        baseline_peak_je_monat = baseline_series.groupby(monat_key).max()
        baseline_kosten_je_monat = baseline_peak_je_monat * netznutzung
        # anteile-Werte sind bereits NEGATIV (= -tatsaechliche Kosten MIT
        # Batterie), je Monat identisch wiederholt -- .first() genuegt.
        kosten_mit_je_monat = -anteile["peak_shaving_kosten_monat"].groupby(monat_key).first()
        ersparnis_je_monat = baseline_kosten_je_monat - kosten_mit_je_monat
        anzahl_je_monat = monat_key.value_counts()
        ersparnis_pro_schritt = monat_key.map(
            lambda m: ersparnis_je_monat[m] / anzahl_je_monat[m]
        )
        peak_shaving_ertrag = ersparnis_pro_schritt.values
    else:
        peak_shaving_ertrag = np.zeros(len(idx))

    fluesse = pd.DataFrame(
        {
            "Arbitrage": anteile["ertrag_arbitrage"].values,
            "SRL": anteile["ertrag_srl"].values,
            "Peak-Shaving": peak_shaving_ertrag,
            "Eigenverbrauchsoptimierung": anteile["ertrag_eigenverbrauch"].values,
            "Einspeiseoptimierung": anteile["ertrag_einspeiseoptimierung"].values,
        },
        index=idx,
    )

    tages_summe = fluesse.groupby(fluesse.index.date).sum()
    tages_summe.index = pd.to_datetime(tages_summe.index)
    return tages_summe.cumsum()


def build_pdf_report(
    output_path: str,
    module: dict,
    ts: pd.DataFrame,
    anteile: pd.DataFrame,
    zeitreihen: pd.DataFrame,
    params: dict,
    pv_profile: pd.Series,
    last_profile: pd.Series,
    kapitalkosten: dict,
    dt_hours: float = 0.25,
) -> dict:
    """Baut den PDF-Report im Fleco-Letterhead-Stil (auf Beats Wunsch an sein
    eigenes Fleco-Referenzdokument angenaehert, siehe Farbschema-Kommentar
    oben) auf VIER Seiten (Wirtschaftlichkeit/Charts bewusst auf zwei Seiten
    zusammengelegt, siehe frueherer Meilenstein -- NEU dazugekommen sind die
    Eingabeparameter-Uebersicht als Seite 1 UND die kumulierten Modul-
    Ertraege als Seite 4, beides Beats Wunsch):

      Seite 1 -- Eingabeparameter-Uebersicht: Zeitperiode, Tarifschema
        (inkl. SPOT-Energie-Markup bzw. HT/NT-Preise), PV-Leistung, Batterie
        (Leistung/Kapazitaet/Wirkungsgrad/SOC-Grenzen/Netz-Grenzen),
        Standort-Netzgrenzen/Peak-Shaving, SRL-Modus und -Preisschema --
        siehe build_input_summary_rows().
      Seite 2 -- Wirtschaftlichkeit UND Renditebetrachtung zusammen: Jahres-
        Einnahmen (nur aktivierte Module), Jahres-Ausgaben (Amortisation/
        Kapitalkosten/Unterhalt), Gewinn/Verlust, sowie darunter die drei
        Rendite-Kennzahlen (LCOS, Amortisationsdauer, Kapitalverzinsung) als
        Kennzahlen-Kacheln in einer Reihe.
      Seite 3 -- Alle drei Chart-Themen (Eigenverbrauch/Rueckspeisung pro
        Monat, Netzbezug-Tagesprofil je Quartal, monatliche Lastspitze) in
        einem 3x2-Raster: je Thema ZWEI Mini-Charts nebeneinander ("Ohne
        Batterie" grau als Referenzzustand vs. "Mit Batterie" gruen als
        tatsaechliches Optimierungsergebnis), mit einer knappen Kennzahl-
        Zeile darunter.
      Seite 4 -- NEU: kumulierte Ertraege je Modul (Arbitrage, SRL,
        Peak-Shaving, Eigenverbrauchsoptimierung, Einspeiseoptimierung) als
        additiver Flaechenplot ueber die Zeit (X-Achse = Kalendertag) --
        siehe build_cumulative_ertrag_table().

    `kapitalkosten` ist der Rueckgabewert von
    battery_optimization.capital_costs(params) -- wird hier NICHT selbst
    berechnet, damit output_create.py kein battery_optimization braucht
    (schlanke Abhaengigkeiten, wie bei build_income_expense_summary()).

    Gibt ein Dict mit allen zugrunde liegenden Tabellen zurueck (praktisch
    zum Cross-Check/Testen, ohne die PDF-Datei erneut oeffnen zu muessen).
    """
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_pdf import PdfPages

    plt.rcParams.update({
        "figure.facecolor": _BG,
        "axes.facecolor": _BG,
        "savefig.facecolor": _BG,
        "axes.edgecolor": _LINE,
        "axes.labelcolor": _TEXT_DIM,
        "text.color": _TEXT,
        "xtick.color": _TEXT_DIM,
        "ytick.color": _TEXT_DIM,
        "font.size": 10,
        "font.family": "sans-serif",
    })

    logo_img = _load_fleco_logo()

    # Start-Zeitstempel des Szenarios VOR der Zuercher-Lokalzeit-Umrechnung
    # merken (siehe unten): _scenario_dauer() braucht den urspruenglichen,
    # fest-CET-Start, NICHT den nach _to_zurich_local() umgerechneten --
    # sonst wuerde der Sommerzeit-Versatz (+1h, siehe dort) den angezeigten
    # Zeitraum im Titel systematisch nach vorne verschieben (z.B. "endet am
    # 01.07." statt "30.06.", obwohl das Szenario laut Eingabedatei exakt am
    # 30.06. endet).
    _scenario_start_orig = ts.index[0]

    # NEU (Beats Wunsch): alle Charts/Gruppierungen (Monat/Quartal/Uhrzeit) in
    # ECHTER Zuercher Ortszeit MIT Sommerzeit-Umstellung zeigen, statt der
    # intern verwendeten festen UTC+1-"CET" ohne DST -- siehe
    # _to_zurich_local(). Muss VOR allen Funktionen passieren, die ts.index/
    # pv_profile.index/last_profile.index fuer Monats-/Quartals-/Uhrzeit-
    # Gruppierung verwenden (build_monthly_peak_table(), build_monthly_
    # energy_table(), build_netzbezug_quarterly_profiles(), _build_quarterly_
    # avg_profile_generic() etc., alle weiter unten).
    ts, anteile, zeitreihen, pv_profile, last_profile = _to_zurich_local(
        ts, anteile, zeitreihen, pv_profile, last_profile
    )

    # NEU (Beats Wunsch): reines Batterie-System (kein PV, keine Last) --
    # dort ist ein "Ohne Batterie"-Referenzfall trivial (waere schlicht ein
    # System ohne jede Komponente) und "Eigenverbrauch"/"Rueckspeisung"
    # begrifflich nicht sinnvoll. Seite 2 zeigt in diesem Fall nur die
    # "Mit Batterie"-Charts (einspaltig, volle Breite) und ersetzt den
    # Eigenverbrauch/Rueckspeisung-Chart durch Laden/Entladen pro Monat.
    nur_batterie = float(pv_profile.max()) <= 1e-9 and float(last_profile.max()) <= 1e-9

    # NEU (Beats Frage "wieso sagt der Report eine Jahresbetrachtung, obwohl
    # es nur 3 Monate sind"): tatsaechlich simulierte Dauer ermitteln, damit
    # (a) der Titel den echten Zeitraum zeigt statt immer "Jahresbetrachtung"
    # und (b) Ausgaben/Kennzahlen konsistent mit dieser Dauer gerechnet
    # werden (siehe build_income_expense_summary()/build_rendite_kennzahlen()).
    dauer_jahre, wirtschaftlichkeit_titel = _scenario_dauer(_scenario_start_orig, len(ts.index), dt_hours)
    dauer_ist_vollejahr = wirtschaftlichkeit_titel == "Wirtschaftlichkeit — Jahresbetrachtung"

    zusammenfassung = build_income_expense_summary(module, kapitalkosten, dauer_jahre)
    rendite = build_rendite_kennzahlen(module, kapitalkosten, ts, dt_hours, dauer_jahre)
    peak_tabelle = build_monthly_peak_table(ts)
    if nur_batterie:
        laden_entladen_tabelle = build_monthly_charge_discharge_table(ts, dt_hours)
        profil_mit = _build_quarterly_avg_profile_generic(
            ts.index, ts["netzbezug_total"].values, label="Netzbezug mit Batterie"
        )
        # NEU (Beats Wunsch: "Quartalsplot für Netzeinspeisung hinzufügen,
        # gleiche Farben, gestrichelte Linien") -- im selben Chart wie
        # Netzbezug ueberlagert, siehe _plot_netzbezug() unten.
        profil_einspeisung_mit = _build_quarterly_avg_profile_generic(
            ts.index, ts["export"].values, label="Einspeisung mit Batterie"
        )
    else:
        monatstabelle = build_monthly_energy_table(ts, pv_profile, dt_hours)
        monatstabelle_ohne = build_baseline_monthly_energy_table(pv_profile, last_profile, dt_hours)
        profil_ohne, profil_mit = build_netzbezug_quarterly_profiles(ts, pv_profile, last_profile)
        profil_einspeisung_ohne, profil_einspeisung_mit = build_einspeisung_quarterly_profiles(
            ts, pv_profile, last_profile
        )
        peak_tabelle_ohne = build_baseline_monthly_peak_table(pv_profile, last_profile)

    netznutzung = params.get("netznutzung_leistung")
    peakshaving_hat_tarif = (
        netznutzung is not None
        and not (isinstance(netznutzung, float) and np.isnan(netznutzung))
        and float(netznutzung) != 0.0
    )

    # NEU (Beats Wunsch, neue Seite 4): kumulierte Ertraege je Modul, siehe
    # build_cumulative_ertrag_table(). Nutzt ts/anteile/pv_profile/
    # last_profile NACH der Zuercher-Lokalzeit-Umrechnung oben, damit die
    # Tagesgruppierung echte Kalendertage zeigt.
    kumulierte_ertraege = build_cumulative_ertrag_table(
        anteile, params, pv_profile, last_profile, peakshaving_hat_tarif
    )

    def _fmt_chf(v):
        return f"{v:,.0f} CHF".replace(",", "'")

    def _fmt_kennzahl(v, suffix, decimals=1):
        return f"{v:,.{decimals}f} {suffix}" if v is not None else "nicht amortisierbar"

    def _plot_eigenverbrauch(cax, tabelle, farbe1, farbe2, einheit_divisor=1.0, einheit_label="kWh", ylim=None):
        monate_idx = tabelle.index.astype(int)
        labels = [MONATSNAMEN[m - 1] for m in monate_idx]
        x = np.arange(len(labels))
        breite = 0.35
        cax.bar(x - breite / 2, tabelle["Eigenverbrauch_kWh"] / einheit_divisor, breite,
                label="EV", color=farbe1, zorder=3)
        cax.bar(x + breite / 2, tabelle["Rueckspeisung_kWh"] / einheit_divisor, breite,
                label="RS", color=farbe2, zorder=3)
        cax.set_xticks(x)
        cax.set_xticklabels(labels, fontsize=6)
        cax.set_ylabel(einheit_label, fontsize=7, color=_TEXT_DIM)
        cax.tick_params(labelsize=6)
        cax.legend(frameon=False, fontsize=6, labelcolor=_TEXT_DIM, loc="upper right")
        # NEU (Beats Wunsch): "Ohne"/"Mit"-Batterie-Paar soll dieselbe Y-Skala
        # verwenden, damit die beiden Charts auf einen Blick vergleichbar
        # sind (vorher skalierte jede Achse unabhaengig per Matplotlib-
        # Autoscale, was bei unterschiedlichen Maxima irrefuehrend war).
        if ylim is not None:
            cax.set_ylim(0, ylim)
        _pdf_style_axes(cax)

    def _plot_laden_entladen(cax, tabelle, farbe1, farbe2, einheit_divisor=1.0, einheit_label="kWh"):
        """Analog zu _plot_eigenverbrauch(), aber fuer ein reines
        Batterie-System (kein PV, keine Last) -- siehe
        build_monthly_charge_discharge_table()."""
        monate_idx = tabelle.index.astype(int)
        labels = [MONATSNAMEN[m - 1] for m in monate_idx]
        x = np.arange(len(labels))
        breite = 0.35
        cax.bar(x - breite / 2, tabelle["Laden_kWh"] / einheit_divisor, breite,
                label="Laden", color=farbe1, zorder=3)
        cax.bar(x + breite / 2, tabelle["Entladen_kWh"] / einheit_divisor, breite,
                label="Entladen", color=farbe2, zorder=3)
        cax.set_xticks(x)
        cax.set_xticklabels(labels, fontsize=7)
        cax.set_ylabel(einheit_label, fontsize=8, color=_TEXT_DIM)
        cax.tick_params(labelsize=7)
        cax.legend(frameon=False, fontsize=7, labelcolor=_TEXT_DIM, loc="upper right")
        _pdf_style_axes(cax)

    def _plot_peak_single(cax, peak_mit):
        """Wie _plot_peak_combined(), aber nur die "Mit Batterie"-Saeule --
        fuer ein reines Batterie-System gibt es keinen sinnvollen "Ohne
        Batterie"-Referenzfall (waere trivial 0, da weder PV noch Last)."""
        monate_idx = peak_mit.index.astype(int)
        labels = [MONATSNAMEN[m - 1] for m in monate_idx]
        x = np.arange(len(labels))
        cax.bar(x, peak_mit.values, 0.5, label="Mit Batterie", color=_ACCENT, zorder=3)
        cax.set_xticks(x)
        cax.set_xticklabels(labels, fontsize=7)
        cax.set_ylabel("kW", fontsize=8, color=_TEXT_DIM)
        cax.tick_params(labelsize=7)
        _pdf_style_axes(cax)

    def _plot_peak_combined(cax, peak_ohne, peak_mit):
        """NEU (Beats Wunsch): Ohne/Mit-Lastspitze in EINEM Chart statt zwei
        getrennten -- gruppierte Balken pro Monat (grau=Ohne, gruen=Mit),
        analog zum EV/RS-Chart oben."""
        monate_idx = peak_ohne.index.astype(int)
        labels = [MONATSNAMEN[m - 1] for m in monate_idx]
        x = np.arange(len(labels))
        breite = 0.35
        cax.bar(x - breite / 2, peak_ohne.values, breite, label="Ohne Batterie", color=_GRAU, zorder=3)
        cax.bar(x + breite / 2, peak_mit.values, breite, label="Mit Batterie", color=_ACCENT, zorder=3)
        cax.set_xticks(x)
        cax.set_xticklabels(labels, fontsize=7)
        cax.set_ylabel("kW", fontsize=8, color=_TEXT_DIM)
        cax.tick_params(labelsize=7)
        # NEU (Beats Rueckmeldung "Lastspitze-Legende ist nicht gut, fix
        # rechts der Grafik"): vorher `loc="upper right"` -- lag DIREKT IM
        # Chart und ueberlagerte dort haeufig die hohen Balken (z.B. Jan/Nov/
        # Dez). Jetzt fest RECHTS AUSSERHALB der Achse per `bbox_to_anchor`;
        # der dafuer noetige Platz wird von
        # `_mini_chart_cell(..., legend_right_w_in=...)` reserviert.
        cax.legend(frameon=False, fontsize=7.5, labelcolor=_TEXT_DIM,
                   loc="center left", bbox_to_anchor=(1.02, 0.5), bbox_transform=cax.transAxes)
        _pdf_style_axes(cax)

    def _plot_kumulierte_ertraege(cax, kumuliert):
        x = kumuliert.index
        alle_spalten = list(_ERTRAG_COLORS.keys())

        # NEU (Beats Rueckmeldung: "die Eigenverbrauchsoptimierung ist
        # negativ, wieso ist es in der Grafik dann positiv?"): ROOT CAUSE --
        # matplotlib.stackplot stapelt JEDE uebergebene Reihe additiv, egal
        # ob ihr Wert positiv oder negativ ist. Ein Modul, dessen kumulierter
        # Ertrag das ganze Jahr ueber NEGATIV ist (bestaetigt reproduziert:
        # min/max beide <= 0), wird trotzdem als normale, gefuellte Flaeche
        # OBEN auf dem Stapel gezeichnet -- optisch nicht von einem echten
        # positiven Beitrag zu unterscheiden, obwohl es den Gesamtertrag in
        # Wahrheit SCHMAELERT statt ihn zu vergroessern (bei kleinem Betrag
        # relativ zur Skala [hier: -354 CHF von ueber 30'000 CHF Total] ist
        # der eigentliche "Einbruch" der obersten Kontur mit blossem Auge
        # praktisch nicht erkennbar).
        #
        # FIX: Module, die zu IRGENDEINEM Zeitpunkt im Jahr negativ sind,
        # werden NICHT mehr additiv gestapelt, sondern als eigene, duenne
        # Linie bei ihrem TATSAECHLICHEN (ggf. negativen) kumulierten Wert
        # gezeichnet -- so ist auf den ersten Blick sichtbar, dass diese
        # Linie unterhalb der Null-Achse verlaeuft bzw. den Gesamtertrag
        # reduziert. Alle uebrigen (durchgehend nicht-negativen) Module
        # bleiben im additiven Flaechenstapel. Zusaetzlich zeigt JEDES
        # Legenden-Label den Jahresend-Wert in CHF -- damit ist die Aussage
        # der Tabelle auf Seite 2 (Wirtschaftlichkeit) IMMER konsistent mit
        # der Grafik hier, unabhaengig davon, wie fein der Unterschied
        # optisch sichtbar waere.
        eps = 1e-6
        stack_spalten = [c for c in alle_spalten if kumuliert[c].min() >= -eps]
        linien_spalten = [c for c in alle_spalten if kumuliert[c].min() < -eps]

        def _label_mit_wert(col):
            endwert = float(kumuliert[col].iloc[-1]) if len(kumuliert) else 0.0
            return f"{col} ({_fmt_chf(endwert)})"

        if stack_spalten:
            ys = [kumuliert[c].values for c in stack_spalten]
            farben = [_ERTRAG_COLORS[c] for c in stack_spalten]
            labels = [_label_mit_wert(c) for c in stack_spalten]
            # NEU (Beats Wunsch "kumulierte Erträge ... in additiven
            # Flächenplots"): echtes additives (gestapeltes) Flaechendiagramm
            # -- jede Flaeche zeigt den kumulierten Ertrag DIESES Moduls,
            # oben auf der Flaeche des vorherigen Moduls gestapelt. NUR NOCH
            # fuer Module, die nie negativ werden (siehe Fix-Kommentar oben)
            # -- die oberste Kontur zeigt daher die Summe DIESER Module,
            # NICHT mehr zwingend den gesamten operativen Ertrag (dafuer
            # steht weiterhin die Textzeile unter dem Chart). Duenne
            # Konturlinie in Seitenfarbe (_BG) zwischen den Flaechen
            # (dataviz-Skill: 2px-Flaechentrennung).
            cax.stackplot(x, ys, labels=labels, colors=farben, linewidth=0.6,
                          edgecolor=_BG, alpha=0.92, zorder=2)

        for c in linien_spalten:
            cax.plot(x, kumuliert[c].values, color=_ERTRAG_COLORS[c], linewidth=1.4,
                      linestyle="--", label=_label_mit_wert(c), zorder=3)

        cax.axhline(0, color=_LINE, linewidth=0.7, zorder=1)
        cax.set_ylabel("CHF (kumuliert)", fontsize=8, color=_TEXT_DIM)
        cax.tick_params(labelsize=7)
        import matplotlib.dates as mdates
        # NEU (Beats Wunsch "x-Achse monatlich beschriften"): vorher
        # AutoDateLocator (waehlte z.B. nur jeden 2. Monat) -- jetzt fix ein
        # Tick pro Kalendermonat (1. des Monats), unabhaengig von der
        # Zeitraumlaenge.
        cax.xaxis.set_major_locator(mdates.MonthLocator(bymonthday=1, interval=1))
        cax.xaxis.set_major_formatter(mdates.DateFormatter("%d.%m.%y"))
        for lbl in cax.get_xticklabels():
            lbl.set_rotation(30)
            lbl.set_ha("right")
        cax.legend(frameon=False, fontsize=7.5, labelcolor=_TEXT_DIM, ncol=min(3, len(alle_spalten)),
                   loc="upper center")
        _pdf_style_axes(cax)
        return linien_spalten

    def _plot_netzbezug(cax, profil, profil2=None, label2="Netzeinspeisung", ylim=None):
        """profil = Netzbezug-Quartalsprofile (durchgezogene Linien).
        profil2 (optional, NEU auf Beats Wunsch) = zweite Groesse (z.B.
        Netzeinspeisung) im SELBEN Chart ueberlagert -- gleiche Farbe pro
        Quartal wie profil, aber GESTRICHELT, damit Netzbezug/Netzeinspeisung
        auf einen Blick unterscheidbar bleiben, ohne die Quartals-Farblogik
        zu verdoppeln (4 statt 8 Farben). Zwei zusaetzliche graue
        Legenden-Eintraege (durchgezogen/gestrichelt) erklaeren, was die
        Linienart bedeutet -- Farbe allein codiert nur das Quartal."""
        import matplotlib.lines as mlines

        x = np.arange(len(profil.index))
        for i, col in enumerate(profil.columns):
            cax.plot(x, profil[col].values, color=_QCOLORS[i % len(_QCOLORS)], linewidth=1.3)
        if profil2 is not None:
            for i, col in enumerate(profil2.columns):
                cax.plot(x, profil2[col].values, color=_QCOLORS[i % len(_QCOLORS)], linewidth=1.3,
                          linestyle="--", alpha=0.85)
        tick_step = max(1, len(x) // 6)
        cax.set_xticks(x[::tick_step])
        cax.set_xticklabels(profil.index[::tick_step], rotation=45, ha="right", fontsize=6)
        cax.set_ylabel("kW", fontsize=7, color=_TEXT_DIM)
        cax.tick_params(labelsize=6)
        cax.axhline(0, color=_LINE, linewidth=0.7, zorder=1)

        handles = [mlines.Line2D([0], [0], color=_QCOLORS[i % len(_QCOLORS)], linewidth=1.3)
                   for i in range(len(profil.columns))]
        labels = list(profil.columns)
        if profil2 is not None:
            handles += [
                mlines.Line2D([0], [0], color=_TEXT_DIM, linewidth=1.3, linestyle="-"),
                mlines.Line2D([0], [0], color=_TEXT_DIM, linewidth=1.3, linestyle="--"),
            ]
            labels += ["Netzbezug", label2]
        # NEU (Beats Rueckmeldung "Netzbezug-Legende ist nicht gut, eventuell
        # fix unterhalb der Grafik"): vorher `loc="upper center"` -- lag
        # DIREKT IM Chart und ueberlagerte dort die Linien (gerade im oberen
        # Wertebereich, wo die Kurven tagsueber verlaufen). Jetzt fest
        # UNTERHALB der Achse per `bbox_to_anchor` (Achsen-Koordinaten, y<0
        # liegt unterhalb der x-Achse) -- der dafuer noetige Platz wird von
        # `_mini_chart_cell(..., legend_bottom_h_in=...)` reserviert.
        cax.legend(handles, labels, frameon=False, fontsize=6, labelcolor=_TEXT_DIM,
                   ncol=min(len(labels), 6), loc="upper center",
                   bbox_to_anchor=(0.5, -0.42), bbox_transform=cax.transAxes)
        # NEU (Beats Wunsch): gleiche Y-Skala fuer "Ohne"/"Mit Batterie" --
        # siehe Kommentar bei _plot_eigenverbrauch().
        if ylim is not None:
            cax.set_ylim(0, ylim)
        _pdf_style_axes(cax)


    with PdfPages(output_path) as pdf:
        # ---- Seite 1 (NEU, Beats Wunsch): Eingabeparameter-Uebersicht ------
        # "koennen wir irgendwo im Output noch die 'wichtigen' Inputs
        # festhalten?" -- Zeitperiode, Tarifschema, PV, Batterie, Netz-/
        # Batterie-Grenzen, SRL-Modus/-Preisschema, Energie-Markup. Als
        # eigene erste Seite, damit man beim spaeteren Nachschlagen eines
        # Reports sofort sieht, mit welchen Annahmen er gerechnet wurde,
        # ohne die urspruengliche Inputs.xlsx erneut oeffnen zu muessen.
        fig0, ax0, fig0_w, fig0_h, content0_top_in, content0_bottom_in = _fleco_page(
            (8.27, 11.69), "Eingabeparameter", logo_img, page_no="0"
        )
        margin0_in = 0.5
        table0_width_in = fig0_w - 2 * margin0_in
        input_rows = build_input_summary_rows(
            params, zeitreihen, _scenario_start_orig, len(ts.index), dt_hours, peakshaving_hat_tarif
        )
        _draw_bordered_table(ax0, fig0_w, fig0_h, margin0_in, content0_top_in, table0_width_in, input_rows)
        pdf.savefig(fig0)
        plt.close(fig0)

        # ---- Seite 1 (NEU, Beats Wunsch "Report formeller machen ... auf
        # der ersten Seite kurz zusammenfassen was alles im System ist"):
        # kurze System-Zusammenfassung -- ein Fliesstext-Absatz (Zeitraum,
        # Zweck) gefolgt von einer kompakten, kategorisierten Kennzahlen-
        # Liste (System/Tarif & Netz/Zusatzerloese), im Stil der "Angaben zu
        # Verbrauch und Erzeugung"-Box in Beats Referenzdokument. Bewusst
        # NUR die System-KONFIGURATION (nicht die wirtschaftlichen
        # Ergebnisse, dafuer steht Seite 2) -- Details zu jedem einzelnen
        # Parameter bleiben auf der ausfuehrlichen Seite 0.
        fig1, ax1, fig1_w, fig1_h, content1_top_in, content1_bottom_in = _fleco_page(
            (8.27, 11.69), "System-Zusammenfassung", logo_img, page_no="1"
        )
        margin1_in = 0.5
        box1_width_in = fig1_w - 2 * margin1_in

        pv_vorhanden_zsf = not _ist_wert_null(params.get("dc_leistung"))
        last_vorhanden_raw = params.get("last_vorhanden")
        last_vorhanden_zsf = str(last_vorhanden_raw).strip().lower() == "ja" if last_vorhanden_raw is not None else False
        zeitraum_txt = (
            f"{_scenario_start_orig.strftime('%d.%m.%Y')}"
            f"–{(ts.index[-1]).strftime('%d.%m.%Y')}" if len(ts.index) else "unbekannt"
        )
        komponenten = []
        komponenten.append("einer PV-Anlage" if pv_vorhanden_zsf else None)
        komponenten.append("einem Batteriespeicher")
        # NEU (Beats Rueckmeldung "der Lastgang ist nicht simuliert sondern
        # ist einfach ein Lastgang"): last_profile ist ein ECHTER, von Beat
        # bereitgestellter Lastgang, keine Simulation -- "simuliert" hier
        # entfernt (galt ohnehin nur fuer den ZEITRAUM/die Optimierung selbst,
        # nicht fuer den Lastgang als solchen).
        komponenten.append("einem Lastgang" if last_vorhanden_zsf else None)
        komponenten_txt = ", ".join(k for k in komponenten if k)

        zusammenfassung_text = (
            f"Dieser Bericht vergleicht den Betrieb mit und ohne Batteriespeicher über den "
            f"Zeitraum {zeitraum_txt} ({wirtschaftlichkeit_titel.split('—')[-1].strip()}). "
            f"Das analysierte System besteht aus {komponenten_txt}. Untersucht werden die "
            f"Effekte auf Eigenverbrauch, Netzbezug, Lastspitzen sowie die wirtschaftliche "
            f"Rendite der Investition; die zugrundeliegenden Annahmen sind im Detail auf "
            f"Seite 0 (Eingabeparameter) dokumentiert."
        )
        content1_top_in = _prose_box(
            ax1, fig1_w, fig1_h, margin1_in, content1_top_in, box1_width_in, zusammenfassung_text,
        )

        # NEU (Beats Rueckmeldung): "Last" zeigt statt vorhanden/keine die
        # tatsaechliche Jahreslast in MWh (Summe last_profile * dt_hours,
        # aus kWh in MWh umgerechnet) -- fuer ein reines Batterie-System
        # (kein Lastgang) bleibt es bei "keine".
        if last_vorhanden_zsf:
            jahreslast_mwh_zsf = float(np.sum(last_profile.values) * dt_hours) / 1000.0
            last_wert_zsf = _fmt_param(jahreslast_mwh_zsf, "MWh", 1)
        else:
            last_wert_zsf = "keine"

        # NEU (Beats Rueckmeldung "SRL bitte nur Residual/Optimiert und das
        # Preisschema angeben"): kurzes Label statt des ausfuehrlichen
        # SRL_MODUS_LABELS-Texts (der bleibt auf Seite 0), Preisschema in
        # derselben Zeile angehaengt, falls SRL ueberhaupt vermarktet wird.
        srl_teilnahme_zsf = params.get("srl_teilnahme")
        srl_normalisiert_zsf = str(srl_teilnahme_zsf).strip().lower() if srl_teilnahme_zsf is not None else ""
        srl_kurz_zsf = _SRL_KURZ_LABELS.get(srl_normalisiert_zsf, _fmt_param(srl_teilnahme_zsf, na="unbekannt"))
        if srl_normalisiert_zsf in ("ja_residual", "ja_optimiert"):
            srl_wert_zsf = f"{srl_kurz_zsf} — Preisschema: {_fmt_param(params.get('srl_preisschema'), na='unbekannt')}"
        else:
            srl_wert_zsf = srl_kurz_zsf

        # NEU (Beats Rueckmeldung "bei HT/NT bitte noch die Tarife angeben"):
        # HT/NT-Tarifzeilen nur, wenn das Bezugstarifschema tatsaechlich
        # HT_NT ist (analog zur Weiche in build_input_summary_rows()/Seite 0).
        tarifschema_zsf = params.get("tarifschema")
        tarifschema_norm_zsf = str(tarifschema_zsf).strip().upper() if tarifschema_zsf is not None else ""
        ht_nt_rows_zsf = []
        if tarifschema_norm_zsf == "HT_NT":
            ht_nt_rows_zsf = [
                {"label": "HT-Preis", "value": _fmt_saisonal_werte(params.get("ht_preis"), "CHF/kWh", 3)},
                {"label": "NT-Preis", "value": _fmt_saisonal_werte(params.get("nt_preis"), "CHF/kWh", 3)},
            ]

        # NEU (Beats Rueckmeldung "Vergütungsschema und Tarife der Batterie
        # und Solar angeben"): kompakte Einzeiler ueber _ruecklieferung_kurz()
        # -- Details/HKN weiterhin nur auf Seite 0.
        pv_verguetung_zsf = _ruecklieferung_kurz(
            params, "rueckliefer_pv_schema", "rueckliefer_pv_fixtarif", "rueckliefer_pv_floor"
        )
        batterie_verguetung_zsf = _ruecklieferung_kurz(
            params, "rueckliefer_batterie_schema", "rueckliefer_batterie_fixtarif"
        )

        zsf_rows = [
            {"label": "System", "style": "header"},
            {"label": "PV-Anlage", "value": (
                _fmt_param(params.get("dc_leistung"), "kWp", 1) if pv_vorhanden_zsf else "keine"
            )},
            {"label": "Batteriespeicher", "value": (
                f"{_fmt_param(params.get('leistung'), 'kW')} / {_fmt_param(params.get('kapazitaet'), 'kWh')}"
            )},
            {"label": "Batterie — Netzbezug / -abgabe Limiten", "value": (
                f"{_fmt_param(params.get('max_bezug_batterie'), 'kW')} / "
                f"{_fmt_param(params.get('max_einspeisung_batterie'), 'kW')}"
            )},
            {"label": "Last (Jahreslast)", "value": last_wert_zsf},
            {"label": "Tarif & Netz", "style": "header"},
            {"label": "Tarifschema (Bezug)", "value": _fmt_param(params.get("tarifschema"), na="unbekannt")},
            *ht_nt_rows_zsf,
            {"label": "Netzanschluss (Bezug / Einspeisung)", "value": (
                f"{_fmt_param(params.get('max_bezug'), 'kW')} / {_fmt_param(params.get('max_einspeisung'), 'kW')}"
            )},
            {"label": "PV-Vergütung (Rücklieferung)", "value": pv_verguetung_zsf} if pv_vorhanden_zsf else None,
            {"label": "Batterie-Vergütung (Rücklieferung)", "value": batterie_verguetung_zsf},
            {"label": "Zusätzliche Erlösquellen", "style": "header"},
            {"label": "Peak-Shaving", "value": (
                f"aktiv ({_fmt_param(params.get('netznutzung_leistung'), 'CHF/kW/Monat', 2)})"
                if peakshaving_hat_tarif else "inaktiv"
            )},
            {"label": "Sekundärregelleistung (SRL)", "value": srl_wert_zsf},
        ]
        zsf_rows = [r for r in zsf_rows if r is not None]
        _draw_bordered_table(ax1, fig1_w, fig1_h, margin1_in, content1_top_in, box1_width_in, zsf_rows)
        pdf.savefig(fig1)
        plt.close(fig1)

        # ---- Seite 2: Wirtschaftlichkeit + Renditebetrachtung -------------
        fig, ax, fig_w, fig_h, content_top_in, content_bottom_in = _fleco_page(
            (8.27, 11.69), "Wirtschaftlichkeit & Renditebetrachtung", logo_img, page_no="2"
        )
        margin_in = 0.5
        table_width_in = fig_w - 2 * margin_in

        rows = [{"label": "Einnahmen / Ersparnisse", "style": "header"}]
        for label, val in zusammenfassung["einnahmen"].items():
            rows.append({"label": label, "value": _fmt_chf(val)})
        rows.append({"label": "Total Einnahmen", "value": _fmt_chf(zusammenfassung["total_einnahmen"]),
                     "style": "total"})
        rows.append({"label": "Ausgaben", "style": "header"})
        for label, val in zusammenfassung["ausgaben"].items():
            rows.append({"label": label, "value": _fmt_chf(-val)})
        rows.append({"label": "Total Ausgaben", "value": _fmt_chf(-zusammenfassung["total_ausgaben"]),
                     "style": "total"})
        gewinn = zusammenfassung["gewinn_verlust"]
        rows.append({"label": "Gewinn / Verlust (netto)", "value": _fmt_chf(gewinn), "style": "result",
                     "color": _POS if gewinn >= 0 else _NEG})

        _wirtschaftlichkeit_titel_size = 13 if dauer_ist_vollejahr else 11
        _section_label(ax, fig_w, fig_h, margin_in, content_top_in, wirtschaftlichkeit_titel,
                        size=_wirtschaftlichkeit_titel_size)
        table_top_in = content_top_in - 0.34
        top_in = _draw_bordered_table(ax, fig_w, fig_h, margin_in, table_top_in, table_width_in, rows)

        top_in -= 0.5
        _section_label(ax, fig_w, fig_h, margin_in, top_in, "Renditebetrachtung")
        top_in -= 0.34

        kennzahlen = [
            ("Levelized Cost of Storage", _fmt_kennzahl(rendite["lcos_rp_kwh"], "Rp/kWh")),
            ("Amortisationsdauer", _fmt_kennzahl(rendite["amortisationsdauer_jahre"], "Jahre")),
            ("Durchschnittliche Kapitalverzinsung", _fmt_kennzahl(rendite["kapitalverzinsung_pct"], "%")),
        ]
        tile_gap_in = 0.22
        tile_w_in = (table_width_in - 2 * tile_gap_in) / 3
        tile_h_in = 1.5
        x_in = margin_in
        for label, val in kennzahlen:
            _kpi_tile(ax, fig_w, fig_h, x_in, top_in, tile_w_in, tile_h_in, label, val)
            x_in += tile_w_in + tile_gap_in

        # NEU (Beats Frage "wie kann die Amortisationsdauer 14 Jahre sein und
        # die jaehrliche Wirtschaftsrechnung gleichzeitig negativ?"): die drei
        # Kennzahlen oben rechnen bewusst OHNE Amortisation/Kapitalkosten
        # (sonst waere "wie lange dauert die Amortisation" zirkulaer, siehe
        # build_rendite_kennzahlen()) -- sie zeigen also eine ANDERE
        # Finanzierungsannahme (kreditfreie Investition) als die Tabelle oben
        # ("Gewinn/Verlust netto", die ein Annuitaetendarlehen unterstellt).
        # Beide Zahlen koennen sich deshalb scheinbar widersprechen, ohne dass
        # ein Rechenfehler vorliegt -- das wird hier IMMER (nicht nur bei
        # Teil-Jahr-Szenarien) explizit erklaert, damit das nicht jedes Mal
        # zu Rueckfragen fuehrt.
        import textwrap
        top_in -= tile_h_in + 0.22
        line_h_in = 7.6 * 1.5 / 72.0  # Zeilenhoehe bei fontsize=7.6, linespacing=1.5

        # NEU (Beats Wunsch: "kannst du diese Erklaerung mit in den PDF-Report
        # nehmen -- anstelle des Hinweistextes, oder diesen kuerzen falls er
        # noch Platz hat"): Beat hatte im Chat nach der Bedeutung von
        # "Einspeiseoptimierung" gefragt -- die Zeilen der Einnahmen-Tabelle
        # weiter oben (Eigenverbrauchsoptimierung/Einspeiseoptimierung/
        # Arbitrage/SRL/Peak-Shaving) sind ohne Erklaerung nicht selbsterklaerend.
        # Statt des bisherigen, sehr ausfuehrlichen Methodik-Hinweistexts steht
        # hier jetzt zuerst dieses kompakte Modul-Glossar; der Methodik-Hinweis
        # (Amortisationsdauer vs. Jahresergebnis) bleibt bestehen, aber stark
        # gekuerzt, da fuer beides zusammen sonst kein Platz mehr auf der Seite
        # waere.
        # NEU (Beats Rueckmeldung "Module in der Einnahmen-Tabelle oben:
        # weglassen"): keine Einleitungszeile mehr vor dem Modul-Glossar --
        # die fett gesetzten Modulnamen darin sind selbsterklaerend genug.
        top_in = _modul_glossar_box(fig, ax, fig_w, fig_h, margin_in, top_in, table_width_in, _MODUL_GLOSSAR)
        top_in -= 0.09

        methodik_text = (
            "Hinweis zur Methodik: LCOS/Amortisationsdauer/Kapitalverzinsung rechnen ohne die "
            "Positionen „Amortisation“/„Kapitalkosten“ (kreditfreie Investition) -- „Gewinn/Verlust "
            "(netto)“ oben unterstellt dagegen ein Annuitätendarlehen. Beide Sichtweisen können "
            "deshalb auseinanderlaufen."
        )
        methodik_zeilen = textwrap.wrap(methodik_text, width=100)
        ax.text(
            margin_in / fig_w, top_in / fig_h, "\n".join(methodik_zeilen),
            transform=ax.transAxes, fontsize=7.6, color=_TEXT_DIM,
            va="top", ha="left", style="italic", linespacing=1.5,
        )
        top_in -= line_h_in * len(methodik_zeilen) + 0.12

        # NEU (Beats Frage "wieso sagt der Report eine Jahresbetrachtung,
        # obwohl es nur 3 Monate sind"): bei einem Teil-Jahr-Szenario sind
        # diese drei Kennzahlen auf Basis der beobachteten Werte auf ein
        # volles Jahr HOCHGERECHNET (siehe build_rendite_kennzahlen()) --
        # das wird hier explizit vermerkt, damit es nicht mit einer echten
        # Jahresmessung verwechselt wird.
        if not dauer_ist_vollejahr:
            # Text ist bei realistischen Seitenbreiten (A4, 0.5in Rand) zu
            # lang fuer eine Zeile -- manuell umbrechen (textwrap statt
            # matplotlibs `wrap=True`, das ohne explizite Pixel-Breite
            # unzuverlaessig ist) und mit `va="top"` ausrichten, damit die
            # Position unabhaengig von der Zeilenzahl stimmt.
            hinweis_text = (
                f"Hinweis: simulierter Zeitraum = {dauer_jahre * 365.25:.0f} Tage. "
                "Die Renditekennzahlen oben sind auf ein volles Jahr hochgerechnet "
                "(beobachtete Werte / simulierter Zeitraum) -- keine Messung ueber "
                "ein tatsaechliches Jahr."
            )
            hinweis_zeilen = textwrap.wrap(hinweis_text, width=100)
            ax.text(
                margin_in / fig_w, top_in / fig_h, "\n".join(hinweis_zeilen),
                transform=ax.transAxes, fontsize=7.6, color=_TEXT_DIM,
                va="top", ha="left", style="italic", linespacing=1.5,
            )
        pdf.savefig(fig)
        plt.close(fig)

        # ---- Seite 3: Alle Charts in einem 3x2-Raster (bzw. einspaltig bei
        # einem reinen Batterie-System, siehe nur_batterie oben) -----------
        seite2_titel = (
            "Batterieladung, Netzbezug & Lastspitze"
            if nur_batterie else
            "Eigenverbrauch, Netzbezug & Lastspitze — Ohne/Mit Batterie im Vergleich"
        )
        fig, ax, fig_w, fig_h, content_top_in, content_bottom_in = _fleco_page(
            (11.69, 8.27), seite2_titel, logo_img, page_no="3",
        )
        margin_in = 0.5
        col_gap_in = 0.28
        row_gap_in = 0.30
        col_w_in = (fig_w - 2 * margin_in - col_gap_in) / 2
        content_h_in = content_top_in - content_bottom_in
        row_h_in = (content_h_in - 2 * row_gap_in) / 3
        voll_breite_in = fig_w - 2 * margin_in

        # NEU (Beats Rueckmeldung "Legenden ueberlagern die Grafiken"):
        # reservierter Zusatzplatz fuer die Netzbezug-Legende UNTERHALB
        # bzw. die Lastspitze-Legende RECHTS der jeweiligen Chart-Achse
        # (siehe _mini_chart_cell()/_plot_netzbezug()/_plot_peak_combined()
        # oben). Werte empirisch per Testrender ermittelt: die Netzbezug-
        # Legende ist bei `bbox_to_anchor=(0.5, -0.42)` eine einzeilige
        # Legende (fontsize 6) -- 0.42in Puffer reicht fuer den Abstand
        # plus die Legendenbox selbst. Die Lastspitze-Legende hat nur zwei
        # Eintraege ("Ohne"/"Mit Batterie", fontsize 7.5) vertikal gestapelt
        # rechts der Achse -- 1.15in reicht fuer den laengeren Text.
        _NETZBEZUG_LEGEND_H_IN = 0.42
        _LASTSPITZE_LEGEND_W_IN = 1.15

        if nur_batterie:
            # NEU (Beats Wunsch): reines Batterie-System -- kein Ohne/Mit-
            # Vergleich (der "Ohne Batterie"-Fall waere trivial, da weder PV
            # noch Last existieren), stattdessen drei einspaltige Charts ueber
            # die volle Breite: Laden/Entladen statt Eigenverbrauch/
            # Rueckspeisung, Netzbezug-Tagesprofil nur "Mit Batterie", und die
            # monatliche Lastspitze ebenfalls nur "Mit Batterie".
            max_monatswert_kwh = max(
                laden_entladen_tabelle["Laden_kWh"].max(), laden_entladen_tabelle["Entladen_kWh"].max(),
            )
            einheit_divisor, einheit_label = (1000.0, "MWh") if max_monatswert_kwh >= 1000 else (1.0, "kWh")

            row_top_in = content_top_in
            _mini_chart_cell(
                fig, ax, fig_w, fig_h, margin_in, row_top_in, voll_breite_in, row_h_in,
                title="Laden & Entladen pro Monat — Mit Batterie", title_color=_ACCENT,
                plot_fn=lambda cax: _plot_laden_entladen(
                    cax, laden_entladen_tabelle, _ACCENT, _ACCENT2, einheit_divisor, einheit_label),
                caption=(
                    f"Total Laden: {laden_entladen_tabelle['Laden_kWh'].sum() / einheit_divisor:,.0f} "
                    f"{einheit_label} · Total Entladen: "
                    f"{laden_entladen_tabelle['Entladen_kWh'].sum() / einheit_divisor:,.0f} {einheit_label}"
                ),
            )
            row_top_in -= row_h_in + row_gap_in

            # NEU (Beats Wunsch: "Netzbezug und Einspeisung ist mir zu
            # ueberladen"): frueher EIN Chart mit Netzbezug (durchgezogen) UND
            # Netzeinspeisung (gestrichelt) je Quartal ueberlagert -- bei acht
            # gleichzeitig sichtbaren Linien zu unuebersichtlich. Jetzt ZWEI
            # separate, halbbreite Charts nebeneinander (je 4 Linien/Quartale,
            # kein Ueberlagern/Linienart-Unterscheiden mehr noetig) --
            # _plot_netzbezug() OHNE profil2 zeigt automatisch nur die
            # einfache Q1-Q4-Legende.
            _mini_chart_cell(
                fig, ax, fig_w, fig_h, margin_in, row_top_in, col_w_in, row_h_in,
                title="Netzbezug — Tagesprofil je Quartal — Mit Batterie", title_color=_ACCENT,
                plot_fn=lambda cax: _plot_netzbezug(cax, profil_mit),
                caption=f"Ø Netzbezug: {np.nanmean(profil_mit.values):,.1f} kW",
                legend_bottom_h_in=_NETZBEZUG_LEGEND_H_IN,
            )
            _mini_chart_cell(
                fig, ax, fig_w, fig_h, margin_in + col_w_in + col_gap_in, row_top_in, col_w_in, row_h_in,
                title="Netzeinspeisung — Tagesprofil je Quartal — Mit Batterie", title_color=_ACCENT,
                plot_fn=lambda cax: _plot_netzbezug(cax, profil_einspeisung_mit),
                caption=f"Ø Netzeinspeisung: {np.nanmean(profil_einspeisung_mit.values):,.1f} kW",
                legend_bottom_h_in=_NETZBEZUG_LEGEND_H_IN,
            )
            row_top_in -= row_h_in + row_gap_in

            _mini_chart_cell(
                fig, ax, fig_w, fig_h, margin_in, row_top_in, voll_breite_in, row_h_in,
                title="Monatliche Lastspitze — Mit Batterie", title_color=_ACCENT,
                plot_fn=lambda cax: _plot_peak_single(cax, peak_tabelle),
                caption=f"Ø {peak_tabelle.values.mean():,.0f} kW (Max {peak_tabelle.values.max():,.0f} kW)",
            )
        else:
            total_ev = float(monatstabelle["Eigenverbrauch_kWh"].sum())
            total_ev_ohne = float(monatstabelle_ohne["Eigenverbrauch_kWh"].sum())
            total_rs = float(monatstabelle["Rueckspeisung_kWh"].sum())
            total_rs_ohne = float(monatstabelle_ohne["Rueckspeisung_kWh"].sum())
            quote_mit = (total_ev / (total_ev + total_rs) * 100.0) if (total_ev + total_rs) > 1e-9 else 0.0
            quote_ohne = (
                total_ev_ohne / (total_ev_ohne + total_rs_ohne) * 100.0
            ) if (total_ev_ohne + total_rs_ohne) > 1e-9 else 0.0

            # NEU (Beats Wunsch): EV/Rueckspeisung automatisch in MWh anzeigen,
            # sobald die monatlichen Werte "gross" werden -- Schwelle 1000 kWh
            # (=1 MWh), gemeinsam fuer Ohne/Mit-Chart ermittelt, damit beide
            # Balkendiagramme dieselbe Einheit/Skala verwenden und vergleichbar
            # bleiben.
            max_monatswert_kwh = max(
                monatstabelle["Eigenverbrauch_kWh"].max(), monatstabelle["Rueckspeisung_kWh"].max(),
                monatstabelle_ohne["Eigenverbrauch_kWh"].max(), monatstabelle_ohne["Rueckspeisung_kWh"].max(),
            )
            if max_monatswert_kwh >= 1000:
                einheit_divisor, einheit_label = 1000.0, "MWh"
            else:
                einheit_divisor, einheit_label = 1.0, "kWh"

            ersparnis_txt = ""
            if peakshaving_hat_tarif:
                ersparnis = float(
                    sum((peak_tabelle_ohne[m] - peak_tabelle[m]) * float(netznutzung)
                        for m in peak_tabelle.index)
                )
                ersparnis_txt = f" · Ersparnis: {_fmt_chf(ersparnis)}"

            # NEU (Beats Wunsch): "Ohne"/"Mit Batterie" je Zeile auf DIESELBE
            # Y-Skala bringen, statt jede Achse unabhaengig automatisch
            # skalieren zu lassen (vorher z.B. beim Netzbezug-Tagesprofil:
            # "Ohne" bis 150 kW, "Mit" nur bis 125 kW -- auf den ersten Blick
            # sah es dann so aus, als waeren die Kurvenverlaeufe nicht direkt
            # vergleichbar). Faktor 1.08 = etwas Luft ueber dem hoechsten
            # Balken/der hoechsten Linie, analog zu Matplotlibs eigenem
            # Autoscale-Rand.
            ev_ylim = (max_monatswert_kwh / einheit_divisor) * 1.08
            netzbezug_ylim = max(
                np.nanmax(profil_ohne.values), np.nanmax(profil_mit.values),
                np.nanmax(profil_einspeisung_ohne.values), np.nanmax(profil_einspeisung_mit.values),
            ) * 1.08

            zeilen_zweispaltig = [
                {
                    "titel": "Eigenverbrauch & Rueckspeisung pro Monat",
                    "legend_bottom_h_in": 0.0,
                    "ohne": {"plot": lambda cax: _plot_eigenverbrauch(
                                 cax, monatstabelle_ohne, _GRAU, _GRAU2, einheit_divisor, einheit_label,
                                 ylim=ev_ylim),
                             "caption": f"Eigenverbrauchsquote: {quote_ohne:,.0f} %"},
                    "mit": {"plot": lambda cax: _plot_eigenverbrauch(
                                cax, monatstabelle, _ACCENT, _ACCENT2, einheit_divisor, einheit_label,
                                ylim=ev_ylim),
                            "caption": f"Eigenverbrauchsquote: {quote_mit:,.0f} %"},
                },
                {
                    "titel": "Netzbezug & Netzeinspeisung — Tagesprofil je Quartal",
                    # NEU: diese Zeile nutzt _plot_netzbezug() MIT profil2
                    # (Netzbezug + Netzeinspeisung ueberlagert) -- die Legende
                    # braucht daher reservierten Platz UNTERHALB der Achse
                    # (siehe _NETZBEZUG_LEGEND_H_IN oben). Die Eigenverbrauch-
                    # Zeile oben zeichnet ihre (kurze) Legende weiterhin IM
                    # Chart, dort besteht kein Ueberlagerungsproblem.
                    "legend_bottom_h_in": _NETZBEZUG_LEGEND_H_IN,
                    "ohne": {"plot": lambda cax: _plot_netzbezug(
                                 cax, profil_ohne, profil_einspeisung_ohne, ylim=netzbezug_ylim),
                             "caption": (
                                 f"Ø Netzbezug: {np.nanmean(profil_ohne.values):,.1f} kW · "
                                 f"Ø Netzeinspeisung: {np.nanmean(profil_einspeisung_ohne.values):,.1f} kW"
                             )},
                    "mit": {"plot": lambda cax: _plot_netzbezug(
                                cax, profil_mit, profil_einspeisung_mit, ylim=netzbezug_ylim),
                            "caption": (
                                f"Ø Netzbezug: {np.nanmean(profil_mit.values):,.1f} kW · "
                                f"Ø Netzeinspeisung: {np.nanmean(profil_einspeisung_mit.values):,.1f} kW"
                            )},
                },
            ]

            # NEU (Beats Wunsch "Lastspitzen in EINE Grafik packen"): Ohne/Mit
            # nicht mehr als zwei getrennte Mini-Charts, sondern EIN Chart mit
            # gruppierten Balken (analog zum EV/RS-Chart oben), ueber die volle
            # Zeilenbreite.
            peak_caption = (
                f"Ø Ohne {peak_tabelle_ohne.values.mean():,.0f} kW (Max {peak_tabelle_ohne.values.max():,.0f} kW)"
                f"  ·  Ø Mit {peak_tabelle.values.mean():,.0f} kW (Max {peak_tabelle.values.max():,.0f} kW)"
                f"{ersparnis_txt}"
            )

            row_top_in = content_top_in
            for zeile in zeilen_zweispaltig:
                x_in = margin_in
                for seite, farbe, label in [("ohne", _GRAU, "Ohne Batterie"), ("mit", _ACCENT, "Mit Batterie")]:
                    cell = zeile[seite]
                    _mini_chart_cell(
                        fig, ax, fig_w, fig_h, x_in, row_top_in, col_w_in, row_h_in,
                        title=f"{zeile['titel']} — {label}", title_color=farbe,
                        plot_fn=cell["plot"], caption=cell["caption"],
                        legend_bottom_h_in=zeile.get("legend_bottom_h_in", 0.0),
                    )
                    x_in += col_w_in + col_gap_in
                row_top_in -= row_h_in + row_gap_in

            _mini_chart_cell(
                fig, ax, fig_w, fig_h, margin_in, row_top_in, 2 * col_w_in + col_gap_in, row_h_in,
                title="Monatliche Lastspitze — Ohne/Mit Batterie im Vergleich", title_color=_ACCENT,
                plot_fn=lambda cax: _plot_peak_combined(cax, peak_tabelle_ohne, peak_tabelle),
                caption=peak_caption,
                legend_right_w_in=_LASTSPITZE_LEGEND_W_IN,
            )

        pdf.savefig(fig)
        plt.close(fig)

        # ---- Seite 4 (NEU, Beats Wunsch): kumulierte Ertraege je Modul als
        # additiver Flaechenplot ueber die Zeit ("X-Achse ist die Datum, als
        # Werte hätte ich gerne die kumulierten Erträge der einzelnen
        # Bereiche ... in additiven Flächenplots") ------------------------
        fig, ax, fig_w, fig_h, content_top_in, content_bottom_in = _fleco_page(
            (11.69, 8.27), "Kumulierte Erträge je Modul", logo_img, page_no="4",
        )
        margin_in = 0.6
        # tick_pad_in reserviert Platz UNTERHALB der Chart-Axes-Bbox fuer die
        # rotierten "TT.MM.JJ"-Datums-Ticks (matplotlib zeichnet Tick-Labels
        # AUSSERHALB der Axes-Bbox, analog zu _mini_chart_cell() oben);
        # caption_h_in reserviert zusaetzlich Platz DARUNTER fuer die
        # Total-/Hinweis-Zeile(n), sonst kollidieren beide (siehe erste
        # Test-Iteration). ZWEIZEILIG, wenn mind. ein Modul negativ ist (NEU,
        # siehe Fix-Kommentar in _plot_kumulierte_ertraege) -- eine Zeile
        # reichte dafuer nicht (lief in einem Testrender rechts aus der
        # Seite), daher eigene zweite Zeile statt eines noch laengeren
        # Einzeilers.
        tick_pad_in = 0.42
        # NEU: vorab (unabhaengig vom eigentlichen Plot-Aufruf unten) pruefen,
        # ob mind. ein Modul negativ wird -- dann braucht die Fussnote eine
        # zweite Zeile und damit mehr reservierten Platz (siehe
        # _plot_kumulierte_ertraege fuer dieselbe Pruefung/denselben Epsilon).
        _hat_negativ_modul = any(
            kumulierte_ertraege[c].min() < -1e-6 for c in kumulierte_ertraege.columns
        ) if len(kumulierte_ertraege) else False
        caption_h_in = 0.36 if _hat_negativ_modul else 0.26
        chart_bottom_in = content_bottom_in + tick_pad_in + caption_h_in
        cax4 = fig.add_axes([
            margin_in / fig_w, chart_bottom_in / fig_h,
            (fig_w - 2 * margin_in) / fig_w, (content_top_in - chart_bottom_in) / fig_h,
        ])
        negativ_module = _plot_kumulierte_ertraege(cax4, kumulierte_ertraege)

        total_kumuliert = (
            float(kumulierte_ertraege.iloc[-1].sum()) if len(kumulierte_ertraege) else 0.0
        )
        letzter_tag = (
            kumulierte_ertraege.index[-1].strftime("%d.%m.%Y") if len(kumulierte_ertraege) else "--"
        )
        zeile1 = (
            f"Kumulierter operativer Gesamtertrag per {letzter_tag}: {_fmt_chf(total_kumuliert)} "
            "(Summe aller Module, ohne Kapitalkosten/Amortisation — siehe Wirtschaftlichkeit Seite 2)"
        )
        ax.text(
            margin_in / fig_w, (content_bottom_in + (0.16 if negativ_module else 0.0)) / fig_h,
            zeile1,
            transform=ax.transAxes, fontsize=7.6, color=_TEXT_DIM, va="bottom", ha="left",
        )
        # NEU (Beats Rueckmeldung, siehe Fix-Kommentar in
        # _plot_kumulierte_ertraege): Module, die als eigene (gestrichelte,
        # ggf. negative) Linie statt als Teil der Flaeche gezeichnet werden,
        # hier in einer EIGENEN zweiten Zeile explizit benennen -- sonst
        # bleibt unklar, warum sie nicht Teil des additiven Stapels sind.
        if negativ_module:
            zeile2 = (
                f"{', '.join(negativ_module)}: negativer Jahreswert, daher als gestrichelte "
                "Linie statt additive Fläche dargestellt (siehe Legende oben)."
            )
            ax.text(
                margin_in / fig_w, content_bottom_in / fig_h,
                zeile2,
                transform=ax.transAxes, fontsize=7.6, color=_TEXT_DIM, va="bottom", ha="left",
            )

        pdf.savefig(fig)
        plt.close(fig)

    ergebnis = {
        "zusammenfassung": zusammenfassung,
        "rendite": rendite,
        "profil_mit": profil_mit,
        "peak_tabelle": peak_tabelle,
        "nur_batterie": nur_batterie,
        "kumulierte_ertraege": kumulierte_ertraege,
    }
    # monatstabelle/monatstabelle_ohne/profil_ohne/peak_tabelle_ohne existieren
    # nur im Nicht-nur_batterie-Zweig (siehe oben) -- bei einem reinen
    # Batterie-System gibt es statt dessen laden_entladen_tabelle.
    if nur_batterie:
        ergebnis["laden_entladen_tabelle"] = laden_entladen_tabelle
    else:
        ergebnis["monatstabelle"] = monatstabelle
        ergebnis["monatstabelle_ohne"] = monatstabelle_ohne
        ergebnis["profil_ohne"] = profil_ohne
        ergebnis["peak_tabelle_ohne"] = peak_tabelle_ohne
    return ergebnis


def _pick_file_dialog() -> str:
    """Kleiner Datei-Auswahl-Dialog, analog zu run_battery_analysis.py.
    Bewusst hier dupliziert statt importiert, um einen Zirkelimport zu
    vermeiden (run_battery_analysis.py importiert output_create.py)."""
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError as exc:
        raise RuntimeError(
            "Kein Pfad angegeben und der Datei-Dialog (tkinter) ist in dieser "
            "Python-Installation nicht verfuegbar. Bitte den Pfad direkt als "
            "Argument angeben:\n"
            "  python output_create.py <Pfad-zu-input_lp.xlsx>"
        ) from exc

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askopenfilename(
        title="input_lp.xlsx auswaehlen",
        filetypes=[("Excel-Dateien", "*.xlsx"), ("Alle Dateien", "*.*")],
    )
    root.destroy()
    if not path:
        raise SystemExit("Keine Datei ausgewaehlt -- Abbruch.")
    return path


def main(argv=None):
    """Standalone-Einstiegspunkt: optimiert direkt ein vorhandenes
    input_lp.xlsx (ohne input_prep.py neu laufen zu lassen) und schreibt
    ergebnis.xlsx in denselben Ordner. Siehe Modul-Docstring oben."""
    # Lazy-Import, damit ein blosses `import output_create` (z.B. aus
    # run_battery_analysis.py) nicht sofort oemof.solph/pyomo braucht.
    import battery_optimization

    parser = argparse.ArgumentParser(
        description="Optimierung + Output-Kontrolle direkt aus einem "
                     "vorhandenen input_lp.xlsx erzeugen (ohne input_prep.py)."
    )
    parser.add_argument(
        "input_lp_file", nargs="?", default=None,
        help="Pfad zur input_lp.xlsx. Falls weggelassen: Datei-Auswahl-Dialog.",
    )
    parser.add_argument(
        "--solver", default="cbc",
        help="Solver (Default: cbc, siehe SOLVER-Setup in battery_optimization.py).",
    )
    args = parser.parse_args(argv)

    input_lp_path = args.input_lp_file or _pick_file_dialog()
    input_lp_path = os.path.abspath(input_lp_path)
    if not os.path.isfile(input_lp_path):
        raise SystemExit(f"Datei nicht gefunden: {input_lp_path}")

    data_dir = os.path.dirname(input_lp_path)
    result_path = os.path.join(data_dir, derive_result_filename(input_lp_path))
    pdf_path = os.path.join(data_dir, derive_pdf_filename(input_lp_path))

    print(f"Optimiere direkt auf Basis von {input_lp_path} (ohne input_prep.py) ...")
    module, ts, anteile, zeitreihen, params, pv_profile, last_profile = (
        battery_optimization.main(input_path=input_lp_path, solver=args.solver)
    )

    print("Schreibe Output-Kontrolle ...")
    save_output_excel(
        result_path, module, ts, anteile, zeitreihen, params, pv_profile, last_profile
    )

    print("Erstelle PDF-Report ...")
    # entladeenergie_kwh mitgeben, damit die Degradationskosten (siehe
    # battery_optimization.DEGRADATIONSKOSTEN_CHF_PRO_KWH) auch im PDF-Report
    # als eigene Ausgaben-Zeile auftauchen (siehe capital_costs()-Docstring).
    entladeenergie_kwh = float(ts["dis"].sum() * battery_optimization.DT_HOURS)
    kapitalkosten = battery_optimization.capital_costs(params, entladeenergie_kwh=entladeenergie_kwh)
    build_pdf_report(
        pdf_path, module, ts, anteile, zeitreihen, params, pv_profile, last_profile, kapitalkosten
    )

    print(f"\nFertig. Ergebnis gespeichert unter: {result_path}")
    print(f"PDF-Report gespeichert unter: {pdf_path}")


if __name__ == "__main__":
    main()