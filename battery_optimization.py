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
(mit oemof.solph 0.6.4 + einem Solver, z.B. `cbc` oder `glpk`) laufen lassen
und mir Fehlermeldungen zurueckmelden -- dann kann ich gezielt nachbessern.
Kandidaten fuer API-Abweichungen je nach genauer oemof-Version, auf die beim
Testen besonders zu achten ist (Kommentare an den jeweiligen Stellen im Code):
  - `Flow(custom_attributes={...})` fuer generic_integral_limit
  - `om.objective.expr += ...` zum Erweitern der Zielfunktion
  - exakte Attributnamen fuer Ergebnis-Zugriff (`processing.results`)

Modellumfang
------------
- PV (Source, fixes Profil) + Batterie (GenericStorage) + Netz, mit:
  - PV-Profil je nach Parameter!C17: "normiert" -> Zeitreihen!E (CH_PV_normiert_
    1kWp) * DC Leistung (C16); "Profil_Zeitreihe" -> Zeitreihen!F (PV_Profil)
    direkt (siehe `build_pv_profile()`).
  - "Netzbezug-Bus" als gemeinsamer Netzanschlusspunkt: EIN Zufluss vom Netz
    mit hartem Limit `Maximaler Bezug` (Parameter!C35) fuer Last+Batterie
    zusammen (siehe Klaerung im Gespraech).
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
  - Last(t) je nach Parameter!C43: "Ja" -> Zeitreihen!I (Last, in kWh PRO
    15-MIN-INTERVALL -- Umrechnung in kW via `/DT_HOURS`, siehe
    `build_last_profile()`), "Nein"/leer -> Last = 0 ueberall (Modell
    funktioniert auch ganz ohne Verbrauch).
  - Start-/End-SOC = 50% (`initial_storage_level=0.5`, `balanced=True`).
  - Vollzyklen-Durchsatzbegrenzung (C29) ueber
    `oemof.solph.constraints.generic_integral_limit`.
- Peak-Shaving als custom Pyomo-Block, AUTOMATISCH aktiv/inaktiv je nach
  Parameter!C9 (Netznutzung Leistung): Wert vorhanden und != 0 -> aktiv,
  0/leer -> inaktiv (siehe `determine_peakshaving_aktiv()`). Hilfsvariable
  `peak[monat] >= Netzbezug-Flow(t)`, Kostenterm in Zielfunktion.
- SRL-Vermarktung, Variante AUTOMATISCH aus Parameter!C38 abgeleitet
  (siehe `determine_srl_variante()`):
  - C38 == "Ja_Residual" -> "post_hoc": NICHT Teil der Optimierung. Nach der
    Loesung wird die freie Lade-/Entladeleistung mit den SRL-Preisen bewertet
    (nur wenn SOC_MIN_SRL < SOC/Kapazitaet < SOC_MAX_SRL, beide waehlbar).
  - C38 == "Ja_Optimiert" -> "optimiert": Teil der Optimierung (custom
    Pyomo-Block, NUR Leistungsheadroom, keine SOC-/Energie-Nachhaltigkeitskopplung).
  - C38 == "Nein" -> "keine": kein SRL-Modul.
- Ertragsaufteilung nach der Optimierung in die Module Eigenverbrauch,
  Arbitrage, Anfangsbestand-Verwertung, Peak-Shaving, SRL -- ueber eine
  "gut durchmischter Tank"-Konvention (PV-/Netz-/Anfangsbestand-Anteil des
  Speicherinhalts wird pro Zeitschritt proportional fortgeschrieben).

Benoetigte Pakete: oemof.solph, pyomo, pandas, numpy, openpyxl, ein Solver
(z.B. `pip install pyomo oemof.solph pandas numpy openpyxl` + CBC/GLPK
systemweit installiert).
"""

from __future__ import annotations

import datetime as dt
import os

import numpy as np
import pandas as pd
import pyomo.environ as po
from openpyxl import load_workbook

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
SOLVER = "cbc"  # Alternativen: "glpk", "gurobi", "cplex" ...

# 15-Minuten-Zeitschritte -- an mehreren Stellen gebraucht (Last-Umrechnung
# kWh->kW, SRL-Erloesberechnung, Ertragsaufteilung), daher zentral definiert.
DT_HOURS = 0.25

# Peak-Shaving und SRL-Variante werden standardmaessig automatisch aus dem
# Excel abgeleitet (Parameter!C9 bzw. C38 -- siehe determine_peakshaving_aktiv()
# und determine_srl_variante() unten). Diese beiden Konstanten sind nur noch
# Fallback-Werte, falls man `main()` mit einer expliziten Zeitreihe ohne Excel-
# Bezug aufruft; im Normalfall braucht ihr sie nicht anzufassen.
SOC_MIN_SRL = 0.10          # nur fuer SRL-Variante "post_hoc", waehlbar
SOC_MAX_SRL = 0.90          # nur fuer SRL-Variante "post_hoc", waehlbar

# Siehe input_prep.py: Excel zaehlt Zeit als fixen UTC+1-Offset,
# keine Sommerzeit-Umstellung.
CET_FIXED = dt.timezone(dt.timedelta(hours=1))

# Spaltenlayout im Sheet "Zeitreihen" (wie in input_prep.py)
COL_ZEIT = 1
COL_SWISSIX = 2
COL_RUECKLIEFER = 3
COL_BEZUG = 4
COL_PV_NORMIERT = 5
COL_PV_PROFIL = 6
COL_SRL_NEG = 7
COL_SRL_POS = 8
COL_LAST = 9  # neu: Last/Verbrauch in kWh pro 15-Min-Intervall (nicht kW!)


# --------------------------------------------------------------------------
# 1. Input einlesen
# --------------------------------------------------------------------------

def read_inputs(path: str = INPUT_PATH):
    wb = load_workbook(path, data_only=True)
    wsP = wb["Parameter"]
    wsZ = wb["Zeitreihen"]

    params = {
        "beginn": wsP["C2"].value,
        "ende": wsP["C3"].value,
        "rueckliefer_flat": wsP["C8"].value,
        "dc_leistung": wsP["C16"].value,
        "produktions_input": wsP["C17"].value,
        "kapazitaet": wsP["C20"].value,
        "leistung": wsP["C21"].value,
        "trafo_wirkungsgrad": wsP["C22"].value,
        "lade_wirkungsgrad": wsP["C23"].value,
        "entlade_wirkungsgrad": wsP["C24"].value,
        "invest_kosten_kwh": wsP["C25"].value,
        "unterhalt_kosten": wsP["C26"].value,
        "lebenszeit": wsP["C27"].value,
        "vollzyklen": wsP["C28"].value,
        "ladegrenze_soc": wsP["C30"].value / 100.0,
        "entladegrenze_soc": wsP["C31"].value / 100.0,
        "max_einspeisung": wsP["C34"].value,
        "max_bezug": wsP["C35"].value,
        "netznutzung_leistung": wsP["C9"].value,
        "srl_teilnahme": wsP["C38"].value,
        "last_vorhanden": wsP["C43"].value,  # neu: "Ja" -> Zeitreihen!I, "Nein" -> keine Last
    }

    # Durchsatzbegrenzung/Jahr (C29 ist im Original eine Formel; falls der
    # gecachte Wert beim Excel-Resave verlorengegangen ist, hier neu herleiten)
    durchsatz = wsP["C29"].value
    if not isinstance(durchsatz, (int, float)):
        durchsatz = params["kapazitaet"] * params["vollzyklen"] / params["lebenszeit"]
    params["durchsatz_mwh_jahr"] = durchsatz

    n_rows = wsZ.max_row
    zeit = [wsZ.cell(row=r, column=COL_ZEIT).value for r in range(2, n_rows + 1)]
    index = pd.DatetimeIndex(zeit).tz_localize(CET_FIXED)

    def col(c):
        return pd.Series(
            [wsZ.cell(row=r, column=c).value for r in range(2, n_rows + 1)],
            index=index, dtype=float,
        )

    zeitreihen = pd.DataFrame(
        {
            "swissix": col(COL_SWISSIX),
            "rueckliefertarif": col(COL_RUECKLIEFER),
            "bezugstarif": col(COL_BEZUG),
            "pv_normiert": col(COL_PV_NORMIERT),
            "pv_profil": col(COL_PV_PROFIL),
            "srl_neg": col(COL_SRL_NEG),
            "srl_pos": col(COL_SRL_POS),
            "last_kwh": col(COL_LAST),  # Achtung: kWh pro 15-Min-Intervall, NICHT kW
        }
    )
    return params, zeitreihen


def build_pv_profile(params: dict, zeitreihen: pd.DataFrame) -> pd.Series:
    """kW-Zeitreihe der PV-Erzeugung, je nach Parameter!C17.

    Gross-/Kleinschreibung und Leerzeichen am Rand werden ignoriert (siehe
    determine_srl_variante() fuer denselben Grund -- Excel-Eintraege variieren
    da leicht)."""
    wert = params["produktions_input"]
    normalisiert = str(wert).strip().lower() if wert is not None else ""

    if normalisiert == "normiert":
        profile = zeitreihen["pv_normiert"]
        if profile.isna().all():
            raise RuntimeError(
                "Parameter!C17='normiert', aber Zeitreihen!E (CH_PV_normiert_1kWp) "
                "ist leer. Bitte zuerst das PV-Referenzprofil ergaenzen (siehe "
                "input_prep.py::fetch_pv_reference_profile -- braucht "
                "Standort/Ausrichtung)."
            )
        return profile * params["dc_leistung"]
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
    """Leitet die SRL-Variante aus Parameter!C38 ab (offizielle Werte im
    Hilfen_Listen-Dropdown: 'Ja_Residual', 'Ja_Optimiert', 'Nein'):
    'Ja_Residual' -> 'post_hoc', 'Ja_Optimiert' -> 'optimiert', 'Nein' -> 'keine'.
    Gross-/Kleinschreibung und Leerzeichen am Rand werden zusaetzlich ignoriert
    (falls doch mal jemand manuell 'ja_residual' o.ae. eintippt)."""
    mapping = {"ja_residual": "post_hoc", "ja_optimiert": "optimiert", "nein": "keine"}
    wert = params["srl_teilnahme"]
    normalisiert = str(wert).strip().lower() if wert is not None else ""
    if normalisiert not in mapping:
        raise ValueError(
            f"Parameter!C38 = {wert!r} ist kein bekannter Wert. Erwartet (Gross-/"
            f"Kleinschreibung egal) einen von 'Ja_Residual' (-> post_hoc-Bewertung), "
            f"'Ja_Optimiert' (-> Teil der Optimierung), 'Nein' (-> keine SRL-Vermarktung)."
        )
    return mapping[normalisiert]


def build_last_profile(
    params: dict, zeitreihen: pd.DataFrame, last_series: pd.Series | None = None
) -> pd.Series:
    """kW-Zeitreihe der Last, je nach Parameter!C43.

    - `last_series` explizit uebergeben: nimmt Vorrang (z.B. fuer Tests), muss
      bereits in kW vorliegen.
    - Parameter!C43 == "Ja": Zeitreihen!I (Last) verwenden. WICHTIG: diese
      Spalte ist in kWh PRO 15-MIN-INTERVALL angegeben, nicht in kW -- oemof-
      Flows sind aber Leistungen (kW). Umrechnung: kW = kWh_intervall / DT_HOURS.
    - Parameter!C43 == "Nein" (oder leer): Last = 0 ueberall, das Modell
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
                "Parameter!C43='Ja', aber Zeitreihen!I (Last) ist leer. Bitte "
                "Lastreihe ergaenzen oder C43 auf 'Nein' setzen."
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
    b_ac = solph.Bus(label="b_ac")  # Last/Export-Bus

    # --- Quellen -------------------------------------------------------------
    src_pv = solph.components.Source(
        label="pv",
        outputs={b_pv: solph.Flow(nominal_value=1, fix=pv_profile.values)},
    )

    src_netz = solph.components.Source(
        label="netz",
        outputs={
            b_netzbezug: solph.Flow(
                nominal_value=params["max_bezug"],
                variable_costs=zeitreihen["bezugstarif"].values,
            )
        },
    )

    # --- Senken --------------------------------------------------------------
    snk_last = solph.components.Sink(
        label="last",
        inputs={b_ac: solph.Flow(nominal_value=1, fix=last_profile.values)},
    )

    snk_export = solph.components.Sink(
        label="export",
        inputs={
            b_ac: solph.Flow(
                nominal_value=params["max_einspeisung"],
                variable_costs=-zeitreihen["rueckliefertarif"].values,
            )
        },
    )

    # --- Direktflüsse (unbegrenzt, ausser durch die Busse/Quellen selbst) ---
    # PV -> Hausbus (direkter Verbrauch/Export) und PV -> Ladebus (Batterieladung)
    # ueber je einen 1:1-Converter, damit b_pv/b_netzbezug sich sauber auf die
    # zwei Abgaenge (Last-Deckung vs. Batterieladung) aufteilen lassen.
    link_pv_ac = solph.components.Converter(
        label="pv_zu_ac",
        inputs={b_pv: solph.Flow()},
        outputs={b_ac: solph.Flow(nominal_value=params["dc_leistung"])},
        conversion_factors={b_ac: 1.0},
    )
    link_pv_charge = solph.components.Converter(
        label="pv_zu_ladebus",
        inputs={b_pv: solph.Flow()},
        outputs={b_ladebus: solph.Flow(nominal_value=params["leistung"])},
        conversion_factors={b_ladebus: 1.0},
    )
    link_netz_ac = solph.components.Converter(
        label="netzbezug_zu_ac",
        inputs={b_netzbezug: solph.Flow()},
        outputs={b_ac: solph.Flow(nominal_value=params["max_bezug"])},
        conversion_factors={b_ac: 1.0},
    )
    link_netz_charge = solph.components.Converter(
        label="netzbezug_zu_ladebus",
        inputs={b_netzbezug: solph.Flow()},
        outputs={b_ladebus: solph.Flow(nominal_value=params["leistung"])},
        conversion_factors={b_ladebus: 1.0},
    )
    link_entlade_ac = solph.components.Converter(
        label="entladebus_zu_ac",
        inputs={b_entladebus: solph.Flow()},
        outputs={b_ac: solph.Flow(nominal_value=params["leistung"])},
        conversion_factors={b_ac: 1.0},
    )

    # --- Batterie --------------------------------------------------------------
    # Trafo-Wirkungsgrad (C22) vereinfachend mit Lade-/Entladewirkungsgrad
    # verrechnet (siehe Modulkommentar oben).
    eta_laden = params["lade_wirkungsgrad"] * params["trafo_wirkungsgrad"]
    eta_entladen = params["entlade_wirkungsgrad"] * params["trafo_wirkungsgrad"]

    # custom_attributes fuer generic_integral_limit (Vollzyklen-Durchsatz) --
    # Attributname je nach oemof-Version ggf. anzupassen (siehe Hinweis oben).
    discharge_flow = solph.Flow(
        nominal_value=params["leistung"],
        custom_attributes={"vollzyklen_keyword": 1},
    )

    storage = solph.components.GenericStorage(
        label="batterie",
        nominal_storage_capacity=params["kapazitaet"],
        inputs={b_ladebus: solph.Flow(nominal_value=params["leistung"])},
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
        b_pv, b_netzbezug, b_ladebus, b_entladebus, b_ac,
        src_pv, src_netz, snk_last, snk_export,
        link_pv_ac, link_pv_charge, link_netz_ac, link_netz_charge, link_entlade_ac,
        storage,
    )

    om = solph.Model(es)

    node_refs = {
        "b_pv": b_pv, "b_netzbezug": b_netzbezug, "b_ladebus": b_ladebus,
        "b_entladebus": b_entladebus, "b_ac": b_ac,
        "src_netz": src_netz, "storage": storage,
        "link_pv_charge": link_pv_charge, "link_netz_charge": link_netz_charge,
        "link_entlade_ac": link_entlade_ac,
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
    # Zielfunktion erweitern -- siehe Hinweis im Modulkopf zu API-Unsicherheit.
    om.objective.expr += peak_kosten

    return block


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
    om.objective.expr += -erloes  # Erloes = negative Kosten

    return block


# --------------------------------------------------------------------------
# 5. Vollzyklen-Durchsatzbegrenzung (nativ ueber oemof.solph.constraints)
# --------------------------------------------------------------------------

def add_vollzyklen_limit(om, storage, entladebus, durchsatz_mwh_jahr: float):
    flows = {(storage, entladebus): storage.outputs[entladebus]}
    solph.constraints.generic_integral_limit(
        om, "vollzyklen_keyword", flows, upper_limit=durchsatz_mwh_jahr * 1000.0
    )


# --------------------------------------------------------------------------
# 6. Loesen & Ergebnisse extrahieren
# --------------------------------------------------------------------------

def solve(om, solver: str = SOLVER):
    om.solve(solver=solver)
    return solph.processing.results(om)


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

    out = pd.DataFrame(
        {
            "soc": soc,
            "ch_pv": ch_pv,
            "ch_grid": ch_grid,
            "dis": dis,
            "netzbezug_total": netzbezug_total,
        }
    )
    return out


# --------------------------------------------------------------------------
# 7. Modul-Ertragsaufteilung (Tank-Konvention)
# --------------------------------------------------------------------------

def allocate_modules(ts: pd.DataFrame, zeitreihen: pd.DataFrame, params: dict,
                      dt_hours: float, srl_variante: str,
                      soc_min_srl: float = SOC_MIN_SRL, soc_max_srl: float = SOC_MAX_SRL):
    kapazitaet = params["kapazitaet"]
    eta_laden = params["lade_wirkungsgrad"] * params["trafo_wirkungsgrad"]

    n = len(ts)
    pv_anteil = np.zeros(n + 1)
    grid_anteil = np.zeros(n + 1)
    anfang_anteil = np.zeros(n + 1)
    anfang_anteil[0] = 1.0  # Start-SOC (50%) ist zu 100% "Anfangsbestand"

    soc = ts["soc"].values
    soc0 = kapazitaet * 0.5
    soc_full = np.concatenate([[soc0], soc])

    dis_pv = np.zeros(n)
    dis_grid = np.zeros(n)
    dis_anfang = np.zeros(n)

    for i in range(n):
        soc_before = soc_full[i]
        soc_after = soc_full[i + 1]
        ch_pv_i = ts["ch_pv"].values[i] * eta_laden
        ch_grid_i = ts["ch_grid"].values[i] * eta_laden
        dis_i = ts["dis"].values[i]  # bereits nach Entladewirkungsgrad (oemof-Output)

        pv_energy_before = pv_anteil[i] * soc_before
        grid_energy_before = grid_anteil[i] * soc_before
        anfang_energy_before = anfang_anteil[i] * soc_before

        dis_pv[i] = dis_i * pv_anteil[i]
        dis_grid[i] = dis_i * grid_anteil[i]
        dis_anfang[i] = dis_i * anfang_anteil[i]

        pv_energy_after = pv_energy_before + ch_pv_i - dis_pv[i]
        grid_energy_after = grid_energy_before + ch_grid_i - dis_grid[i]
        anfang_energy_after = anfang_energy_before - dis_anfang[i]

        if soc_after > 1e-9:
            pv_anteil[i + 1] = max(0.0, pv_energy_after / soc_after)
            grid_anteil[i + 1] = max(0.0, grid_energy_after / soc_after)
            anfang_anteil[i + 1] = max(0.0, anfang_energy_after / soc_after)
        else:
            pv_anteil[i + 1] = grid_anteil[i + 1] = anfang_anteil[i + 1] = 0.0

    bezugstarif = zeitreihen["bezugstarif"].values
    rueckliefertarif = zeitreihen["rueckliefertarif"].values
    netzbezug = ts["netzbezug_total"].values
    # Bewertungspreis: vermiedener Bezug, falls die Anlage zu diesem Zeitpunkt
    # sonst Netz bezogen haette, sonst erzielter Einspeiseerlös.
    bewertungspreis = np.where(netzbezug > 1e-9, bezugstarif, rueckliefertarif)

    eigenverbrauch_ertrag = float(
        np.sum(dis_pv * bewertungspreis) - np.sum(ts["ch_pv"].values * rueckliefertarif)
    )
    arbitrage_ertrag = float(
        np.sum(dis_grid * bewertungspreis) - np.sum(ts["ch_grid"].values * bezugstarif)
    )
    anfangsbestand_ertrag = float(np.sum(dis_anfang * bewertungspreis))

    module = {
        "Eigenverbrauchsoptimierung": eigenverbrauch_ertrag,
        "Arbitrage der Batterie": arbitrage_ertrag,
        "Anfangsbestand-Verwertung": anfangsbestand_ertrag,
    }

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

        srl_mehrertrag = float(
            np.sum(residual_pos * srl_pos_preis) * dt_hours
            + np.sum(residual_neg * srl_neg_preis) * dt_hours
        )
        module["SRL (Post-hoc-Mehrertrag)"] = srl_mehrertrag

    return module, pd.DataFrame(
        {"pv_anteil": pv_anteil[1:], "grid_anteil": grid_anteil[1:], "anfang_anteil": anfang_anteil[1:]},
        index=ts.index,
    )


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
):
    """
    peakshaving_aktiv=None (Default): automatisch aus Parameter!C9 abgeleitet
        (aktiv, falls C9 != 0/NaN). Explizit True/False uebersteuert das Excel.
    srl_variante=None (Default): automatisch aus Parameter!C38 abgeleitet
        ('Ja_Residual' -> 'post_hoc', 'Ja_Optimiert' -> 'optimiert',
        'Nein' -> 'keine'). Explizite Angabe uebersteuert das Excel.
    """
    print(f"Lese {input_path} ...")
    params, zeitreihen = read_inputs(input_path)

    if peakshaving_aktiv is None:
        peakshaving_aktiv = determine_peakshaving_aktiv(params)
        print(f"Peak-Shaving automatisch aus Parameter!C9 abgeleitet: {peakshaving_aktiv}")

    if srl_variante is None:
        srl_variante = determine_srl_variante(params)
        print(f"SRL-Variante automatisch aus Parameter!C38 abgeleitet: {srl_variante!r}")

    pv_profile = build_pv_profile(params, zeitreihen)
    last_profile = build_last_profile(params, zeitreihen, last_series)

    print("Baue Energiesystem ...")
    es, om, node_refs, dt_hours = build_energysystem(params, zeitreihen, pv_profile, last_profile)

    add_vollzyklen_limit(om, node_refs["storage"], node_refs["b_entladebus"], params["durchsatz_mwh_jahr"])

    if peakshaving_aktiv:
        print("Aktiviere Peak-Shaving ...")
        add_peak_shaving(
            om, node_refs["src_netz"], node_refs["b_netzbezug"],
            params["netznutzung_leistung"], zeitreihen.index,
        )

    if srl_variante == "optimiert":
        print("Aktiviere SRL als Teil der Optimierung (einfache Variante) ...")
        add_srl_optimiert(
            om, node_refs["storage"], node_refs["b_ladebus"], node_refs["b_entladebus"],
            params["leistung"], zeitreihen["srl_pos"].values, zeitreihen["srl_neg"].values,
            dt_hours,
        )

    print(f"Loese mit Solver '{solver}' ...")
    results = solve(om, solver)

    print("Werte Ergebnisse aus ...")
    ts = extract_timeseries(results, node_refs, zeitreihen)
    module, anteile = allocate_modules(
        ts, zeitreihen, params, dt_hours, srl_variante, soc_min_srl, soc_max_srl
    )

    if peakshaving_aktiv:
        peak_block = om.PeakShaving
        peak_kosten = sum(
            po.value(peak_block.peak[m]) * params["netznutzung_leistung"] for m in peak_block.MONATE
        )
        module["Peak-Shaving"] = -peak_kosten  # als Kosten (negativer Ertrag)

    if srl_variante == "optimiert":
        srl_block = om.SRLReserve
        srl_erloes = sum(
            (po.value(srl_block.r_pos[t]) * zeitreihen["srl_pos"].values[t]
             + po.value(srl_block.r_neg[t]) * zeitreihen["srl_neg"].values[t]) * dt_hours
            for t in om.TIMESTEPS
        )
        module["SRL (Teil der Optimierung)"] = srl_erloes

    total = sum(module.values())
    module["Total (operativ)"] = total

    print("\n=== Ertrag pro Modul (CHF/Jahr) ===")
    for k, v in module.items():
        print(f"  {k:45s} {v:>12,.0f}")

    # --- Kapitalkosten/Unterhalt (informativ, nicht Teil der LP-Zielfunktion) --
    wacc = 0.03  # Annahme, bitte pruefen/anpassen
    capex = params["kapazitaet"] * params["invest_kosten_kwh"]
    lebenszeit = params["lebenszeit"]
    annuitaet = capex * (wacc * (1 + wacc) ** lebenszeit) / ((1 + wacc) ** lebenszeit - 1)
    unterhalt = params["kapazitaet"] * params["unterhalt_kosten"]

    print("\n=== Kapitalkosten/Unterhalt (informativ, Annahme WACC=3%) ===")
    print(f"  {'Investitions-Annuitaet':45s} {-annuitaet:>12,.0f}")
    print(f"  {'Unterhaltkosten':45s} {-unterhalt:>12,.0f}")
    print(f"  {'Netto-Ergebnis (operativ - Kapital/Unterhalt)':45s} {total - annuitaet - unterhalt:>12,.0f}")

    return module, ts, anteile


if __name__ == "__main__":
    main()