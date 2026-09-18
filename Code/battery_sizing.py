"""
optimale_batteriegroesse.py
============================

NEU (Beat, 15.9.2026: "wenn ich die Konfiguration habe (PV+Last) zum
Beispiel, koennten wir dann eine Optimierung darum bauen um die 'optimale
Batterie' dazu zu finden (Leistung und Kapazitaet)? Mir schwebt ein neues
Skript vor, das ein Inputfile nimmt und als Output so eine Heatmap erstellt
mit Kombinationen aus Kapazitaet und Leistung fuer die optimale
Batteriegroesse." -- NEU GESCHRIEBEN 18.9.2026: die urspruengliche Version
ging mit einem Container-Reset einer frueheren Session verloren; Beat hatte
seine lokale Kopie ebenfalls nicht mehr. Diese Version ist 1:1 aus der im
Projekt-Verlauf dokumentierten Spezifikation rekonstruiert.

WAS DIESES SKRIPT TUT
----------------------
Nimmt eine normale Inputs.xlsx (gleiches Format wie run_battery_analysis.py),
bereitet sie EINMAL ueber input_prep.py auf (Tarife/SwissIX/PV-Referenzprofil
haengen nicht von der Batteriegroesse ab -- nur EIN input_prep-Lauf noetig,
nicht einer pro Rasterpunkt) und rechnet dann fuer ein RASTER aus
Kapazitaet x C-Rate je eine VOLLSTAENDIGE oemof.solph-Optimierung ueber
`battery_optimization.main(params_override=...)`. Dieser Hook existiert im
aktuellen battery_optimization.py bereits genau fuer diesen Zweck (siehe
dessen main()-Docstring: "NEU (15.9.2026, fuer optimale_batteriegroesse.py)").

Ergebnis:
  - Ein Excel mit der vollen Ergebnistabelle (eine Zeile pro Rasterpunkt):
    Kapazitaet, C-Rate, Leistung, Capex, Unterhalt, Annuitaet, Degradation,
    operativer Gesamtertrag, Gewinn/Verlust netto, Kapitalverzinsung,
    Entladeenergie -- oder ein Fehlertext, falls ein Rasterpunkt nicht loesbar
    war (z.B. Infeasibility bei einer sehr kleinen Batterie).
  - Ein PDF mit drei Heatmaps (Gewinn/Verlust netto, durchschnittliche
    Kapitalverzinsung, operativer Gesamtertrag) -- Zeilen = Kapazitaet (kWh),
    Spalten = C-Rate, jede Zelle zusaetzlich mit der resultierenden Leistung
    (kW) beschriftet. Die Zelle mit dem hoechsten Gewinn/Verlust wird in allen
    drei Heatmaps schwarz umrandet hervorgehoben (dieselbe Position in jeder
    der drei Heatmaps, damit man sie direkt vergleichen kann).

RASTER-BESTIMMUNG (komplett automatisch aus dem Inputfile abgeleitet, Beats
ausdruecklicher Wunsch: "ich moechte eigentlich keine Eingabe machen")
--------------------------------------------------------------------------
  - LEISTUNG wird nicht unabhaengig variiert, sondern ueber eine C-Rate:
    Leistung[kW] = C-Rate * Kapazitaet[kWh]. Default drei C-Raten: 0.25,
    0.5, 1.0 (Batterie-Fachbegriff: "wie schnell relativ zur Kapazitaet
    geladen/entladen werden kann") -- siehe C_RATEN unten, aenderbar.
  - KAPAZITAET wird aus dem PV-/Lastprofil abgeleitet (siehe
    bestimme_kapazitaets_basis()): primaer die typische TAGESENERGIE des
    PV-UEBERSCHUSSES (PV minus Last, positiver Teil) -- die Energiemenge, die
    an einem Durchschnittstag ueberhaupt sinnvoll zwischengespeichert werden
    koennte. Kein nennenswerter PV-Ueberschuss (PV < Last oder keine
    PV-Anlage) -> typische taegliche LAST-Energie (Batterie zur
    Lastspitzenkappung/Arbitrage). Auch keine Last vorhanden (reines
    Batterie-Arbitrage-/SRL-System ohne PV und ohne Last) -> nichts aus dem
    Profil ableitbar, dann greift ein fester Default-Bereich
    (KAPAZITAETEN_FALLBACK_KWH) mit einer klaren Konsolen-Warnung. Aus der
    Tagesenergie-Basis werden 6 Kapazitaets-Werte gebildet (0.5x/0.75x/1x/
    1.5x/2x/3x der Basis, auf "schoene" Zehner-/Fuenfziger-/Hunderter-
    Schritte gerundet, siehe _rund_schoen()) -- macht zusammen 6 Kapazitaeten
    x 3 C-Raten = 18 Rasterpunkte im Default (alles ueber die Konstanten am
    Skriptanfang aenderbar).

UNTERHALTKOSTEN-ANNAHME (bitte mit Beat pruefen)
--------------------------------------------------
`battery_optimization.capital_costs()` liest die Unterhaltkosten als FIXEN
CHF/Jahr-Wert direkt aus Parameter!C26, unabhaengig von der Kapazitaet. Bei
einem Kapazitaets-Sweep waere ein fixer Wert fuer jede Groesse irrefuehrend
(eine 10x groessere Batterie haette dieselben Unterhaltkosten wie eine
winzige). Dieses Skript skaliert die Unterhaltkosten deshalb STANDARDMAESSIG
PROPORTIONAL zur Investitionssumme (derselbe Prozentsatz vom Capex wie im
Original-Excel-Wert, siehe sweep()) -- steuerbar ueber
UNTERHALT_PROPORTIONAL_SKALIEREN (Default True) unten.

TEIL-JAHR-SZENARIEN (z.B. "..._Q2.xlsx", nur 3 Monate)
---------------------------------------------------------
`_kombination_rechnen()` rechnet die IST-Einnahmen/Entladeenergie/
Degradation jedes Rasterpunkts auf ein volles Jahr hoch (geteilt durch die
tatsaechlich simulierte Dauer in Jahren), bevor sie gegen die -- ohnehin
bereits jaehrlichen -- Annuitaet/Unterhalt-Kosten aus capital_costs()
gerechnet werden. Exakt dieselbe Logik wie in output_create.py's
build_rendite_kennzahlen()/build_income_expense_summary() (dort aus Beats
Frage "wieso sagt der Report eine Jahresbetrachtung, obwohl es nur 3 Monate
sind" entstanden) -- ohne diese Hochrechnung waeren Gewinn/Verlust und
Kapitalverzinsung bei einem Teil-Jahr-Input systematisch zu tief (z.B. bei
einem 3-Monats-Szenario um den Faktor ~4).

LAUFZEIT-HINWEIS
------------------
Jeder Rasterpunkt ist ein VOLLSTAENDIGER oemof.solph-Solve ueber das ganze
Jahr (15-Minuten-Schritte) -- bei 18 Rasterpunkten sind das 18 volle
Optimierungslaeufe NACHEINANDER, je nach Rechner von wenigen Minuten bis
mehreren Stunden. Fuer einen ersten schnellen Test die Konstanten
KAPAZITAETS_MULTIPLIKATOREN/C_RATEN unten auf weniger Werte kuerzen (z.B.
nur `C_RATEN = [1.0]` und `KAPAZITAETS_MULTIPLIKATOREN = [1.0]` fuer einen
einzelnen Testlauf).

TESTABDECKUNG / SCHLANKE ABHAENGIGKEITEN
-------------------------------------------
Wie bei output_create.py bewusst KEIN Import von `battery_optimization`/
`input_prep` auf Modulebene (siehe deren `import oemof.solph`, das in dieser
Cloud-Sandbox nicht installierbar ist) -- die reine Raster-/Rundungs-/
Heatmap-/Excel-Logik (`_rund_schoen()`, `_tagesenergie_typisch()`,
`bestimme_kapazitaets_basis()`, `baue_raster()`, `baue_heatmap_pdf()`,
`speichere_ergebnis_excel()`) haengt NUR von numpy/pandas/matplotlib/
output_create ab und ist dadurch in dieser Sandbox isoliert mit synthetischen
Daten testbar (alle drei Zweige von bestimme_kapazitaets_basis(): PV-
Ueberschuss, Last-Fallback, kein Bezug/Default-Werte). `battery_optimization`
und `input_prep` werden NUR innerhalb von `_kombination_rechnen()`, `sweep()`
und `main()` lazy importiert (genau wie in output_create.py's eigenem
Standalone-`main()`). NICHT TESTBAR IN DIESER SANDBOX (kein lauffaehiges
oemof.solph/pyomo/CBC): der eigentliche Sweep-Aufruf selbst. Bitte lokal
testen, ggf. zuerst mit gekuerztem Raster (siehe oben).
"""

from __future__ import annotations

import argparse
import os
import warnings

import matplotlib

matplotlib.use("Agg")
import matplotlib.colors as mcolors
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import numpy as np
import pandas as pd

import output_create

# NEU (siehe input_prep.py/battery_optimization.py): dieselbe gezielte
# Unterdrueckung der harmlosen openpyxl-"Data Validation"-Warnung, die sonst
# auch hier bei jedem input_prep/battery_optimization-Aufruf erscheinen wuerde.
warnings.filterwarnings(
    "ignore", message=r".*Data Validation extension is not supported.*", category=UserWarning
)

# --------------------------------------------------------------------------
# Konfiguration (Beats Wunsch: "ich moechte eigentlich keine Eingabe machen")
# --------------------------------------------------------------------------
C_RATEN = [0.25, 0.5, 1.0]
KAPAZITAETS_MULTIPLIKATOREN = [0.5, 0.75, 1.0, 1.5, 2.0, 3.0]
KAPAZITAETEN_FALLBACK_KWH = [100.0, 200.0, 300.0, 500.0, 750.0, 1000.0]
UNTERHALT_PROPORTIONAL_SKALIEREN = True
WACC = 0.03  # gleiche Default-Annahme wie battery_optimization.capital_costs()

# Fleco-Markenfarben (dieselben wie in output_create.py) fuer die Heatmaps.
_POS = "#196B24"
_NEG = "#B23B3B"
_TEXT = "#1A1A1A"


# --------------------------------------------------------------------------
# 1. Raster-Bestimmung (reine Logik, kein oemof noetig)
# --------------------------------------------------------------------------

def _rund_schoen(x: float) -> float:
    """Rundet auf eine 'schoene' Schrittweite, abhaengig von der
    Groessenordnung (5er/10er/50er/100er/500er-Schritte) -- damit die
    Rasterwerte wie plausible Batteriegroessen aussehen (z.B. 150 kWh statt
    147.3 kWh)."""
    if x <= 0:
        return 0.0
    if x < 50:
        schritt = 5.0
    elif x < 200:
        schritt = 10.0
    elif x < 1000:
        schritt = 50.0
    elif x < 5000:
        schritt = 100.0
    else:
        schritt = 500.0
    return round(x / schritt) * schritt


def _tagesenergie_typisch(profil_kw: pd.Series | np.ndarray, dt_hours: float = 0.25) -> float:
    """Durchschnittliche taegliche Energie (kWh/Tag) einer kW-Zeitreihe:
    Gesamtenergie ueber den simulierten Zeitraum geteilt durch die Anzahl
    Tage (aus der Anzahl Zeitschritte hergeleitet, nicht aus Kalendertagen --
    robust auch fuer Teil-Jahr-Szenarien)."""
    werte = np.asarray(profil_kw, dtype=float)
    anzahl_tage = len(werte) * dt_hours / 24.0
    if anzahl_tage <= 0:
        return 0.0
    return float(np.nansum(werte) * dt_hours / anzahl_tage)


def bestimme_kapazitaets_basis(
    pv_profile: pd.Series | np.ndarray,
    last_profile: pd.Series | np.ndarray,
    dt_hours: float = 0.25,
) -> tuple[float, str]:
    """Leitet die 'typische' Tagesenergie-Basis (kWh) fuer den Kapazitaets-
    Raster her, in drei Stufen (siehe Modulkopf):
      1. PV-Ueberschuss (PV - Last, positiver Teil) -- die Batterie speichert
         PV-Ueberschuss fuer spaeter.
      2. Last-Energie -- kein nennenswerter PV-Ueberschuss (z.B. PV < Last
         oder keine PV-Anlage), Batterie dient v.a. Lastspitzenkappung/
         Arbitrage relativ zur Last.
      3. Fallback -- weder PV-Ueberschuss noch Last vorhanden (reines
         Arbitrage-/SRL-System) -- Basis 0.0, Aufrufer (baue_raster())
         verwendet dann den festen Default-Bereich statt Multiplikatoren.
    Gibt (basis_kwh, herkunft_text) zurueck."""
    pv = np.asarray(pv_profile, dtype=float)
    last = np.asarray(last_profile, dtype=float)
    ueberschuss = np.maximum(pv - last, 0.0)

    basis_pv = _tagesenergie_typisch(ueberschuss, dt_hours)
    if basis_pv > 1e-6:
        return basis_pv, "typische taegliche PV-Ueberschussenergie (PV - Last, positiver Teil)"

    basis_last = _tagesenergie_typisch(last, dt_hours)
    if basis_last > 1e-6:
        return basis_last, "typische taegliche Last-Energie (kein nennenswerter PV-Ueberschuss)"

    return 0.0, "kein PV-Ueberschuss und keine Last vorhanden -- Fallback-Bereich wird verwendet"


def baue_raster(
    pv_profile: pd.Series | np.ndarray,
    last_profile: pd.Series | np.ndarray,
    dt_hours: float = 0.25,
) -> tuple[list[float], str]:
    """Baut die Liste der Kapazitaets-Rasterpunkte (kWh) -- entweder aus der
    automatisch abgeleiteten Basis (x KAPAZITAETS_MULTIPLIKATOREN, 'schoen'
    gerundet, Duplikate entfernt, aufsteigend sortiert) oder, falls keine
    Basis herleitbar ist, direkt aus dem festen Fallback-Bereich
    KAPAZITAETEN_FALLBACK_KWH (keine Multiplikatoren -- der Fallback-Bereich
    deckt bereits eine sinnvolle Bandbreite ab)."""
    basis_kwh, herkunft = bestimme_kapazitaets_basis(pv_profile, last_profile, dt_hours)
    if basis_kwh <= 1e-6:
        print(f"  HINWEIS: {herkunft} -- verwende feste Default-Kapazitaeten "
              f"{KAPAZITAETEN_FALLBACK_KWH} kWh.")
        return list(KAPAZITAETEN_FALLBACK_KWH), herkunft

    print(f"  Kapazitaets-Basis: {basis_kwh:,.1f} kWh/Tag ({herkunft})")
    kapazitaeten = sorted({_rund_schoen(basis_kwh * m) for m in KAPAZITAETS_MULTIPLIKATOREN})
    kapazitaeten = [k for k in kapazitaeten if k > 0]
    print(f"  Kapazitaets-Raster: {kapazitaeten} kWh (je kombiniert mit C-Raten {C_RATEN})")
    return kapazitaeten, herkunft


# --------------------------------------------------------------------------
# 2. Ein Rasterpunkt rechnen (braucht battery_optimization -> lazy Import)
# --------------------------------------------------------------------------

def _kombination_rechnen(
    input_lp_path: str,
    kapazitaet_kwh: float,
    c_rate: float,
    solver: str,
    invest_kosten_kwh: float,
    unterhalt_pct_von_capex: float | None,
) -> dict:
    """Rechnet EINE Kapazitaet/C-Rate-Kombination komplett durch (ein voller
    oemof.solph-Solve via battery_optimization.main(params_override=...)) und
    fasst das Ergebnis in einem flachen Dict zusammen (eine Zeile der
    spaeteren Ergebnistabelle). Faengt Solver-/Feasibility-Fehler ab (z.B.
    Infeasibility bei einer sehr kleinen Batterie) und gibt dann ein Dict mit
    gesetztem `fehler` zurueck, statt den ganzen Sweep abzubrechen."""
    import battery_optimization  # lazy, siehe Modulkopf

    leistung_kw = round(c_rate * kapazitaet_kwh, 3)
    params_override = {"kapazitaet": kapazitaet_kwh, "leistung": leistung_kw}
    if unterhalt_pct_von_capex is not None:
        capex_kombination = kapazitaet_kwh * invest_kosten_kwh
        params_override["unterhalt_kosten"] = unterhalt_pct_von_capex * capex_kombination

    zeile = {
        "kapazitaet_kwh": kapazitaet_kwh,
        "c_rate": c_rate,
        "leistung_kw": leistung_kw,
        "fehler": None,
    }
    try:
        module, ts, anteile, zeitreihen, params, pv_profile, last_profile = battery_optimization.main(
            input_path=input_lp_path, solver=solver, params_override=params_override,
        )
    except Exception as exc:  # noqa: BLE001 -- bewusst breit: ein Rasterpunkt darf den Sweep nicht stoppen
        zeile["fehler"] = f"{type(exc).__name__}: {exc}"
        return zeile

    # NEU: Hochrechnung auf ein volles Jahr, falls das Szenario kein volles
    # Jahr abdeckt (z.B. "..._Q2.xlsx", nur 3 Monate) -- exakt dieselbe Logik
    # wie output_create.build_rendite_kennzahlen()/build_income_expense_summary()
    # (dort Beats Frage "wieso sagt der Report eine Jahresbetrachtung, obwohl
    # es nur 3 Monate sind" ausgeloest). annuitaet/unterhalt aus capital_costs()
    # sind bereits echte Jahresgroessen (unabhaengig von der Simulationsdauer);
    # nur die IST-Werte aus der Simulation selbst (Einnahmen, Entladeenergie,
    # Degradation) muessen hochgerechnet werden, sonst wuerde z.B. bei einem
    # 3-Monats-Szenario nur ein Viertel Jahr Einnahmen gegen volle
    # Jahres-Kapitalkosten gerechnet (Gewinn/Verlust faelschlich ~4x zu tief).
    dauer_tage = len(ts) * battery_optimization.DT_HOURS / 24.0
    dauer_jahre = dauer_tage / 365.25 if dauer_tage > 0 else 1.0

    entladeenergie_kwh_ist = float(ts["dis"].sum() * battery_optimization.DT_HOURS)
    kk = battery_optimization.capital_costs(params, wacc=WACC, entladeenergie_kwh=entladeenergie_kwh_ist)

    total_operativ_ist = sum(module[k] for k in output_create.EINNAHMEN_LABELS if k in module)
    total_operativ = total_operativ_ist / dauer_jahre
    entladeenergie_kwh = entladeenergie_kwh_ist / dauer_jahre
    degradation_jahr = kk.get("degradation", 0.0) / dauer_jahre

    gewinn_verlust = total_operativ - kk["annuitaet"] - kk["unterhalt"] - degradation_jahr
    netto_cashflow = total_operativ - kk["unterhalt"] - degradation_jahr
    kapitalverzinsung_pct = (netto_cashflow / kk["capex"] * 100.0) if kk["capex"] > 1e-9 else None

    zeile.update({
        "capex_chf": kk["capex"],
        "unterhalt_chf_jahr": kk["unterhalt"],
        "annuitaet_chf_jahr": kk["annuitaet"],
        "degradation_chf_jahr": degradation_jahr,
        "total_operativ_chf_jahr": total_operativ,
        "gewinn_verlust_chf_jahr": gewinn_verlust,
        "kapitalverzinsung_pct": kapitalverzinsung_pct,
        "entladeenergie_kwh_jahr": entladeenergie_kwh,
    })
    return zeile


def sweep(input_lp_path: str, solver: str) -> pd.DataFrame:
    """Fuehrt den kompletten Raster-Sweep durch: Parameter/Profile EINMAL
    ohne Solve einlesen (fuer die Raster-Bestimmung), dann pro Rasterpunkt
    einen vollen Solve via _kombination_rechnen(). Gibt eine DataFrame mit
    einer Zeile pro Rasterpunkt zurueck (Basis fuer Excel + Heatmap-PDF)."""
    import battery_optimization  # lazy, siehe Modulkopf

    print("Lese Parameter/Zeitreihen (ohne Solve, nur fuer Raster-Bestimmung) ...")
    params_original, zeitreihen_original = battery_optimization.read_inputs(input_lp_path)
    pv_profile = battery_optimization.build_pv_profile(params_original, zeitreihen_original)
    last_profile = battery_optimization.build_last_profile(params_original, zeitreihen_original)

    kapazitaeten, _ = baue_raster(pv_profile, last_profile, battery_optimization.DT_HOURS)

    invest_kosten_kwh = params_original["invest_kosten_kwh"]
    capex_original = params_original["kapazitaet"] * invest_kosten_kwh
    if UNTERHALT_PROPORTIONAL_SKALIEREN and capex_original > 1e-9:
        unterhalt_pct_von_capex = params_original["unterhalt_kosten"] / capex_original
        print(
            f"  Unterhaltkosten werden proportional zur Investitionssumme skaliert "
            f"({unterhalt_pct_von_capex * 100:.2f}% des Capex, hergeleitet aus dem "
            f"Original-Excel-Wert {params_original['unterhalt_kosten']:,.0f} CHF/Jahr bei "
            f"{capex_original:,.0f} CHF Capex)."
        )
    else:
        unterhalt_pct_von_capex = None
        print(
            f"  Unterhaltkosten bleiben FIX bei {params_original['unterhalt_kosten']:,.0f} CHF/Jahr "
            f"fuer alle Rasterpunkte (UNTERHALT_PROPORTIONAL_SKALIEREN=False)."
        )

    rasterpunkte = [(kap, c) for kap in kapazitaeten for c in C_RATEN]
    print(
        f"\n{len(rasterpunkte)} Rasterpunkte ({len(kapazitaeten)} Kapazitaeten x "
        f"{len(C_RATEN)} C-Raten) -- jeder ist ein vollstaendiger oemof.solph-Solve, "
        f"das kann eine Weile dauern.\n"
    )

    zeilen = []
    for i, (kap, c_rate) in enumerate(rasterpunkte, start=1):
        print(
            f"[{i}/{len(rasterpunkte)}] Kapazitaet={kap:,.0f} kWh, C-Rate={c_rate} "
            f"(Leistung={c_rate * kap:,.0f} kW) ..."
        )
        zeile = _kombination_rechnen(
            input_lp_path, kap, c_rate, solver, invest_kosten_kwh, unterhalt_pct_von_capex,
        )
        if zeile["fehler"]:
            print(f"    FEHLER: {zeile['fehler']}")
        else:
            kv_txt = (
                f"{zeile['kapitalverzinsung_pct']:.1f}%"
                if zeile["kapitalverzinsung_pct"] is not None else "n/a"
            )
            print(
                f"    Gewinn/Verlust: {zeile['gewinn_verlust_chf_jahr']:,.0f} CHF/Jahr, "
                f"Kapitalverzinsung: {kv_txt}"
            )
        zeilen.append(zeile)

    return pd.DataFrame(zeilen)


# --------------------------------------------------------------------------
# 3. Ergebnis-Excel (reine Logik, kein oemof noetig)
# --------------------------------------------------------------------------

_EXCEL_SPALTEN_REIHENFOLGE = [
    "kapazitaet_kwh", "c_rate", "leistung_kw",
    "capex_chf", "unterhalt_chf_jahr", "annuitaet_chf_jahr", "degradation_chf_jahr",
    "total_operativ_chf_jahr", "gewinn_verlust_chf_jahr", "kapitalverzinsung_pct",
    "entladeenergie_kwh_jahr", "fehler",
]


def speichere_ergebnis_excel(df: pd.DataFrame, xlsx_path: str) -> None:
    """Schreibt die volle Sweep-Ergebnistabelle (eine Zeile pro Rasterpunkt)
    als Excel-Datei. Fehlende Spalten (z.B. wenn ALLE Rasterpunkte
    fehlgeschlagen sind) werden mit NaN ergaenzt, damit die Spaltenreihenfolge
    immer stabil bleibt."""
    df = df.reindex(columns=_EXCEL_SPALTEN_REIHENFOLGE)
    df.to_excel(xlsx_path, index=False, sheet_name="Sweep-Ergebnisse")
    print(f"Ergebnis-Excel geschrieben: {xlsx_path}")


# --------------------------------------------------------------------------
# 4. Heatmap-PDF (reine Logik, kein oemof noetig)
# --------------------------------------------------------------------------

def _heatmap_seite(
    pdf: PdfPages,
    werte_pivot: pd.DataFrame,
    leistung_pivot: pd.DataFrame,
    titel: str,
    einheit: str,
    diverging: bool,
    best_pos: tuple[int, int] | None,
) -> None:
    """Zeichnet EINE Heatmap-Seite (Kapazitaet x C-Rate) ins PDF. `diverging`
    steuert die Farbskala: True = zweifarbig (rot/gruen) um 0 zentriert
    (fuer Gewinn/Verlust und Kapitalverzinsung, die negativ werden koennen),
    False = einfarbig hell->dunkelgruen (fuer den operativen Gesamtertrag,
    der praktisch nie negativ ist -- siehe dataviz-Skill: sequentiell = eine
    Farbe, divergierend = zwei Farben + neutrale Mitte, nie ein Regenbogen)."""
    kapazitaeten = werte_pivot.index.tolist()
    c_raten = werte_pivot.columns.tolist()
    werte = werte_pivot.values.astype(float)

    if diverging:
        finite = werte[np.isfinite(werte)]
        vmax = float(np.max(np.abs(finite))) if finite.size else 1.0
        vmax = vmax if vmax > 1e-9 else 1.0
        norm = mcolors.TwoSlopeNorm(vmin=-vmax, vcenter=0.0, vmax=vmax)
        cmap = mcolors.LinearSegmentedColormap.from_list("fleco_div", [_NEG, "#FFFFFF", _POS])
    else:
        finite = werte[np.isfinite(werte)]
        vmax = float(np.max(finite)) if finite.size else 1.0
        vmax = vmax if vmax > 1e-9 else 1.0
        norm = mcolors.Normalize(vmin=0.0, vmax=vmax)
        cmap = mcolors.LinearSegmentedColormap.from_list("fleco_seq", ["#FFFFFF", _POS])

    fig, ax = plt.subplots(figsize=(9.0, 0.62 * len(kapazitaeten) + 2.4))
    masked = np.ma.masked_invalid(werte)
    im = ax.imshow(masked, cmap=cmap, norm=norm, aspect="auto")

    ax.set_xticks(range(len(c_raten)))
    ax.set_xticklabels([f"{c:g}C" for c in c_raten])
    ax.set_yticks(range(len(kapazitaeten)))
    ax.set_yticklabels([f"{k:,.0f} kWh" for k in kapazitaeten])
    ax.set_xlabel("C-Rate")
    ax.set_ylabel("Kapazitaet")
    ax.set_title(titel, fontsize=13, fontweight="bold", color=_TEXT)
    for spine in ax.spines.values():
        spine.set_visible(False)

    for i in range(len(kapazitaeten)):
        for j in range(len(c_raten)):
            wert = werte[i, j]
            leistung = leistung_pivot.values[i, j]
            if not np.isfinite(wert):
                text = "n/a" if not np.isfinite(leistung) else f"n/a\n({leistung:,.0f} kW)"
            else:
                text = f"{wert:,.0f}{einheit}\n({leistung:,.0f} kW)"
            ax.text(
                j, i, text, ha="center", va="center", fontsize=8, color=_TEXT,
                bbox=dict(facecolor="white", alpha=0.65, edgecolor="none", pad=1.5),
            )

    if best_pos is not None:
        bi, bj = best_pos
        ax.add_patch(
            plt.Rectangle((bj - 0.5, bi - 0.5), 1, 1, fill=False, edgecolor="black", linewidth=2.5)
        )

    fig.colorbar(im, ax=ax, shrink=0.85, label=f"{titel} [{einheit.strip()}]" if einheit.strip() else titel)
    fig.tight_layout()
    pdf.savefig(fig)
    plt.close(fig)


def baue_heatmap_pdf(df: pd.DataFrame, pdf_path: str) -> None:
    """Baut das PDF mit den drei Heatmaps (Gewinn/Verlust, Kapitalverzinsung,
    operativer Gesamtertrag) aus der Sweep-Ergebnistabelle. Rasterpunkte mit
    `fehler` gesetzt erscheinen automatisch als 'n/a'-Zelle (fehlender Wert
    im Pivot -> NaN -> maskiert/beschriftet, siehe _heatmap_seite()). Die
    Zelle mit dem hoechsten Gewinn/Verlust wird in ALLEN DREI Heatmaps an
    derselben Position schwarz umrandet."""
    leistung_pivot = df.pivot(index="kapazitaet_kwh", columns="c_rate", values="leistung_kw")
    gv_pivot = df.pivot(index="kapazitaet_kwh", columns="c_rate", values="gewinn_verlust_chf_jahr")
    kv_pivot = df.pivot(index="kapazitaet_kwh", columns="c_rate", values="kapitalverzinsung_pct")
    op_pivot = df.pivot(index="kapazitaet_kwh", columns="c_rate", values="total_operativ_chf_jahr")

    try:
        best_pos = tuple(int(x) for x in np.unravel_index(np.nanargmax(gv_pivot.values), gv_pivot.shape))
    except ValueError:
        best_pos = None
        print("  HINWEIS: kein einziger Rasterpunkt erfolgreich geloest -- keine 'beste' Zelle markierbar.")

    with PdfPages(pdf_path) as pdf:
        _heatmap_seite(pdf, gv_pivot, leistung_pivot, "Gewinn/Verlust netto", " CHF/Jahr", True, best_pos)
        _heatmap_seite(pdf, kv_pivot, leistung_pivot, "Durchschnittliche Kapitalverzinsung", "%", True, best_pos)
        _heatmap_seite(pdf, op_pivot, leistung_pivot, "Operativer Gesamtertrag", " CHF/Jahr", False, best_pos)

    print(f"Heatmap-PDF geschrieben: {pdf_path}")


# --------------------------------------------------------------------------
# 5. CLI / Orchestrierung (braucht input_prep/battery_optimization -> lazy)
# --------------------------------------------------------------------------

def pick_file_dialog() -> str:
    """Wie run_battery_analysis.py::pick_file_dialog() -- identisches
    Verhalten, hier dupliziert, damit dieses Skript unabhaengig lauffaehig
    bleibt (kein Import von run_battery_analysis.py noetig)."""
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError as exc:
        raise RuntimeError(
            "Kein Pfad angegeben und der Datei-Dialog (tkinter) ist in dieser "
            "Python-Installation nicht verfuegbar. Bitte den Pfad direkt als "
            "Argument angeben:\n"
            "  python optimale_batteriegroesse.py <Pfad-zu-Inputs.xlsx>"
        ) from exc

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askopenfilename(
        title="Inputs.xlsx auswaehlen (fuer den Batteriegroessen-Sweep)",
        filetypes=[("Excel-Dateien", "*.xlsx"), ("Alle Dateien", "*.*")],
    )
    root.destroy()
    if not path:
        raise SystemExit("Keine Datei ausgewaehlt -- Abbruch.")
    return path


def main(argv=None):
    import input_prep  # lazy, siehe Modulkopf

    parser = argparse.ArgumentParser(
        description="Raster-Sweep + Heatmap zur 'optimalen' Batteriegroesse (Kapazitaet x Leistung)."
    )
    parser.add_argument(
        "input_file", nargs="?", default=None,
        help="Pfad zur Inputs.xlsx. Falls weggelassen: Datei-Auswahl-Dialog.",
    )
    parser.add_argument("--solver", default="cbc", help="Solver (Default: cbc).")
    parser.add_argument(
        "--kein-swissix-fetch", action="store_true",
        help="ENTSO-E-Abfrage ueberspringen (z.B. ohne Internetzugang testen).",
    )
    args = parser.parse_args(argv)

    input_path = args.input_file or pick_file_dialog()
    input_path = os.path.abspath(input_path)
    if not os.path.isfile(input_path):
        raise SystemExit(f"Datei nicht gefunden: {input_path}")

    # Gleiche Konvention wie run_battery_analysis.py: der Sweep braucht einen
    # bereits vorhandenen "Output"-Unterordner neben der Inputs.xlsx (wird
    # bewusst NICHT automatisch angelegt) fuer input_lp/Sweep-Excel; nur das
    # Heatmap-PDF (das eigentliche "Ergebnis" fuer den Nutzer) landet direkt
    # im Inputs-Ordner, analog zum Bericht_*.pdf.
    data_dir = os.path.dirname(input_path)
    output_dir = os.path.join(data_dir, "Output")
    if not os.path.isdir(output_dir):
        raise SystemExit(
            f"Output-Ordner nicht gefunden: {output_dir}\n"
            "Bitte den Ordner 'Output' manuell neben der Inputs.xlsx anlegen "
            "und das Skript erneut starten (siehe run_battery_analysis.py, "
            "gleiche Konvention)."
        )

    input_lp_path = os.path.join(output_dir, output_create.derive_input_lp_filename(input_path))
    rest = output_create._strip_input_prefix(input_path)
    sweep_xlsx_path = os.path.join(output_dir, f"Sweep_{rest}.xlsx" if rest else "Sweep.xlsx")
    sweep_pdf_path = os.path.join(data_dir, f"Sweep_{rest}.pdf" if rest else "Sweep.pdf")

    print("=" * 70)
    print("Optimale Batteriegroesse -- Raster-Sweep")
    print(f"  Input:        {input_path}")
    print(f"  input_lp:     {input_lp_path}")
    print(f"  Sweep-Excel:  {sweep_xlsx_path}")
    print(f"  Heatmap-PDF:  {sweep_pdf_path}")
    print("=" * 70)

    print("\n=== Schritt 1: Input aufbereiten (input_prep, EINMAL fuer alle Rasterpunkte) ===")
    input_prep.main(
        input_path=input_path, output_path=input_lp_path, fetch_swissix=not args.kein_swissix_fetch,
    )

    print("\n=== Schritt 2: Raster-Sweep (ein voller Solve je Rasterpunkt) ===")
    df = sweep(input_lp_path, args.solver)

    print("\n=== Schritt 3: Ergebnis-Excel + Heatmap-PDF schreiben ===")
    speichere_ergebnis_excel(df, sweep_xlsx_path)
    baue_heatmap_pdf(df, sweep_pdf_path)

    anzahl_fehler = int(df["fehler"].notna().sum())
    if anzahl_fehler:
        print(f"\nHINWEIS: {anzahl_fehler} von {len(df)} Rasterpunkten sind fehlgeschlagen "
              f"(siehe 'fehler'-Spalte im Sweep-Excel bzw. 'n/a'-Zellen im Heatmap-PDF).")
    print("\nFertig.")


if __name__ == "__main__":
    main()
