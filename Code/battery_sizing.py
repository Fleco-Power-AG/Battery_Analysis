"""
battery_sizing.py
==================

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
dessen main()-Docstring: "NEU (15.9.2026, fuer battery_sizing.py)").

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

RASTER-BESTIMMUNG (NEU GEAENDERT 22.9.2026 -- ZWEITE KORREKTUR, Beat: "bei
einer durchschnittlichen Last von 30 kWh hat der code Batteriegroessen von
1500kWh bis 9000kWh getestet, was voellig absurd ist ... zurueck zur 'auto'
groessendetection aber in einem sinnvollen Verhaeltnis von Last und PV")
--------------------------------------------------------------------------
  GESCHICHTE (damit klar ist, was warum nochmals geaendert wurde):
    1. Urspruenglich: Kapazitaets-Basis aus der typischen PV-Ueberschuss-/
       Lastenergie hergeleitet, x 0.5/0.75/1/1.5/2/3 multipliziert -- bei
       grossen Anlagen kombiniert mit C-Rate 1.0 zu unrealistischen
       MW-Leistungen gefuehrt (Beats erste Beobachtung).
    2. Als Korrektur (kurzzeitig): Ausgangspunkt war die im Excel manuell
       ANGEGEBENE Batterieleistung, +/- feste 100-kW-Schritte, Kapazitaet
       daraus als Leistung/C-Rate zurueckgerechnet. PROBLEM (Beats zweite
       Beobachtung): das Teilen durch eine kleine C-Rate (0.25) vervierfacht
       den Wert, und wenn die im Excel angegebene/hergeleitete Leistung
       schon nicht gut zur tatsaechlichen Last/PV-Groesse passte (z.B. ein
       Platzhalterwert), explodierte die Kapazitaet trotzdem unrealistisch
       (30 kWh Last -> 1500-9000 kWh Vorschlaege).
    3. AKTUELL: zurueck zu einer vollautomatischen, PROFIL-basierten
       Kapazitaets-Basis (kein Excel-Eintrag noetig, Beats urspruenglicher
       Wunsch "ich moechte eigentlich keine Eingabe machen") -- siehe
       bestimme_kapazitaets_basis() -- KOMBINIERT mit relativen (prozentualen)
       statt absoluten Rasterschritten (siehe baue_raster()), damit die
       Spannweite IMMER proportional zur tatsaechlichen Last/PV-Groesse
       bleibt, egal ob es sich um eine 30-kWh- oder eine 3-MWh-Anlage
       handelt -- das behebt strukturell BEIDE oben genannten Symptome (kein
       fixer Absolutbetrag mehr, der je nach Anlagengroesse zu klein oder zu
       gross ist; kein Teilen durch C-Rate mehr, das kleine Ungenauigkeiten
       vervierfacht).

  KAPAZITAETS-BASIS (bestimme_kapazitaets_basis(), automatisch aus dem
  Profil, in drei Stufen):
    1. Typische taegliche PV-UEBERSCHUSSENERGIE (PV minus Last, positiver
       Teil) -- die Energiemenge, die an einem Durchschnittstag ueberhaupt
       sinnvoll zwischengespeichert werden koennte.
    2. Kein nennenswerter PV-Ueberschuss (PV < Last oder keine PV-Anlage)
       -> typische taegliche LAST-Energie (Batterie zur Lastspitzenkappung/
       Arbitrage relativ zur Last).
    3. Weder PV-Ueberschuss noch Last vorhanden (reines Arbitrage-/
       SRL-System) -> nichts aus dem Profil ableitbar, dann greift ein
       fester Default-Bereich (KAPAZITAETEN_FALLBACK_KWH) mit Konsolen-
       Hinweis.

  RASTER UM DIE BASIS (baue_raster(), RELATIV/PROZENTUAL statt absolut --
  NEU, Beat 24.9.2026: "koennen wir die range der kapazitaet etwas erhoehen
  zu 50, 100, 200% schritten?"):
    Kapazitaetsstufen = Basis * (1 +/- p), fuer jeden Prozentsatz p in
    KAPAZITAETS_SCHRITT_PROZENTE (Default [0.50, 1.00, 2.00] -> Multipli-
    katoren 1 +/- 0.5/1.0/2.0 = {-1.0, 0.0, 0.5, 1.0, 1.5, 2.0, 3.0}; die
    beiden <= 0 werden verworfen, es bleiben 0.5x/1.0x/1.5x/2.0x/3.0x der
    Basis -- 5 Kapazitaetsstufen), auf "schoene" Zehner-/Fuenfziger-/
    Hunderter-Schritte gerundet, siehe _rund_schoen(). Jede Kapazitaetsstufe
    wird mit denselben drei C-Raten (0.25/0.5/1.0, Batterie-Fachbegriff: "wie
    schnell relativ zur Kapazitaet geladen/entladen werden kann", siehe
    C_RATEN) kombiniert; die LEISTUNG ergibt sich daraus wie urspruenglich
    als Leistung[kW] = C-Rate * Kapazitaet[kWh] (KEINE Division mehr --
    das war der Verstaerkungseffekt in Version 2 oben). Macht zusammen 5
    Kapazitaetsstufen x 3 C-Raten = 15 Rasterpunkte im Default (alles ueber
    die Konstanten am Skriptanfang aenderbar).

  AUTOMATISCHE RASTER-ERWEITERUNG (NEU, Beat 24.9.2026: "bei allen Runs die
  ich gemacht habe war C=1 und die groesste Batterie-Konfiguration die beste
  Wahl ... kann es sein dass das Autosizing auf zu kleine Batterien geht
  oder/oder die Kosten fuer die groessere Batterie nicht sauber eingelesen
  werden?"):
    Diagnose anhand eines konkreten Sweeps von Beat: die Kosten wurden
    korrekt gelesen/angewendet (Capex/Annuitaet skalierten sauber linear mit
    Kapazitaet * dem C-Rate-spezifischen CHF/kWh-Preis, siehe
    INVEST_KOSTEN_KWH_JE_C_RATE) -- das Problem war stattdessen, dass die
    Gewinnkurve am oberen Rand des getesteten Rasters (3x Basis) noch
    STIEG, statt ein inneres Maximum zu erreichen. "Die groesste getestete
    Batterie gewinnt" ist in diesem Fall zwangslaeufig, unabhaengig von den
    Kosten -- das Raster hat den eigentlichen wirtschaftlichen Peak schlicht
    nie getestet. sweep() prueft deshalb jetzt nach jeder Rasterrunde, ob das
    bisher beste Ergebnis bei der GROESSTEN getesteten Kapazitaet liegt, und
    haengt in diesem Fall automatisch eine weitere, noch groessere
    Kapazitaetsstufe an (+1x Basis pro Runde, siehe
    KAPAZITAETS_ERWEITERUNGS_SCHRITT) -- so lange, bis entweder ein echtes
    inneres Maximum gefunden wird oder KAPAZITAETS_MAX_ERWEITERUNGEN erreicht
    ist (Sicherheitslimit fuer Szenarien ohne natuerliche Saettigung).

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
Jahr (15-Minuten-Schritte). NEU (Beat, 22.9.2026: "das geht zeitweise sehr
lange, kann man das verschnellern?"): die 18 Standard-Rasterpunkte laufen
jetzt standardmaessig PARALLEL in separaten Prozessen (ein Rasterpunkt pro
CPU-Kern, minus 1 -- siehe SWEEP_PARALLEL_WORKERS_DEFAULT unten), statt
nacheinander. Das skaliert praktisch mit der Kernzahl (4 Kerne -> grob 4x
schneller), da jeder Rasterpunkt ein komplett unabhaengiger Solve ist. Steuerbar
per CLI-Flag `--parallel N` (`--parallel 1` = alter, sequenzieller Ablauf,
z.B. zum Debuggen, da die Konsolenausgabe dann strikt in Raster-Reihenfolge
erscheint). Zusaetzlich fuer einen ersten schnellen Test weiterhin moeglich:
die Konstanten KAPAZITAETS_SCHRITT_PROZENTE/C_RATEN unten auf weniger Werte
kuerzen (z.B. nur `C_RATEN = [1.0]` und `KAPAZITAETS_SCHRITT_PROZENTE = []`
fuer einen einzelnen Testlauf mit nur der Basis-Kapazitaet).

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
import concurrent.futures
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

# NEU (Beat, 22.9.2026, zweite Korrektur: "zurueck zur 'auto' groessen-
# detection aber in einem sinnvollen Verhaeltnis von Last und PV") -- die
# Kapazitaets-Basis kommt wieder automatisch aus dem PV-/Lastprofil (siehe
# bestimme_kapazitaets_basis()), das Raster darum ist aber jetzt PROZENTUAL
# statt in absoluten kWh-Bloecken gestaffelt, damit die Spannweite IMMER
# proportional zur tatsaechlichen Anlagengroesse bleibt (siehe Modulkopf/
# Historie -- das behebt sowohl die MW-Ausreisser bei grossen Anlagen als
# auch die absurden 1000er-kWh-Vorschlaege bei einer kleinen 30-kWh-Last).
#
# NEU (Beat, 24.9.2026: "koennen wir die range der kapazitaet etwas erhoehen
# zu 50, 100, 200% schritten?") -- statt eines EINZELNEN, gleich grossen
# Prozentsatzes pro Schritt (vorher: 3x 10%) jetzt eine LISTE unterschiedlich
# grosser Schritte, symmetrisch um die Basis (+/- je Eintrag). Mit den
# Default-Werten unten ergibt das Multiplikatoren 1 +/- 0.5/1.0/2.0 =
# {-1.0, 0.0, 0.5, 1.0, 1.5, 2.0, 3.0} -- die beiden <= 0 (Basis -100%/-200%)
# werden wie gehabt automatisch verworfen (siehe baue_raster()), es bleiben
# 0.5x/1.0x/1.5x/2.0x/3.0x der Basis.
KAPAZITAETS_SCHRITT_PROZENTE = [0.50, 1.00, 2.00]
# Nur falls das Profil GAR NICHTS liefert (weder PV-Ueberschuss noch Last,
# reines Arbitrage-/SRL-System ohne beides) -- siehe bestimme_kapazitaets_basis().
KAPAZITAETEN_FALLBACK_KWH = [100.0, 200.0, 300.0, 500.0, 750.0, 1000.0]

# NEU (Beat, 24.9.2026: "bei allen Runs die ich gemacht habe war C=1 und die
# groesste Batterie-Konfiguration die beste Wahl ... kann es sein dass das
# Autosizing auf zu kleine Batterien geht?") -- wenn das beste Ergebnis eines
# Sweeps bei der GROESSTEN getesteten Kapazitaet liegt, ist das Raster
# hoechstwahrscheinlich zu schmal (die Gewinnkurve war am oberen Rand noch
# steigend, der eigentliche Peak wurde also nie getestet) -- NICHT
# zwangslaeufig ein Kosten-Bug. sweep() erweitert das Kapazitaets-Raster in
# diesem Fall jetzt AUTOMATISCH um weitere, groessere Stufen (je +1x Basis),
# bis entweder ein echtes (inneres) Maximum gefunden wird oder dieses
# Sicherheitslimit erreicht ist (falls die Gewinnkurve z.B. wegen sehr hohem
# Arbitrage-Potenzial gar kein natuerliches Maximum hat). Gilt NUR, wenn eine
# Profil-Basis > 0 hergeleitet werden konnte (nicht fuer den festen
# KAPAZITAETEN_FALLBACK_KWH-Bereich, da dort keine proportionale Erweiterung
# moeglich ist).
KAPAZITAETS_AUTO_ERWEITERN = True
KAPAZITAETS_ERWEITERUNGS_SCHRITT = 1.0  # je Runde: + 1x Basis oben drauf
KAPAZITAETS_MAX_ERWEITERUNGEN = 3

# NEU (Beat, 24.9.2026: "jetzt muessen wir noch die batteriekosten anpassen,
# fuer 0.25 C 400CHF/kWh, 0.5C 500 CHF/kWh und fuer 1C 700 CHF/kWh") --
# schnellere Batterien (hoehere C-Rate) brauchen leistungsfaehigere Zellen/
# Leistungselektronik und sind pro kWh teurer. Bisher wurde EIN einziger
# CHF/kWh-Preis (aus Parameter!C34, oder aus C41 hergeleitet, siehe
# battery_optimization.capital_costs()) fuer ALLE C-Raten im Sweep verwendet
# -- das ignorierte diesen realen Kostenunterschied. Ersetzt/hat Vorrang vor
# dem Excel-Wert FUER DEN SWEEP (nicht fuer den normalen Einzellauf via
# run_battery_analysis.py -- dort bleibt C34/C41 unveraendert maassgeblich,
# da es dort keine C-Rate-Achse gibt). Wird eine C-Rate hier verwendet
# (siehe C_RATEN oben), die NICHT in dieser Tabelle steht, faellt sweep()
# mit einer Konsolen-Warnung auf den Excel-Wert zurueck (siehe unten).
INVEST_KOSTEN_KWH_JE_C_RATE = {
    0.25: 400.0,
    0.5: 500.0,
    1.0: 700.0,
}

UNTERHALT_PROPORTIONAL_SKALIEREN = True
WACC = 0.03  # gleiche Default-Annahme wie battery_optimization.capital_costs()

# NEU (Beat, 22.9.2026: "das geht zeitweise sehr lange, kann man das
# verschnellern?"): jeder Rasterpunkt ist ein komplett unabhaengiger
# oemof.solph-Solve (eigener CBC-Subprozess) -- die 18 Standard-Rasterpunkte
# liessen sich bisher aber sequenziell (einer nach dem anderen) rechnen.
# Default hier: alle Kerne bis auf einen (damit der Rechner waehrenddessen
# noch bedienbar bleibt) -- ueberschreibbar per CLI-Flag "--parallel N".
# "--parallel 1" schaltet zurueck auf den alten, sequenziellen Ablauf.
SWEEP_PARALLEL_WORKERS_DEFAULT = max(1, (os.cpu_count() or 2) - 1)

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
    Raster her, VOLLAUTOMATISCH aus dem PV-/Lastprofil (Beat: "ich moechte
    eigentlich keine Eingabe machen" / "zurueck zur 'auto' groessendetection
    aber in einem sinnvollen Verhaeltnis von Last und PV"), in drei Stufen:
      1. PV-Ueberschuss (PV - Last, positiver Teil) -- die Batterie speichert
         PV-Ueberschuss fuer spaeter.
      2. Last-Energie -- kein nennenswerter PV-Ueberschuss (z.B. PV < Last
         oder keine PV-Anlage), Batterie dient v.a. Lastspitzenkappung/
         Arbitrage relativ zur Last.
      3. Fallback -- weder PV-Ueberschuss noch Last vorhanden (reines
         Arbitrage-/SRL-System) -- Basis 0.0, Aufrufer (baue_raster())
         verwendet dann den festen Default-Bereich statt Prozent-Schritten.
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
    """Baut die Liste der Kapazitaets-Rasterpunkte (kWh): PROZENTUALE Schritte
    um die automatisch abgeleitete Basis (siehe bestimme_kapazitaets_basis()),
    gemaess KAPAZITAETS_SCHRITT_PROZENTE symmetrisch (+/- je Eintrag) --
    Default [0.50, 1.00, 2.00] -> Multiplikatoren 0.5x/1.0x/1.5x/2.0x/3.0x
    der Basis (die rechnerisch <= 0 liegenden Stufen, hier Basis-100%/-200%,
    werden automatisch verworfen, siehe unten), 'schoen' gerundet, Duplikate
    entfernt, aufsteigend sortiert. Bewusst PROZENTUAL (nicht in fixen kWh-
    oder kW-Bloecken wie in frueheren Versionen -- siehe Modulkopf/Historie):
    das haelt die Spannweite IMMER proportional zur tatsaechlichen Last-/
    PV-Groesse, egal ob 30 kWh oder 3 MWh Basis. Ist keine Basis herleitbar
    (kein PV-Ueberschuss und keine Last), wird stattdessen der feste
    Fallback-Bereich KAPAZITAETEN_FALLBACK_KWH verwendet (keine Prozent-
    Schritte moeglich, da jeder Prozentsatz von 0 immer 0 waere)."""
    basis_kwh, herkunft = bestimme_kapazitaets_basis(pv_profile, last_profile, dt_hours)
    if basis_kwh <= 1e-6:
        print(f"  HINWEIS: {herkunft} -- verwende feste Default-Kapazitaeten "
              f"{KAPAZITAETEN_FALLBACK_KWH} kWh.")
        return list(KAPAZITAETEN_FALLBACK_KWH), herkunft

    print(f"  Kapazitaets-Basis: {basis_kwh:,.1f} kWh/Tag ({herkunft})")
    alle_multiplikatoren = sorted(
        {1.0} | {1.0 + vz * p for p in KAPAZITAETS_SCHRITT_PROZENTE for vz in (1, -1)}
    )
    alle_stufen = [_rund_schoen(basis_kwh * m) for m in alle_multiplikatoren]
    kapazitaeten = sorted({k for k in alle_stufen if k > 0})

    verworfen = len(alle_stufen) - len(kapazitaeten)
    if verworfen:
        print(f"  HINWEIS: {verworfen} Kapazitaetsstufe(n) <= 0 kWh wurden verworfen "
              f"(Schritt-Prozentsatz >= 100% ist groesser als die Basis selbst).")

    print(f"  Kapazitaets-Raster: {kapazitaeten} kWh (Multiplikatoren {alle_multiplikatoren} der "
          f"Basis, je kombiniert mit C-Raten {C_RATEN})")
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
    params_original: dict | None = None,
    zeitreihen_original: pd.DataFrame | None = None,
) -> dict:
    """Rechnet EINE Kapazitaet/C-Rate-Kombination komplett durch (ein voller
    oemof.solph-Solve via battery_optimization.main(params_override=...)) und
    fasst das Ergebnis in einem flachen Dict zusammen (eine Zeile der
    spaeteren Ergebnistabelle). Faengt Solver-/Feasibility-Fehler ab (z.B.
    Infeasibility bei einer sehr kleinen Batterie) und gibt dann ein Dict mit
    gesetztem `fehler` zurueck, statt den ganzen Sweep abzubrechen.

    NEU (Beat, 22.9.2026, zweite Korrektur): `kapazitaet_kwh` ist wieder die
    UNABHAENGIGE Groesse (aus dem Kapazitaets-Raster, siehe baue_raster()) --
    die Leistung ergibt sich rechnerisch als C-Rate * Kapazitaet, NICHT mehr
    umgekehrt (das Teilen durch C-Rate in der Zwischenversion hat kleine
    Ungenauigkeiten der Basis vervierfacht, siehe Modulkopf/Historie).

    `invest_kosten_kwh` (NEU, Beat 24.9.2026: "batteriekosten anpassen, fuer
    0.25 C 400CHF/kWh, 0.5C 500 CHF/kWh und fuer 1C 700 CHF/kWh"): sweep()
    reicht hier bereits den zu DIESER c_rate passenden Preis durch (aus
    INVEST_KOSTEN_KWH_JE_C_RATE, mit Fallback auf den Excel-/C41-Wert fuer
    nicht hinterlegte C-Raten) -- diese Funktion selbst muss die C-Rate also
    nicht mehr extra nachschlagen.

    `params_original`/`zeitreihen_original` (Beat, 22.9.2026: "man muss doch
    nicht fuer jedes Problem das ganze Input neu laden, da sich ja jeweils
    nur die Batterie-Parameter aendern"): wenn beide gesetzt sind (sweep()
    liest sie ohnehin schon EINMAL fuer die Raster-Bestimmung ein), werden
    sie an battery_optimization.main() als preloaded_params/
    preloaded_zeitreihen durchgereicht -- main() liest dann das Excel NICHT
    nochmals ein, sondern nutzt direkt diese bereits geparsten Objekte. Nur
    die Batteriegroesse (kapazitaet/leistung/unterhalt_kosten) unterscheidet
    sich ohnehin zwischen den Rasterpunkten, siehe params_override oben --
    Tarife/SwissIX/PV-Referenzprofil/Zeitreihen sind fuer alle Rasterpunkte
    identisch. Bei parallelem Betrieb (ProcessPoolExecutor) werden diese
    Objekte statt einer Excel-Neuparsen also nur noch (deutlich billiger)
    an den jeweiligen Worker-Prozess durchgereicht/gepickelt."""
    import battery_optimization  # lazy, siehe Modulkopf

    leistung_kw = round(c_rate * kapazitaet_kwh, 3)
    capex_kombination = kapazitaet_kwh * invest_kosten_kwh
    params_override = {
        "kapazitaet": kapazitaet_kwh,
        "leistung": leistung_kw,
        # NEU (Beat, 22.9.2026, Parameter!C41 "gegebenenfalls die
        # Investitionskosten" als fixer Totalbetrag statt CHF/kWh): ein
        # fixer Totalbetrag aus dem Original-Excel gilt nur fuer DESSEN
        # Original-Kapazitaet -- fuer andere Rasterpunkte waere er falsch
        # (zu hoch/zu tief). `invest_kosten_kwh` hier ist bereits der von
        # sweep() aufgeloeste EFFEKTIVE CHF/kWh-Satz (entweder direkt aus
        # C34, oder -- falls C41 gesetzt war -- aus C41 / Original-Kapazitaet
        # hergeleitet, siehe sweep()). "investitionskosten_fix" wird hier
        # deshalb explizit geleert, damit battery_optimization.capital_costs()
        # zuverlaessig kapazitaet_kwh * invest_kosten_kwh fuer JEDEN
        # Rasterpunkt neu rechnet, statt den fixen Original-Totalbetrag
        # unveraendert fuer alle Groessen zu uebernehmen.
        "invest_kosten_kwh": invest_kosten_kwh,
        "investitionskosten_fix": None,
    }
    if unterhalt_pct_von_capex is not None:
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
            preloaded_params=params_original, preloaded_zeitreihen=zeitreihen_original,
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


def sweep(input_lp_path: str, solver: str, parallel_workers: int = 1) -> pd.DataFrame:
    """Fuehrt den kompletten Raster-Sweep durch: Parameter/Profile EINMAL
    ohne Solve einlesen (fuer die Raster-Bestimmung, siehe baue_raster()),
    dann pro Rasterpunkt (Kapazitaetsstufe x C-Rate) einen vollen Solve via
    _kombination_rechnen(). Gibt eine DataFrame mit einer Zeile pro
    Rasterpunkt zurueck (Basis fuer Excel + Heatmap-PDF).

    `parallel_workers` (NEU, siehe SWEEP_PARALLEL_WORKERS_DEFAULT oben):
    > 1 rechnet die Rasterpunkte in separaten Prozessen (ProcessPoolExecutor)
    gleichzeitig, statt einen nach dem anderen -- jeder Rasterpunkt ist ein
    komplett unabhaengiger Solve (kein gemeinsamer Zustand ausser der rein
    LESEND geteilten input_lp-Datei), das skaliert daher praktisch linear mit
    der Kernzahl. Reihenfolge der Ergebnis-DataFrame ist bei Parallelbetrieb
    NICHT mehr die Raster-Reihenfolge (haengt davon ab, welcher Solve zuerst
    fertig wird) -- unproblematisch, da speichere_ergebnis_excel()/
    baue_heatmap_pdf() ohnehin ueber df.pivot(index=..., columns=...) gehen,
    was automatisch nach Kapazitaet/C-Rate sortiert.

    NEU (siehe KAPAZITAETS_AUTO_ERWEITERN oben): die zurueckgegebene Anzahl
    Zeilen kann groesser sein als das urspruengliche Kapazitaeten x C-Raten-
    Raster, falls das Raster automatisch um weitere Kapazitaetsstufen
    erweitert wurde (weil das beste Ergebnis sonst am oberen Rand gelegen
    haette) -- siehe Konsolenausgabe fuer Details."""
    import battery_optimization  # lazy, siehe Modulkopf

    print("Lese Parameter/Zeitreihen (ohne Solve, nur fuer Raster-Bestimmung) ...")
    params_original, zeitreihen_original = battery_optimization.read_inputs(input_lp_path)
    pv_profile = battery_optimization.build_pv_profile(params_original, zeitreihen_original)
    last_profile = battery_optimization.build_last_profile(params_original, zeitreihen_original)

    kapazitaeten, _ = baue_raster(pv_profile, last_profile, battery_optimization.DT_HOURS)
    # NEU (siehe KAPAZITAETS_AUTO_ERWEITERN oben): fuer die automatische
    # Raster-Erweiterung wird die Basis separat noch einmal gebraucht (reine
    # numpy-Berechnung, vernachlaessigbar teuer -- billiger als baue_raster()'s
    # Rueckgabewert dafuer extra umzubauen).
    basis_kwh, _ = bestimme_kapazitaets_basis(pv_profile, last_profile, battery_optimization.DT_HOURS)

    # NEU (Beat, 22.9.2026): Parameter!C41 kann "gegebenenfalls" einen FIXEN
    # Total-Investitionsbetrag (CHF) vorgeben statt des CHF/kWh-Ansatzes (C34)
    # -- siehe battery_optimization.capital_costs(). Dieser fixe Betrag gilt
    # aber nur fuer die im Excel angegebene ORIGINAL-Kapazitaet; fuer den
    # Sweep (andere Kapazitaeten pro Rasterpunkt) leiten wir daraus einen
    # EFFEKTIVEN CHF/kWh-Satz her (Fixbetrag / Original-Kapazitaet) und
    # rechnen damit fuer jeden Rasterpunkt einzeln neu (siehe
    # _kombination_rechnen(), die "investitionskosten_fix" je Rasterpunkt
    # explizit leert, damit der fixe Original-Betrag nicht faelschlich fuer
    # ALLE Groessen uebernommen wird).
    kapazitaet_original = params_original["kapazitaet"]
    investitionskosten_fix = params_original.get("investitionskosten_fix")
    if (
        isinstance(investitionskosten_fix, (int, float))
        and investitionskosten_fix > 1e-9
        and kapazitaet_original > 1e-9
    ):
        invest_kosten_kwh_fallback = investitionskosten_fix / kapazitaet_original
        capex_original = float(investitionskosten_fix)
        print(
            f"  HINWEIS: Investitionskosten sind im Input-Excel als fixer Totalbetrag "
            f"angegeben (Parameter!C41 = {investitionskosten_fix:,.0f} CHF) statt als "
            f"CHF/kWh-Ansatz. Fuer den Sweep wird daraus ein effektiver Satz von "
            f"{invest_kosten_kwh_fallback:,.2f} CHF/kWh hergeleitet ({investitionskosten_fix:,.0f} CHF / "
            f"{kapazitaet_original:,.0f} kWh Original-Kapazitaet) und je Rasterpunkt neu skaliert."
        )
    else:
        invest_kosten_kwh_fallback = params_original["invest_kosten_kwh"]
        capex_original = kapazitaet_original * invest_kosten_kwh_fallback

    # NEU (Beat, 24.9.2026: "batteriekosten anpassen, fuer 0.25 C 400CHF/kWh,
    # 0.5C 500 CHF/kWh und fuer 1C 700 CHF/kWh") -- fuer JEDE im Sweep
    # verwendete C-Rate wird hier EINMAL der passende CHF/kWh-Preis aufgeloest
    # (aus INVEST_KOSTEN_KWH_JE_C_RATE, sonst Fallback auf den Excel-/
    # C41-Wert oben, mit Warnung). Wird unten pro Rasterpunkt weitergegeben,
    # statt eines einzigen, fuer alle C-Raten gleichen Preises.
    invest_kosten_kwh_je_c_rate = {}
    for c in C_RATEN:
        if c in INVEST_KOSTEN_KWH_JE_C_RATE:
            invest_kosten_kwh_je_c_rate[c] = INVEST_KOSTEN_KWH_JE_C_RATE[c]
        else:
            print(f"  HINWEIS: kein Preis in INVEST_KOSTEN_KWH_JE_C_RATE fuer C-Rate {c} "
                  f"hinterlegt -- verwende Fallback {invest_kosten_kwh_fallback:,.2f} CHF/kWh.")
            invest_kosten_kwh_je_c_rate[c] = invest_kosten_kwh_fallback
    print(f"  Batteriekosten je C-Rate: " + ", ".join(
        f"{c:g}C = {invest_kosten_kwh_je_c_rate[c]:,.0f} CHF/kWh" for c in C_RATEN
    ))

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

    def _ergebnis_zeile_drucken(nr: int, gesamt: int, kap: float, c_rate: float, zeile: dict) -> None:
        praefix = f"[{nr}/{gesamt}] Kapazitaet={kap:,.0f} kWh, C-Rate={c_rate}"
        if zeile["fehler"]:
            print(f"{praefix}: FEHLER: {zeile['fehler']}")
        else:
            kv_txt = (
                f"{zeile['kapitalverzinsung_pct']:.1f}%"
                if zeile["kapitalverzinsung_pct"] is not None else "n/a"
            )
            print(
                f"{praefix}: Gewinn/Verlust {zeile['gewinn_verlust_chf_jahr']:,.0f} CHF/Jahr, "
                f"Kapitalverzinsung {kv_txt}"
            )

    def _punkte_rechnen(punkte: list[tuple[float, float]]) -> list[dict]:
        """Rechnet eine Liste von (Kapazitaet, C-Rate)-Rasterpunkten durch --
        sequenziell oder parallel (ProcessPoolExecutor), je nach dem an
        sweep() uebergebenen `parallel_workers`. Ausgelagert aus dem
        Hauptkoerper von sweep(), damit sowohl das urspruengliche Raster als
        auch spaeter automatisch angehaengte Erweiterungsrunden (siehe
        KAPAZITAETS_AUTO_ERWEITERN) dieselbe Sequenziell-/Parallel-Logik
        verwenden -- jeder Aufruf bestimmt die Worker-Zahl neu anhand der
        Anzahl UEBERGEBENER Punkte (z.B. nur 3 bei einer Erweiterungsrunde mit
        einer einzelnen neuen Kapazitaetsstufe x 3 C-Raten)."""
        gesamt = len(punkte)
        workers = max(1, min(parallel_workers, gesamt))
        ergebnisse = []
        if workers <= 1:
            # Unveraendertes, sequenzielles Verhalten (z.B. bei --parallel 1,
            # oder wenn nur 1 Rasterpunkt existiert) -- einfacher zu
            # debuggen, da die Konsolenausgabe strikt in Raster-Reihenfolge
            # erscheint.
            for i, (kap, c_rate) in enumerate(punkte, start=1):
                print(f"[{i}/{gesamt}] Kapazitaet={kap:,.0f} kWh, C-Rate={c_rate} "
                      f"(Leistung={c_rate * kap:,.0f} kW, "
                      f"{invest_kosten_kwh_je_c_rate[c_rate]:,.0f} CHF/kWh) ...")
                zeile = _kombination_rechnen(
                    input_lp_path, kap, c_rate, solver, invest_kosten_kwh_je_c_rate[c_rate],
                    unterhalt_pct_von_capex,
                    params_original=params_original, zeitreihen_original=zeitreihen_original,
                )
                _ergebnis_zeile_drucken(i, gesamt, kap, c_rate, zeile)
                ergebnisse.append(zeile)
        else:
            print(f"  Rechne mit {workers} parallelen Prozessen "
                  f"(einer je CPU-Kern, minus 1) -- Reihenfolge der Meldungen unten "
                  f"entspricht der Fertigstellung, nicht der Raster-Reihenfolge.\n")
            with concurrent.futures.ProcessPoolExecutor(max_workers=workers) as pool:
                future_zu_punkt = {
                    pool.submit(
                        _kombination_rechnen, input_lp_path, kap, c_rate, solver,
                        invest_kosten_kwh_je_c_rate[c_rate], unterhalt_pct_von_capex,
                        params_original, zeitreihen_original,
                    ): (kap, c_rate)
                    for kap, c_rate in punkte
                }
                for nr, future in enumerate(concurrent.futures.as_completed(future_zu_punkt), start=1):
                    kap, c_rate = future_zu_punkt[future]
                    try:
                        zeile = future.result()
                    except Exception as exc:  # noqa: BLE001 -- _kombination_rechnen faengt normalerweise
                        # selbst schon ab; dieser Fallback greift nur bei einem
                        # unerwarteten Absturz DES WORKER-PROZESSES selbst (z.B.
                        # Speicherfehler) -- auch dann soll der restliche Sweep
                        # weiterlaufen statt komplett abzubrechen.
                        zeile = {
                            "kapazitaet_kwh": kap, "c_rate": c_rate,
                            "leistung_kw": round(c_rate * kap, 3),
                            "fehler": f"{type(exc).__name__}: {exc}",
                        }
                    _ergebnis_zeile_drucken(nr, gesamt, kap, c_rate, zeile)
                    ergebnisse.append(zeile)
        return ergebnisse

    rasterpunkte = [(kap, c) for kap in kapazitaeten for c in C_RATEN]
    print(
        f"\n{len(rasterpunkte)} Rasterpunkte ({len(kapazitaeten)} Kapazitaeten x "
        f"{len(C_RATEN)} C-Raten) -- jeder ist ein vollstaendiger oemof.solph-Solve, "
        f"das kann eine Weile dauern."
    )
    zeilen = _punkte_rechnen(rasterpunkte)

    # NEU (Beat, 24.9.2026: "bei allen Runs die ich gemacht habe war C=1 und
    # die groesste Batterie-Konfiguration die beste Wahl ... kann es sein
    # dass das Autosizing auf zu kleine Batterien geht?") -- siehe
    # KAPAZITAETS_AUTO_ERWEITERN oben: wenn das bisher beste Ergebnis bei der
    # GROESSTEN getesteten Kapazitaet liegt, war das Raster vermutlich zu
    # schmal (Gewinnkurve am oberen Rand noch steigend) -- automatisch eine
    # weitere, groessere Kapazitaetsstufe anhaengen und neu rechnen, bis ein
    # echtes inneres Maximum gefunden wird oder das Sicherheitslimit erreicht
    # ist.
    if KAPAZITAETS_AUTO_ERWEITERN and basis_kwh > 1e-6:
        erweiterungsrunde = 0
        while erweiterungsrunde < KAPAZITAETS_MAX_ERWEITERUNGEN:
            gueltige = [z for z in zeilen if z.get("fehler") is None]
            if not gueltige:
                break
            beste_zeile = max(gueltige, key=lambda z: z["gewinn_verlust_chf_jahr"])
            aktuelle_max_kapazitaet = max(kapazitaeten)
            if beste_zeile["kapazitaet_kwh"] < aktuelle_max_kapazitaet - 1e-6:
                # Bestes Ergebnis liegt INNERHALB des Rasters -- echtes
                # (inneres) Maximum gefunden, keine weitere Erweiterung noetig.
                break

            aktueller_multiplikator = aktuelle_max_kapazitaet / basis_kwh
            neuer_multiplikator = round(aktueller_multiplikator) + KAPAZITAETS_ERWEITERUNGS_SCHRITT
            neue_kapazitaet = _rund_schoen(basis_kwh * neuer_multiplikator)
            versuch = 0
            while (neue_kapazitaet <= aktuelle_max_kapazitaet or neue_kapazitaet in kapazitaeten) and versuch < 5:
                neuer_multiplikator += KAPAZITAETS_ERWEITERUNGS_SCHRITT
                neue_kapazitaet = _rund_schoen(basis_kwh * neuer_multiplikator)
                versuch += 1
            if neue_kapazitaet <= aktuelle_max_kapazitaet or neue_kapazitaet in kapazitaeten:
                print("  HINWEIS: konnte keine neue, noch nicht getestete (und "
                      "groessere) Kapazitaetsstufe finden -- automatische "
                      "Erweiterung wird hier abgebrochen.")
                break

            erweiterungsrunde += 1
            print(
                f"\n  HINWEIS: das bisher beste Ergebnis liegt bei der GROESSTEN "
                f"getesteten Kapazitaet ({aktuelle_max_kapazitaet:,.0f} kWh) -- das "
                f"wirtschaftliche Optimum koennte also noch groesser sein. "
                f"Automatische Raster-Erweiterung {erweiterungsrunde}/"
                f"{KAPAZITAETS_MAX_ERWEITERUNGEN}: teste zusaetzlich "
                f"{neue_kapazitaet:,.0f} kWh ..."
            )
            kapazitaeten = sorted(set(kapazitaeten) | {neue_kapazitaet})
            zeilen.extend(_punkte_rechnen([(neue_kapazitaet, c) for c in C_RATEN]))
        else:
            print(
                f"\n  HINWEIS: auch nach {KAPAZITAETS_MAX_ERWEITERUNGEN} "
                f"automatischen Raster-Erweiterungen liegt das beste Ergebnis "
                f"weiterhin bei der groessten getesteten Kapazitaet "
                f"({max(kapazitaeten):,.0f} kWh). Das kann bedeuten, dass hier "
                f"tatsaechlich noch groessere Batterien wirtschaftlich sinnvoll "
                f"sind, oder dass die Gewinnkurve in diesem Szenario kein klares "
                f"Maximum hat (z.B. sehr hohes Arbitrage-Potenzial) -- bitte "
                f"Ergebnis pruefen bzw. bei Bedarf KAPAZITAETS_MAX_ERWEITERUNGEN "
                f"erhoehen oder manuell noch groessere Kapazitaeten testen."
            )

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
    Farbe, divergierend = zwei Farben + neutrale Mitte, nie ein Regenbogen).

    NEU (Beat, 22.9.2026, zweite Korrektur): wieder Kapazitaet als Zeile
    (jetzt automatisch aus dem PV-/Lastprofil, siehe baue_raster()), C-Rate
    als Spalte, `leistung_pivot` liefert die daraus resultierende Leistung
    (kW) als Sekundaer-Beschriftung je Zelle -- die Zwischenversion hatte
    das kurzzeitig vertauscht (siehe Modulkopf/Historie)."""
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
    derselben Position schwarz umrandet.

    NEU (Beat, 22.9.2026, zweite Korrektur): Pivot-Index ist wieder
    Kapazitaet (kWh), nicht mehr Leistung (kW) -- siehe _heatmap_seite()."""
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
            "  python battery_sizing.py <Pfad-zu-Inputs.xlsx>"
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
    parser.add_argument(
        "--parallel", type=int, default=SWEEP_PARALLEL_WORKERS_DEFAULT,
        help=(
            "Anzahl Rasterpunkte, die GLEICHZEITIG in separaten Prozessen "
            "gerechnet werden (Default auf diesem Rechner: "
            f"{SWEEP_PARALLEL_WORKERS_DEFAULT}, d.h. alle Kerne bis auf einen). "
            "--parallel 1 schaltet zurueck auf den alten, sequenziellen Ablauf."
        ),
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
    df = sweep(input_lp_path, args.solver, parallel_workers=args.parallel)

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