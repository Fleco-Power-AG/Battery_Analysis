# -*- coding: utf-8 -*-
"""
fleco_kundenreport.py
=====================
Erzeugt aus einem Berechnungsreport (Batterieanalyse) und optional einer
Sizing-Analyse (Sweep) einen standardisierten Kundenbericht im Fleco-Layout.

Aufruf (Spyder / Konsole):
    python fleco_kundenreport.py --bericht inputs/Bericht_Meccano.pdf \
                                 --sweep   inputs/Sweep_Meccano_1.pdf \
                                 --kunde   inputs/kunde_meccano.json \
                                 --out     outputs/Kundenbericht_Meccano.pdf

Inputs
------
--bericht : PDF aus Batterieanalyse (Seiten "Eingabeparameter", "System-Zusammenfassung",
            "Wirtschaftlichkeit", "Eigenverbrauch/Lastspitze", "Kumulierte Erträge")
            ODER eine JSON-Datei mit denselben Schlüsseln wie parse_bericht() liefert.
--sweep   : (optional) Sweep-PDF mit den drei Heatmaps (Gewinn netto, Kapitalverzinsung,
            operativer Gesamtertrag) ODER CSV (Spalten: kapazitaet_kwh;c_rate;leistung_kw;
            netto_chf;roi_pct;operativ_chf), Semikolon-getrennt.
--kunde   : JSON mit Kundendaten, Texten und optionalen Zusatzdaten (Monatsspitzen, Investition).

Abhängigkeiten: pdfplumber, matplotlib, numpy, pandas, reportlab, pillow
"""

from __future__ import annotations

import argparse
import json
import math
import os
import re
import sys
from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.colors import LinearSegmentedColormap, TwoSlopeNorm, Normalize
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch, Circle

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph

HERE = Path(__file__).resolve().parent
ASSETS = HERE / "assets"

# =============================================================================
# 1) Corporate Design
# =============================================================================
C = dict(
    dark="#1D6B2C", green="#3F8F3A", mid="#78B85A", light="#B7DCA3", pale="#EFF6EB",
    ink="#1F2A24", grey="#66756D", grey2="#A3AEA8", line="#DCE3DE", bg="#F6F8F6",
    amber="#E3A72F", amber_pale="#FDF5E3", blue="#3E7CB1", red="#C8553D",
)
MODUL_FARBEN = {
    "SRL": C["dark"],
    "Peak-Shaving": "#5FA84A",
    "Eigenverbrauchsoptimierung": "#A9D18E",
    "Einspeiseoptimierung": C["amber"],
    "Arbitrage": C["blue"],
    "Übrige": C["grey2"],
}
MODUL_LANG = {
    "SRL": "Sekundärregelleistung (SRL)",
    "Peak-Shaving": "Peak-Shaving",
    "Eigenverbrauchsoptimierung": "Eigenverbrauchsoptimierung",
    "Einspeiseoptimierung": "Einspeiseoptimierung",
    "Arbitrage": "Arbitrage",
}

FIRMA_FOOTER = ("Fleco Power AG  |  Technoparkstrasse 2  |  8406 Winterthur  |  "
                "052 209 04 04  |  info@flecopower.ch  |  www.flecopower.ch")


def _register_fonts():
    """Calibri (Windows) bevorzugt, sonst Carlito (metrisch identisch), sonst DejaVu."""
    candidates = [
        ("C:/Windows/Fonts/calibri.ttf", "C:/Windows/Fonts/calibrib.ttf",
         "C:/Windows/Fonts/calibrii.ttf", "Calibri"),
        ("/usr/share/fonts/truetype/crosextra/Carlito-Regular.ttf",
         "/usr/share/fonts/truetype/crosextra/Carlito-Bold.ttf",
         "/usr/share/fonts/truetype/crosextra/Carlito-Italic.ttf", "Carlito"),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
         "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
         "/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf", "DejaVu Sans"),
    ]
    for reg, bold, ital, family in candidates:
        if all(os.path.exists(p) for p in (reg, bold, ital)):
            pdfmetrics.registerFont(TTFont("F", reg))
            pdfmetrics.registerFont(TTFont("F-B", bold))
            pdfmetrics.registerFont(TTFont("F-I", ital))
            pdfmetrics.registerFontFamily("F", normal="F", bold="F-B", italic="F-I", boldItalic="F-B")
            for p in (reg, bold, ital):
                font_manager.fontManager.addfont(p)
            return family
    raise RuntimeError("Keine geeignete Schrift gefunden.")


FONT_FAMILY = _register_fonts()
plt.rcParams.update({
    "font.family": FONT_FAMILY, "font.size": 9, "axes.edgecolor": C["line"],
    "axes.labelcolor": C["grey"], "xtick.color": C["grey"], "ytick.color": C["grey"],
    "axes.spines.top": False, "axes.spines.right": False, "axes.spines.left": False,
    "axes.grid": True, "grid.color": C["line"], "grid.linewidth": 0.6, "axes.axisbelow": True,
    "axes.titleweight": "bold", "axes.titlesize": 10, "axes.titlecolor": C["ink"],
    "legend.frameon": False, "savefig.dpi": 300,
})


# =============================================================================
# 2) Formatierung
# =============================================================================
def num(x, dec=0):
    """Schweizer Zahlenformat mit typografischem Apostroph."""
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return "–"
    s = f"{x:,.{dec}f}".replace(",", "’")
    return s.replace("-", "−")


def sgn(x):
    """Vorzeichenbehaftete Zahl mit typografischem Minus."""
    return (f"{x:+.0f}").replace("-", "−")


def pct(x):
    return "<1 %" if 0 < x < 0.5 else f"{x:.0f} %"


def chf(x, dec=0):
    return f"{num(x, dec)} CHF"


def _f(s: str) -> float:
    s = s.replace("’", "").replace("'", "").replace(",", "").replace("−", "-").strip()
    return float(s)


# =============================================================================
# 3) Parser
# =============================================================================
def _pdf_text(path) -> list[str]:
    import pdfplumber
    with pdfplumber.open(path) as pdf:
        return [p.extract_text() or "" for p in pdf.pages]


def _rx(pattern, text, cast=_f, default=None, group=1):
    m = re.search(pattern, text, flags=re.M)
    if not m:
        return default
    return cast(m.group(group)) if cast else m.group(group)


NUM = r"(-?[\d'’.,]+)"


def parse_bericht(path) -> dict:
    """Liest den Berechnungsreport (PDF oder JSON) in ein flaches Dict."""
    path = Path(path)
    if path.suffix.lower() == ".json":
        return json.loads(path.read_text(encoding="utf-8"))

    pages = _pdf_text(path)
    t = "\n".join(pages)
    b = {}
    b["zeitraum_von"] = _rx(r"Simulierter Zeitraum\s+(\S+)–(\S+)", t, cast=str)
    b["zeitraum_bis"] = _rx(r"Simulierter Zeitraum\s+(\S+)–(\S+)", t, cast=str, group=2)
    b["tarifschema"] = _rx(r"^Schema\s+(\S+)", t, cast=str)
    b["ht_preis"] = _rx(r"HT-Preis\s+(.+?CHF/kWh)", t, cast=str)
    b["nt_preis"] = _rx(r"NT-Preis\s+(.+?CHF/kWh)", t, cast=str)
    b["ht_fenster"] = _rx(r"HT-Zeitfenster\s+(.+)$", t, cast=str)
    b["ht_tage"] = _rx(r"HT-Tage\s+(.+)$", t, cast=str)
    b["pv_schema"] = _rx(r"PV-Schema\s+(\S+)", t, cast=str)
    b["bat_schema"] = _rx(r"Batterie-Schema\s+(\S+)", t, cast=str)
    b["pv_kwp"] = _rx(r"DC-Leistung\s+" + NUM + r"\s*kWp", t)
    b["pv_abschaltung_neg"] = _rx(r"Abschaltung bei Negativpreisen\s+(\S+)", t, cast=str)
    b["bat_kw"] = _rx(r"Batteriespeicher\s+" + NUM + r"\s*kW\s*/", t)
    b["bat_kwh"] = _rx(r"Batteriespeicher\s+[\d'’.,]+\s*kW\s*/\s*" + NUM + r"\s*kWh", t)
    b["rte_pct"] = _rx(r"Rundwirkungsgrad.*?\s" + NUM + r"\s*%", t)
    b["soc_min"] = _rx(r"SOC-Betriebsgrenzen\s+" + NUM + r"\s*%", t)
    b["soc_max"] = _rx(r"SOC-Betriebsgrenzen\s+[\d.]+\s*%\s*–\s*" + NUM + r"\s*%", t)
    b["netz_standort"] = _rx(r"Netz-Grenzen Standort.*?\)\s+(.+)$", t, cast=str)
    b["last_mwh"] = _rx(r"Last \(Jahreslast\)\s+" + NUM + r"\s*MWh", t)
    b["leistungspreis"] = _rx(r"Peak-Shaving\s+aktiv,?\s*\(?" + NUM + r"\s*CHF/kW/Monat", t)
    b["srl_modus"] = _rx(r"Modus\s+(.+)$", t, cast=str)
    b["srl_preise"] = _rx(r"Preise Sekundärregelleistung \(SRL\)\s+(.+)$", t, cast=str)

    e = {}
    for key in ["Eigenverbrauchsoptimierung", "Einspeiseoptimierung", "Peak-Shaving", "Arbitrage", "SRL"]:
        e[key] = _rx(rf"^{re.escape(key)}\s+{NUM}\s*CHF\s*$", t, default=0.0)
    b["einnahmen"] = e
    b["total_einnahmen"] = _rx(r"Total Einnahmen\s+" + NUM + r"\s*CHF", t)
    b["amortisation"] = -_rx(r"^Amortisation\s+" + NUM + r"\s*CHF", t)
    b["kapitalkosten"] = -_rx(r"Kapitalkosten.*?\s" + NUM + r"\s*CHF", t)
    b["kapitalzins_pct"] = _rx(r"Kapitalkosten \(Annahme\s*" + NUM + r"\s*%\)", t)
    b["unterhalt"] = -_rx(r"^Unterhalt\s+" + NUM + r"\s*CHF", t)
    b["degradation"] = -_rx(r"Batterie-Degradation\s+" + NUM + r"\s*CHF", t)
    b["total_ausgaben"] = -_rx(r"Total Ausgaben\s+" + NUM + r"\s*CHF", t)
    b["netto"] = _rx(r"Gewinn / Verlust \(netto\)\s+" + NUM + r"\s*CHF", t)
    b["lcos_rp"] = _rx(NUM + r"\s*Rp/kWh\s+[\d.]+\s*Jahre", t)
    b["amortisationsdauer_j"] = _rx(r"Rp/kWh\s+" + NUM + r"\s*Jahre", t)
    b["kapitalverzinsung_pct"] = _rx(r"Jahre\s+" + NUM + r"\s*%", t)
    b["wert_pv_abschaltung"] = _rx(r"Wert der Abschaltung ohne Batterie\s+" + NUM + r"\s*CHF", t)

    evq = re.findall(r"Eigenverbrauchsquote:\s*" + NUM + r"\s*%", t)
    b["evq_ohne"], b["evq_mit"] = (_f(evq[0]), _f(evq[1])) if len(evq) >= 2 else (None, None)
    nb = re.findall(r"Ø Netzbezug:\s*" + NUM + r"\s*kW\s*·\s*Ø Netzeinspeisung:\s*" + NUM, t)
    if len(nb) >= 2:
        b["netzbezug_ohne"], b["einspeisung_ohne"] = _f(nb[0][0]), _f(nb[0][1])
        b["netzbezug_mit"], b["einspeisung_mit"] = _f(nb[1][0]), _f(nb[1][1])
    m = re.search(r"Ø Ohne\s+" + NUM + r"\s*kW\s*\(Max\s*" + NUM + r"\s*kW\)\s*·\s*Ø Mit\s+" + NUM +
                  r"\s*kW\s*\(Max\s*" + NUM + r"\s*kW\)\s*·\s*Ersparnis:\s*" + NUM, t)
    if m:
        (b["spitze_avg_ohne"], b["spitze_max_ohne"], b["spitze_avg_mit"],
         b["spitze_max_mit"], b["peak_ersparnis"]) = map(_f, m.groups())
    # Kontrollwerte aus der Grafik "Kumulierte Erträge"
    b["kumuliert_total"] = _rx(r"Kumulierter operativer Gesamtertrag per [\d.]+:\s*" + NUM, t)
    kum = {}
    for key in ["Arbitrage", "SRL", "Peak-Shaving", "Eigenverbrauchsoptimierung", "Einspeiseoptimierung"]:
        kum[key] = _rx(rf"{re.escape(key)} \({NUM} CHF\)", t)
    b["kumuliert_module"] = kum
    return b


def parse_sweep(path) -> pd.DataFrame | None:
    """Liest die Sizing-Analyse (PDF mit 3 Heatmaps oder CSV)."""
    if path is None:
        return None
    path = Path(path)
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path, sep=";", encoding="utf-8-sig")

    pages = _pdf_text(path)
    out = {}
    caps = rates = None
    for t in pages:
        caps = [_f(x) for x in re.findall(r"^([\d,]+) kWh$", t, flags=re.M)] or caps
        rr = re.search(r"^((?:[\d.]+C\s*)+)$", t, flags=re.M)
        rates = [_f(x) for x in re.findall(r"([\d.]+)C", rr.group(1))] if rr else rates
        kw = [_f(x) for x in re.findall(r"\(([\d,]+) kW\)", t)]
        if "Gewinn/Verlust netto" in t:
            out["netto_chf"] = [_f(x) for x in re.findall(r"(-?[\d,]+) CHF/Jahr", t)]
            out["leistung_kw"] = kw
        elif "Kapitalverzinsung" in t:
            out["roi_pct"] = [_f(x) for x in re.findall(r"(-?\d+)%", t)]
        elif "Operativer Gesamtertrag" in t:
            out["operativ_chf"] = [_f(x) for x in re.findall(r"(-?[\d,]+) CHF/Jahr", t)]
    n = len(caps) * len(rates)
    rows = []
    for i in range(n):
        rows.append(dict(
            kapazitaet_kwh=caps[i // len(rates)], c_rate=rates[i % len(rates)],
            leistung_kw=out["leistung_kw"][i], netto_chf=out["netto_chf"][i],
            roi_pct=out["roi_pct"][i], operativ_chf=out["operativ_chf"][i]))
    return pd.DataFrame(rows)


# =============================================================================
# 4) Analytik
# =============================================================================
@dataclass
class Kennzahlen:
    investition: float
    investition_abgeleitet: bool
    lebensdauer_j: float | None
    opex: float
    finanzierung: float
    op_ertrag: float
    module: dict
    anteile: dict
    srl: float
    payback_ohne_srl: float
    roi_ohne_srl: float
    netto_ohne_srl: float
    breakeven_srl_pct: float | None


def kennzahlen(b: dict, kunde: dict) -> Kennzahlen:
    op = b["total_einnahmen"]
    opex = b["unterhalt"] + b["degradation"]
    fin = b["amortisation"] + b["kapitalkosten"]
    inv = kunde.get("investition_chf")
    abgeleitet = False
    if not inv:
        # Kapitalkosten = Zins x Investition  ->  Investition rückrechnen
        inv = b["kapitalkosten"] / (b["kapitalzins_pct"] / 100)
        abgeleitet = True
    # Lebensdauer aus Annuität (Amortisation + Zins) rückrechnen
    r = b["kapitalzins_pct"] / 100
    ann = fin / inv
    n = None
    if ann > r:
        n = -math.log(1 - r / ann) / math.log(1 + r)
    module = {k: v for k, v in b["einnahmen"].items()}
    anteile = {k: v / op * 100 for k, v in module.items()}
    srl = module.get("SRL", 0.0)
    cf_ohne = op - srl - opex
    return Kennzahlen(
        investition=inv, investition_abgeleitet=abgeleitet, lebensdauer_j=n,
        opex=opex, finanzierung=fin, op_ertrag=op, module=module, anteile=anteile, srl=srl,
        payback_ohne_srl=inv / cf_ohne if cf_ohne > 0 else float("inf"),
        roi_ohne_srl=cf_ohne / inv * 100, netto_ohne_srl=b["netto"] - srl,
        breakeven_srl_pct=(-b["netto"] / srl * 100) if srl > 0 else None,
    )


def szenario(b, k: Kennzahlen, delta_srl_pct):
    """Linear: SRL-Erlös skaliert mit Preisniveau, Mengen/Kosten unverändert."""
    d = delta_srl_pct / 100
    netto = b["netto"] + k.srl * d
    cf = k.op_ertrag + k.srl * d - k.opex
    return netto, (k.investition / cf if cf > 0 else float("inf")), cf / k.investition * 100


def sizing(df: pd.DataFrame, ziel_roi: float) -> dict:
    df = df.copy()
    df["label"] = df.apply(lambda r: f"{num(r.kapazitaet_kwh)} kWh / {num(r.leistung_kw)} kW", axis=1)
    best_net = df.loc[df.netto_chf.idxmax()]
    best_roi = df.sort_values(["roi_pct", "netto_chf"], ascending=False).iloc[0]
    kand = df[df.roi_pct >= ziel_roi]
    ausgewogen = kand.loc[kand.netto_chf.idxmax()] if len(kand) else best_roi
    caps = sorted(df.kapazitaet_kwh.unique())
    rates = sorted(df.c_rate.unique())
    # dominiert eine C-Rate in allen Zeilen?
    dom = df.loc[df.groupby("kapazitaet_kwh").netto_chf.idxmax(), "c_rate"].unique()
    return dict(
        df=df, caps=caps, rates=rates, best_net=best_net, best_roi=best_roi,
        ausgewogen=ausgewogen, ziel_roi=ziel_roi,
        dominante_crate=dom[0] if len(dom) == 1 else None,
        rand_optimum=bool(best_net.kapazitaet_kwh == max(caps)),
        negative=df[df.netto_chf < 0],
    )


def detail_konsistent(b, sz, tol=0.03) -> bool:
    """Liegt die Detailvariante (Bericht) auf der Sweep-Kurve? Sonst nicht im Kundenbericht plotten."""
    df = sz["df"]
    sub = df[np.isclose(df.c_rate, b["bat_kw"] / b["bat_kwh"])].sort_values("kapazitaet_kwh")
    if not len(sub) or not (sub.kapazitaet_kwh.min() <= b["bat_kwh"] <= sub.kapazitaet_kwh.max()):
        return False
    net_i = np.interp(b["bat_kwh"], sub.kapazitaet_kwh, sub.netto_chf)
    op_i = np.interp(b["bat_kwh"], sub.kapazitaet_kwh, sub.operativ_chf)
    return abs(net_i - b["netto"]) <= tol * abs(net_i) and abs(op_i - b["total_einnahmen"]) <= tol * op_i


def plausibilitaet(b, k: Kennzahlen, sz: dict | None) -> list[str]:
    """Interne Checks — erscheinen NICHT im Kundenbericht, nur in Konsole/QA-Datei."""
    msgs = []

    def chk(name, a, bb, tol=1.0):
        ok = abs(a - bb) <= tol
        dec = 2 if max(abs(a), abs(bb)) < 100 else 0
        msgs.append(f"[{'OK ' if ok else '!! '}] {name}: {num(a, dec)} vs. {num(bb, dec)} (Diff {num(a - bb, dec)})")

    chk("Summe Module = Total Einnahmen", sum(k.module.values()), b["total_einnahmen"])
    chk("Einnahmen − Ausgaben = Netto", b["total_einnahmen"] - b["total_ausgaben"], b["netto"])
    if b.get("kumuliert_total"):
        chk("Total Einnahmen = Kumuliert (Grafik S.4)", b["total_einnahmen"], b["kumuliert_total"])
        for m, v in b["kumuliert_module"].items():
            if v is not None:
                chk(f"Modul {m} Tabelle vs. Grafik", k.module.get(m, 0), v)
    chk("Amortisationsdauer (Inv/(Ertrag−OPEX))", k.investition / (k.op_ertrag - k.opex),
        b["amortisationsdauer_j"], tol=0.06)
    chk("Kapitalverzinsung ((Ertrag−OPEX)/Inv)", (k.op_ertrag - k.opex) / k.investition * 100,
        b["kapitalverzinsung_pct"], tol=0.06)
    if b.get("spitze_avg_ohne") is not None:
        est = (b["spitze_avg_ohne"] - b["spitze_avg_mit"]) * 12 * b["leistungspreis"]
        chk("Peak-Shaving ≈ ΔØSpitze×12×Leistungspreis (gerundete Ø)", est, b["einnahmen"]["Peak-Shaving"], tol=150)
    if k.investition_abgeleitet:
        msgs.append(f"[i  ] Investition aus Kapitalkosten abgeleitet: {chf(k.investition)} "
                    f"({num(k.investition / b['bat_kwh'])} CHF/kWh), Lebensdauer ≈ {num(k.lebensdauer_j, 1)} J. "
                    f"-> besser 'investition_chf' im Kunden-JSON setzen.")
    if sz is not None:
        df = sz["df"]
        cr = b["bat_kw"] / b["bat_kwh"]
        sub = df[np.isclose(df.c_rate, cr)].sort_values("kapazitaet_kwh")
        if len(sub) and sub.kapazitaet_kwh.min() <= b["bat_kwh"] <= sub.kapazitaet_kwh.max():
            op_i = np.interp(b["bat_kwh"], sub.kapazitaet_kwh, sub.operativ_chf)
            net_i = np.interp(b["bat_kwh"], sub.kapazitaet_kwh, sub.netto_chf)
            chk("Sweep (interpoliert) vs. Bericht: operativer Ertrag", op_i, b["total_einnahmen"], tol=0.03 * op_i)
            chk("Sweep (interpoliert) vs. Bericht: Netto (Kostenbasis!)", net_i, b["netto"], tol=0.03 * abs(net_i))
            if not (df.kapazitaet_kwh == b["bat_kwh"]).any():
                msgs.append(f"[i  ] Bericht-Variante {num(b['bat_kwh'])} kWh ist nicht im Sweep-Raster enthalten.")
        if sz["rand_optimum"]:
            msgs.append("[!! ] Max. Nettogewinn liegt am Rand des Sweep-Rasters -> Raster erweitern "
                        "oder Begrenzung (Netzanschluss/Präqualifikation) begründen.")
    return msgs


# =============================================================================
# 5) Grafiken (matplotlib -> PNG)
# =============================================================================
def _save(fig, path):
    fig.savefig(path, bbox_inches="tight", pad_inches=0.04, transparent=True)
    plt.close(fig)
    return path


def module_sortiert(k: Kennzahlen, schwelle_pct=2.0):
    """Module absteigend; Kleinstmodule (< schwelle) zu 'Übrige' zusammenfassen."""
    items = sorted(k.module.items(), key=lambda x: -x[1])
    gross = [(n, v) for n, v in items if k.anteile[n] >= schwelle_pct]
    klein = [(n, v) for n, v in items if k.anteile[n] < schwelle_pct and v != 0]
    if klein:
        gross.append(("Übrige", sum(v for _, v in klein)))
    return gross, [n for n, _ in klein]


def chart_donut(k: Kennzahlen, path, b):
    items, klein = module_sortiert(k)
    fig, ax = plt.subplots(figsize=(7.2, 2.15))
    ax.set_position([0.0, 0.0, 0.34, 1.0])
    vals = [v for _, v in items]
    cols = [MODUL_FARBEN[n] for n, _ in items]
    ax.pie(vals, colors=cols, startangle=90, counterclock=False,
           wedgeprops=dict(width=0.30, edgecolor="white", linewidth=2))
    ax.text(0, 0.10, num(k.op_ertrag), ha="center", va="center", fontsize=17, fontweight="bold", color=C["ink"])
    ax.text(0, -0.20, "CHF / Jahr", ha="center", va="center", fontsize=8.5, color=C["grey"])
    ax.set_aspect("equal")
    # Legende als Tabelle rechts
    lx = fig.add_axes([0.40, 0.04, 0.59, 0.92])
    lx.axis("off")
    lx.set_xlim(0, 1)
    lx.set_ylim(0, 1)
    step = 1 / max(len(items), 4)
    for i, (n, v) in enumerate(items):
        y = 1 - (i + 0.5) * step
        lx.add_patch(FancyBboxPatch((0, y - 0.035), 0.03, 0.07, boxstyle="round,pad=0,rounding_size=0.012",
                                    color=MODUL_FARBEN[n], transform=lx.transAxes))
        name = MODUL_LANG.get(n, n)
        if n == "Übrige":
            name = "Übrige (" + ", ".join(klein) + ")"
        lx.text(0.06, y, name, va="center", fontsize=9, color=C["ink"])
        lx.text(0.80, y, num(v), va="center", ha="right", fontsize=9, fontweight="bold", color=C["ink"])
        lx.text(0.99, y, pct(k.anteile.get(n, v / k.op_ertrag * 100)), va="center", ha="right",
                fontsize=9, color=C["grey"])
        lx.plot([0, 1], [y - step / 2, y - step / 2], color=C["line"], lw=0.6)
    return _save(fig, path)


def chart_waterfall(b, k: Kennzahlen, path):
    items, _ = module_sortiert(k)
    kosten = [("Unterhalt", -b["unterhalt"]), ("Degradation", -b["degradation"]),
              ("Kapitalkosten", -b["kapitalkosten"]), ("Amortisation", -b["amortisation"])]
    labels, starts, heights, cols = [], [], [], []
    cum = 0
    for n, v in items:
        labels.append({"SRL": "SRL", "Eigenverbrauchsoptimierung": "Eigen-\nverbrauch",
                       "Peak-Shaving": "Peak-\nShaving", "Übrige": "Übrige"}.get(n, n))
        starts.append(cum); heights.append(v); cols.append(MODUL_FARBEN[n]); cum += v
    cum = b["total_einnahmen"]  # gemeldeter Wert (Rundungsdifferenzen der Module)
    labels.append("Operativer\nErtrag"); starts.append(0); heights.append(cum); cols.append(C["ink"])
    for n, v in kosten:
        labels.append(n.replace("Kapitalkosten", "Kapital-\nkosten").replace("Amortisation", "Amorti-\nsation")
                      .replace("Degradation", "Degra-\ndation"))
        starts.append(cum + v); heights.append(-v); cols.append(C["red"]); cum += v
    cum = b["netto"]
    labels.append("Netto-\ngewinn"); starts.append(0); heights.append(cum)
    cols.append(C["dark"] if cum >= 0 else C["red"])

    fig, ax = plt.subplots(figsize=(7.0, 2.9))
    x = np.arange(len(labels))
    for i in range(len(labels)):
        alpha = 0.55 if cols[i] == C["red"] else 1
        ax.bar(x[i], heights[i], bottom=starts[i], color=cols[i], width=0.62, alpha=alpha,
               edgecolor="none", zorder=3)
        top = starts[i] + max(heights[i], 0) if cols[i] != C["red"] else starts[i] + heights[i]
        val = heights[i] if cols[i] != C["red"] else -heights[i]
        ax.text(x[i], top + cum * 0.015 + 300, num(val), ha="center", va="bottom", fontsize=7.6,
                color=C["red"] if val < 0 else C["ink"], fontweight="bold" if i in (len(items), len(labels) - 1) else None)
        if i < len(labels) - 1:
            lvl = starts[i] + heights[i] if cols[i] != C["red"] else starts[i]
            if i == len(items):
                lvl = heights[i]
            ax.plot([x[i] + 0.31, x[i + 1] - 0.31], [lvl, lvl], color=C["grey2"], lw=0.7, ls=":", zorder=2)
    ax.set_xticks(x, labels, fontsize=7.6)
    ax.tick_params(axis="x", length=0)
    ax.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, _: num(v)))
    ax.set_ylabel("CHF / Jahr")
    ax.axvspan(len(items) + 0.5, len(labels) - 1.5, color=C["bg"], zorder=0)
    ax.set_ylim(0, max(starts[i] + heights[i] for i in range(len(labels))) * 1.12)
    return _save(fig, path)


def chart_sensitivitaet(b, k: Kennzahlen, path, bereich=(-50, 50)):
    lo, hi = bereich
    d = np.linspace(lo, hi, int(hi - lo) + 1)
    net = np.array([szenario(b, k, x)[0] for x in d])
    pb = np.array([szenario(b, k, x)[1] for x in d])
    fig, (a1, a2) = plt.subplots(1, 2, figsize=(7.2, 2.6), gridspec_kw=dict(wspace=0.32))
    for ax in (a1, a2):
        ax.axvspan(lo, hi, color=C["bg"], zorder=0)
        ax.axvline(0, color=C["grey2"], lw=0.8, ls=":")
        ax.set_xlim(lo, hi)
        ax.set_xticks(np.linspace(lo, hi, 5))
        ax.xaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, _: f"{sgn(v)} %" if v else "0 %"))
        ax.set_xlabel("SRL-Erlöse ggü. 2025")
    a1.fill_between(d, net, 0, where=net >= 0, color=C["light"], alpha=0.55, lw=0)
    a1.plot(d, net, color=C["dark"], lw=2)
    a1.scatter([0], [b["netto"]], color=C["dark"], zorder=5, s=28)
    a1.annotate("Basis 2025", (0, b["netto"]), xytext=(8, -12), textcoords="offset points", fontsize=7.5,
                color=C["dark"])
    a1.set_ylim(min(0, net.min() * 1.1), net.max() * 1.12)
    a1.set_title("Nettogewinn pro Jahr [CHF]", loc="left")
    a1.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, _: num(v)))

    a2.plot(d, pb, color=C["dark"], lw=2)
    a2.scatter([0], [b["amortisationsdauer_j"]], color=C["dark"], zorder=5, s=28)
    a2.annotate("Basis 2025", (0, b["amortisationsdauer_j"]), xytext=(8, 6), textcoords="offset points",
                fontsize=7.5, color=C["dark"])
    a2.set_ylim(0, np.nanmax(pb) * 1.2)
    a2.set_title("Amortisationsdauer [Jahre]", loc="left")
    return _save(fig, path)


def chart_lastspitzen(ohne, mit, path):
    mon = ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"]
    x = np.arange(12)
    fig, ax = plt.subplots(figsize=(7.2, 2.8))
    ax.bar(x - 0.19, ohne, 0.36, color=C["grey2"], label="ohne Batterie", zorder=3)
    ax.bar(x + 0.19, mit, 0.36, color=C["dark"], label="mit Batterie", zorder=3)
    for val, col, lab in [(np.mean(ohne), C["grey"], "Ø ohne"), (np.mean(mit), C["dark"], "Ø mit")]:
        ax.axhline(val, color=col, lw=0.9, ls="--", zorder=4)
        ax.text(11.5, val + 1.5, f"{lab}\n{val:.0f} kW", va="bottom", fontsize=7.2, color=col, linespacing=1.0)
    ax.set_xticks(x, mon)
    ax.tick_params(axis="x", length=0)
    ax.set_ylabel("kW")
    ax.set_ylim(0, max(ohne) * 1.28)
    ax.legend(loc="upper center", ncol=2, fontsize=8)
    ax.set_xlim(-0.6, 12.5)
    return _save(fig, path)


def chart_schema(b, path):
    fig, ax = plt.subplots(figsize=(7.2, 2.9))
    ax.set_xlim(0, 10); ax.set_ylim(0, 4); ax.axis("off")

    def node(x, y, titel, wert, farbe, w=2.3, h=1.05):
        ax.add_patch(FancyBboxPatch((x - w / 2, y - h / 2), w, h, boxstyle="round,pad=0.02,rounding_size=0.18",
                                    fc="white", ec=C["line"], lw=1.0, zorder=3))
        ax.add_patch(FancyBboxPatch((x - w / 2, y - h / 2), 0.14, h, boxstyle="round,pad=0,rounding_size=0.07",
                                    fc=farbe, ec="none", zorder=4))
        ax.text(x - w / 2 + 0.32, y + 0.17, titel, fontsize=9.5, fontweight="bold", color=C["ink"], va="center", zorder=5)
        ax.text(x - w / 2 + 0.32, y - 0.20, wert, fontsize=8.5, color=C["grey"], va="center", zorder=5)

    def arr(p1, p2, farbe, txt=None, both=False, rad=0.0, off=(0, 0.16)):
        ax.add_patch(FancyArrowPatch(p1, p2, arrowstyle="<|-|>" if both else "-|>", mutation_scale=11,
                                     color=farbe, lw=1.8, connectionstyle=f"arc3,rad={rad}", zorder=2))
        if txt:
            ax.text((p1[0] + p2[0]) / 2 + off[0], (p1[1] + p2[1]) / 2 + off[1], txt, fontsize=7.5,
                    color=farbe, ha="center", zorder=6,
                    bbox=dict(fc="white", ec="none", pad=1.2))

    node(5.0, 3.35, "PV-Anlage", f"{num(b['pv_kwp'])} kWp", C["amber"])
    node(5.0, 0.65, "Batteriespeicher", f"{num(b['bat_kw'])} kW / {num(b['bat_kwh'])} kWh", C["dark"])
    node(8.75, 2.0, "Verbraucher", f"{num(b['last_mwh'], 1)} MWh/Jahr", C["mid"], w=2.3)
    node(1.25, 2.0, "Stromnetz", "Bezug & Rücklieferung", C["blue"], w=2.3)
    arr((6.2, 3.2), (7.75, 2.55), C["amber"], "Direktverbrauch", off=(0.25, 0.1))
    arr((5.0, 2.8), (5.0, 1.2), C["amber"], "Laden", off=(0.42, 0.5))
    arr((6.2, 0.8), (7.75, 1.45), C["dark"], "Eigenverbrauch\n& Peak-Shaving", off=(0.35, -0.35))
    arr((2.25, 1.45), (3.8, 0.8), C["dark"], "SRL · Arbitrage", both=True, off=(-0.35, -0.3))
    arr((3.8, 3.2), (2.25, 2.55), C["amber"], "Einspeisung", off=(-0.25, 0.1))
    ax.add_patch(FancyArrowPatch((2.45, 2.0), (7.55, 2.0), arrowstyle="-|>", mutation_scale=11,
                                 color=C["blue"], lw=1.4, ls=(0, (4, 3)), zorder=1))
    ax.text(3.35, 2.12, "Netzbezug", fontsize=7.5, color=C["blue"], ha="center",
            bbox=dict(fc="white", ec="none", pad=1.2), zorder=6)
    return _save(fig, path)


def _cell_grid(ax, df, sz, wert, fmt, cmap, norm, titel, highlight):
    caps, rates = sz["caps"], sz["rates"]
    ax.set_xlim(-0.5, len(rates) - 0.5); ax.set_ylim(len(caps) - 0.5, -0.5)
    ax.grid(False)
    for spine in ax.spines.values():
        spine.set_visible(False)
    for _, r in df.iterrows():
        i, j = caps.index(r.kapazitaet_kwh), rates.index(r.c_rate)
        v = r[wert]
        fc = cmap(norm(v))
        ax.add_patch(FancyBboxPatch((j - 0.46, i - 0.43), 0.92, 0.86,
                                    boxstyle="round,pad=0,rounding_size=0.12", fc=fc, ec="white", lw=1))
        lum = 0.299 * fc[0] + 0.587 * fc[1] + 0.114 * fc[2]
        tc = "white" if lum < 0.55 else C["ink"]
        ax.text(j, i - 0.09, fmt(v), ha="center", va="center", fontsize=8.6, fontweight="bold", color=tc)
        ax.text(j, i + 0.22, f"{num(r.leistung_kw)} kW", ha="center", va="center", fontsize=6.6, color=tc, alpha=0.85)
        for tag, row in highlight.items():
            if row is not None and row.kapazitaet_kwh == r.kapazitaet_kwh and row.c_rate == r.c_rate:
                ax.add_patch(FancyBboxPatch((j - 0.46, i - 0.43), 0.92, 0.86,
                                            boxstyle="round,pad=0,rounding_size=0.12", fc="none",
                                            ec=C["ink"], lw=1.6))
                ax.add_patch(Circle((j + 0.36, i - 0.31), 0.11, fc=C["ink"], ec="white", lw=0.8, zorder=6))
                ax.text(j + 0.36, i - 0.31, tag, ha="center", va="center", fontsize=6.5, color="white",
                        fontweight="bold", zorder=7)
    ax.set_xticks(range(len(rates)), [f"{r:g}C" for r in rates])
    ax.set_yticks(range(len(caps)), [f"{num(c)} kWh" for c in caps])
    ax.tick_params(length=0)
    ax.xaxis.tick_top()
    ax.set_title(titel, loc="left", pad=18)


def chart_heatmaps(sz, path):
    df = sz["df"]
    hl = {"R": sz["best_roi"], "A": sz["ausgewogen"], "E": sz["best_net"]}
    fig, (a1, a2) = plt.subplots(1, 2, figsize=(7.2, 4.4), gridspec_kw=dict(wspace=0.38))
    vmax = max(abs(df.netto_chf.max()), abs(df.netto_chf.min()))
    cmap_net = LinearSegmentedColormap.from_list("net", [C["red"], "#F4E3DE", "#F4F7F3", C["light"], C["dark"]])
    _cell_grid(a1, df, sz, "netto_chf", lambda v: "0.0k" if abs(v) < 50 else f"{num(v / 1000, 1)}k", cmap_net,
               TwoSlopeNorm(vmin=-vmax, vcenter=0, vmax=vmax), "Nettogewinn [CHF/Jahr]", hl)
    cmap_roi = LinearSegmentedColormap.from_list("roi", ["#F4F7F3", C["light"], C["dark"]])
    _cell_grid(a2, df, sz, "roi_pct", lambda v: f"{v:.0f} %", cmap_roi,
               Normalize(vmin=min(0, df.roi_pct.min()), vmax=df.roi_pct.max()), "Kapitalverzinsung [%]", hl)
    a2.set_yticklabels([])
    return _save(fig, path)


def chart_tradeoff(sz, b, path):
    """Kapitalverzinsung vs. Nettogewinn. Im Diagramm nur Badges R/A/E, die Konfigurationen
    stehen in der Legendenzeile unter dem Diagramm (seite_dimensionierung)."""
    df = sz["df"]
    fig, ax = plt.subplots(figsize=(7.2, 3.0))
    farben = {r: c for r, c in zip(sz["rates"], [C["grey2"], C["mid"], C["dark"], C["blue"]])}
    for r in sz["rates"]:
        s = df[df.c_rate == r].sort_values("kapazitaet_kwh")
        ax.plot(s.netto_chf / 1000, s.roi_pct, color=farben[r], lw=1.4, zorder=2)
        ax.scatter(s.netto_chf / 1000, s.roi_pct, s=18 + s.kapazitaet_kwh / 9, color=farben[r],
                   edgecolor="white", lw=0.8, zorder=3, label=f"{r:g}C", alpha=0.85)
    for tag, row in {"R": sz["best_roi"], "A": sz["ausgewogen"], "E": sz["best_net"]}.items():
        px, py = row.netto_chf / 1000, row.roi_pct
        ax.scatter([px], [py], s=700, color=C["amber"], alpha=0.25, lw=0, zorder=1)            # Halo
        ax.scatter([px], [py], s=260, facecolor="none", edgecolor=C["ink"], lw=1.8, zorder=6)  # Ring
        ax.annotate(tag, (px, py), xytext=(0, 17), textcoords="offset points", ha="center", va="center",
                    fontsize=8, fontweight="bold", color="white", zorder=8,
                    bbox=dict(boxstyle="circle,pad=0.3", fc=C["ink"], ec="white", lw=0.8),
                    arrowprops=dict(arrowstyle="-", color=C["ink"], lw=0.9, shrinkA=6, shrinkB=8))
    if sz.get("detail_konsistent"):
        ax.scatter([b["netto"] / 1000], [b["kapitalverzinsung_pct"]], marker="D", s=46, color=C["amber"],
                   edgecolor="white", zorder=5, label="Detailanalyse")
    ax.axvline(0, color=C["grey"], lw=0.8)
    ax.axhline(sz["ziel_roi"], color=C["grey2"], lw=0.8, ls="--", label=f"Zielrendite {sz['ziel_roi']:g} %")
    ax.set_ylim(min(0, df.roi_pct.min() - 1), df.roi_pct.max() * 1.18 + 1)
    ax.set_xlabel("Nettogewinn [1’000 CHF/Jahr]")
    ax.set_ylabel("Kapitalverzinsung [%]")
    ax.plot([], [], ls="none", label="Punktgrösse = Kapazität")
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.2), ncol=6, fontsize=7.5, markerscale=0.7,
              handletextpad=0.4, columnspacing=1.4)
    return _save(fig, path)


# =============================================================================
# 6) PDF-Layout (reportlab)
# =============================================================================
W, H = A4
M = 18 * mm  # Seitenrand
CW = W - 2 * M  # Inhaltsbreite


def S(size=9.5, color=C["ink"], lead=None, font="F", align=0):
    return ParagraphStyle("s", fontName=font, fontSize=size, leading=lead or size * 1.38,
                          textColor=HexColor(color), alignment=align)


_NBSP = re.compile(r"(\d) (%|CHF|kWh|kWp|kW|MWh|Rp|Jahren|Jahre|Prozentpunkte)")


def para(c, text, x, y_top, w, style=None):
    text = _NBSP.sub(lambda m: m.group(1) + "&nbsp;" + m.group(2), text)
    p = Paragraph(text, style or S())
    _, h = p.wrap(w, 1000)
    p.drawOn(c, x, y_top - h)
    return h


def rrect(c, x, y, w, h, r, fill=None, stroke=None, lw=0.8):
    c.saveState()
    if fill:
        c.setFillColor(HexColor(fill))
    if stroke:
        c.setStrokeColor(HexColor(stroke)); c.setLineWidth(lw)
    c.roundRect(x, y, w, h, r, fill=1 if fill else 0, stroke=1 if stroke else 0)
    c.restoreState()


def img(c, path, x, y_top, w):
    ir = ImageReader(str(path))
    iw, ih = ir.getSize()
    h = w * ih / iw
    c.drawImage(ir, x, y_top - h, w, h, mask="auto")
    return h


class Doc:
    def __init__(self, out, kunde, b):
        self.c = canvas.Canvas(out if not isinstance(out, (str, Path)) else str(out), pagesize=A4)
        self.kunde, self.b = kunde, b
        proj = kunde["projekt"]["bezeichnung"]
        self.c.setTitle(f"Wirtschaftlichkeitsanalyse Batteriespeicher – {proj}")
        self.c.setAuthor("Fleco Power AG")
        self.page = 0
        self.total = None
        self.kapitel = 0

    def header(self, nr, titel):
        self.kapitel = nr
        c = self.c
        logo = ASSETS / "fleco_logo_rund.png"
        if logo.exists():
            c.drawImage(ImageReader(str(logo)), M, H - 21 * mm, 11 * mm, 11 * mm, mask="auto")
        c.setFont("F", 8); c.setFillColor(HexColor(C["grey"]))
        c.drawRightString(W - M, H - 15.5 * mm,
                          "Wirtschaftlichkeitsanalyse Batteriespeicher"
                          + (f"  ·  Projekt {projekt(self.kunde)}" if projekt(self.kunde) else ""))
        c.setStrokeColor(HexColor(C["line"])); c.setLineWidth(0.6)
        c.line(M, H - 24 * mm, W - M, H - 24 * mm)
        # Kapiteltitel
        y = H - 38 * mm
        rrect(c, M, y - 1.5 * mm, 11 * mm, 8 * mm, 2 * mm, fill=C["dark"])
        c.setFont("F-B", 11); c.setFillColor(white)
        c.drawCentredString(M + 5.5 * mm, y + 0.9 * mm, f"{nr:02d}")
        c.setFont("F-B", 19); c.setFillColor(HexColor(C["dark"]))
        c.drawString(M + 15 * mm, y, titel)
        return H - 46 * mm

    def footer(self):
        c = self.c
        c.setStrokeColor(HexColor(C["line"])); c.setLineWidth(0.6)
        c.line(M, 14 * mm, W - M, 14 * mm)
        c.setFont("F", 7); c.setFillColor(HexColor(C["grey"]))
        c.drawString(M, 9.5 * mm, FIRMA_FOOTER)
        c.drawRightString(W - M, 9.5 * mm, f"Seite {self.page} / {self.total}")

    def new_page(self):
        if self.page:
            self.c.showPage()
        self.page += 1

    def h2(self, text, y, x=M):
        self.c.setFont("F-B", 11.5); self.c.setFillColor(HexColor(C["ink"]))
        self.c.drawString(x, y - 4 * mm, text)
        return y - 7.5 * mm

    def kpi(self, x, y_top, w, h, wert, einheit, label, farbe=C["dark"], hint=None, fill=C["pale"]):
        c = self.c
        rrect(c, x, y_top - h, w, h, 3 * mm, fill=fill)
        rrect(c, x, y_top - h, 1.4 * mm, h, 0.7 * mm, fill=farbe)
        c.setFont("F", 8); c.setFillColor(HexColor(C["grey"]))
        c.drawString(x + 5 * mm, y_top - 6.5 * mm, label.upper())
        c.setFont("F-B", 20); c.setFillColor(HexColor(farbe))
        c.drawString(x + 5 * mm, y_top - 15.5 * mm, wert)
        tw = pdfmetrics.stringWidth(wert, "F-B", 20)
        c.setFont("F", 9); c.setFillColor(HexColor(C["grey"]))
        c.drawString(x + 6 * mm + tw, y_top - 15.5 * mm, einheit)
        if hint:
            c.setFont("F", 7.5); c.drawString(x + 5 * mm, y_top - h + 4 * mm, hint)

    def vergleich(self, x, y_top, w, h, label, ohne, mit, einheit, besser_hoch):
        c = self.c
        rrect(c, x, y_top - h, w, h, 3 * mm, fill="#FFFFFF", stroke=C["line"])
        c.setFont("F", 8); c.setFillColor(HexColor(C["grey"]))
        c.drawString(x + 4 * mm, y_top - 6 * mm, label.upper())
        c.setFont("F", 13); c.setFillColor(HexColor(C["grey2"]))
        c.drawString(x + 4 * mm, y_top - 14 * mm, f"{ohne}")
        tw = pdfmetrics.stringWidth(f"{ohne}", "F", 13)
        c.setFont("F-B", 13); c.setFillColor(HexColor(C["mid"]))
        c.drawString(x + 5.5 * mm + tw, y_top - 14 * mm, "→")
        tw2 = pdfmetrics.stringWidth("→", "F-B", 13)
        c.setFont("F-B", 17); c.setFillColor(HexColor(C["dark"]))
        c.drawString(x + 7.5 * mm + tw + tw2, y_top - 14 * mm, f"{mit}")
        c.setFont("F", 8); c.setFillColor(HexColor(C["grey"]))
        c.drawString(x + 4 * mm, y_top - h + 3.5 * mm, f"ohne → mit Batterie, in {einheit}")

    def tabelle(self, x, y_top, w, rows, col1=0.55, size=9, header=None):
        """rows: list of (label, value) oder ('__h__', title) für Zwischentitel."""
        c = self.c
        y = y_top
        rh = 6.4 * mm
        for lab, val in rows:
            if lab in ("__qh__", "__q__"):
                # val = (label, [4 Werte]) -> Quartalsspalten
                txt, werte = val
                hdr = lab == "__qh__"
                c.setFont("F-B" if hdr else "F", size - (1 if hdr else 0))
                c.setFillColor(HexColor(C["grey"] if hdr else C["ink"]))
                c.drawString(x + 1 * mm, y - 4.5 * mm, txt)
                x0, sw = x + w * 0.50, w * 0.50 / 4
                for i, v in enumerate(werte):
                    c.drawRightString(x0 + (i + 1) * sw - 1 * mm, y - 4.5 * mm, v)
                y -= rh
                c.setStrokeColor(HexColor(C["line"])); c.setLineWidth(0.5); c.line(x, y + 0.6 * mm, x + w, y + 0.6 * mm)
                continue
            if lab == "__h__":
                c.setFont("F-B", 9); c.setFillColor(HexColor(C["dark"]))
                c.drawString(x, y - 4.6 * mm, val)
                y -= rh
                c.setStrokeColor(HexColor(C["dark"])); c.setLineWidth(0.8); c.line(x, y + 0.6 * mm, x + w, y + 0.6 * mm)
                continue
            bold = lab.startswith("**")
            lab = lab.strip("*")
            c.setFont("F-B" if bold else "F", size); c.setFillColor(HexColor(C["ink"]))
            c.drawString(x + 1 * mm, y - 4.5 * mm, lab)
            c.drawRightString(x + w - 1 * mm, y - 4.5 * mm, val)
            y -= rh
            c.setStrokeColor(HexColor(C["line"])); c.setLineWidth(0.5); c.line(x, y + 0.6 * mm, x + w, y + 0.6 * mm)
        return y_top - y

    def callout(self, x, y_top, w, text, farbe=C["amber"], fill=C["amber_pale"], titel=None):
        st = S(9, C["ink"])
        inner = (f"<b>{titel}</b><br/>" if titel else "") + text
        p = Paragraph(inner, st)
        _, h = p.wrap(w - 10 * mm, 1000)
        rrect(self.c, x, y_top - h - 7 * mm, w, h + 7 * mm, 3 * mm, fill=fill)
        rrect(self.c, x, y_top - h - 7 * mm, 1.4 * mm, h + 7 * mm, 0.7 * mm, fill=farbe)
        p.drawOn(self.c, x + 6 * mm, y_top - 3.5 * mm - h)
        return h + 7 * mm


# =============================================================================
# 7) Texte
# =============================================================================
# Platzhalter aus der Vorlage (kunde_vorlage.json). Nicht ausgefüllt = fehlt.
VORLAGE_PLATZHALTER = {"firma auftraggeber", "vorname name", "strasse nr.", "plz ort", "projektname",
                       "strasse nr., plz ort", "tt.mm.jjjj", "standort", "funktion"}


def bereinige_kundendaten(o):
    """Entfernt eckige Klammern um ausgefüllte Werte ("[Muster AG]" -> "Muster AG") und
    setzt leere bzw. nicht ausgefüllte Vorlagen-Platzhalter auf None."""
    if isinstance(o, dict):
        return {k: (v if k.startswith("_") else bereinige_kundendaten(v)) for k, v in o.items()}
    if isinstance(o, list):
        return [v for v in (bereinige_kundendaten(x) for x in o) if v is not None]
    if isinstance(o, str):
        t = o.strip()
        m = re.fullmatch(r"\[(.*)\]", t)
        if m:
            t = m.group(1).strip()
            if t.lower() in VORLAGE_PLATZHALTER:
                return None
        return t or None
    return o


def auftraggeber(kunde):
    return (kunde.get("kunde") or {}).get("firma")


def standort(kunde):
    return (kunde.get("projekt") or {}).get("standort")


def projekt(kunde):
    return (kunde.get("projekt") or {}).get("bezeichnung")


def am_standort(kunde):
    """'am Standort X' bzw. 'am Standort', falls nicht angegeben."""
    return f"am Standort {standort(kunde)}" if standort(kunde) else "am untersuchten Standort"


def quartalspreise(txt):
    """'Q1=0.216 / Q2=0.197 / ...' -> [21.6, 19.7, ...] Rp/kWh (None, falls anderes Format)."""
    v = re.findall(r"Q\d\s*=\s*([\d.]+)", txt or "")
    return [float(x) * 100 for x in v] if len(v) == 4 else None


def cfg_label(row):
    return f"{num(row.kapazitaet_kwh)} kWh / {num(row.leistung_kw)} kW"


def text_summary(b, k, sz, kunde):
    items = sorted(k.module.items(), key=lambda x: -x[1])
    top = [(n, k.anteile[n]) for n, v in items if v > 0][:3]
    namen = {"SRL": "der Vermarktung von Sekundärregelleistung (SRL)", "Peak-Shaving": "Peak-Shaving",
             "Eigenverbrauchsoptimierung": "der Optimierung des Eigenverbrauchs",
             "Einspeiseoptimierung": "der Einspeiseoptimierung", "Arbitrage": "Arbitrage"}
    jahr = (b.get("zeitraum_von") or "")[-4:]
    teile = []
    if auftraggeber(kunde):
        teile.append(f"Im Auftrag von <b>{auftraggeber(kunde)}</b> wurde")
    else:
        teile.append("Es wurde")
    if projekt(kunde):
        teile.append(f"für das Projekt {projekt(kunde)}")
    if standort(kunde):
        teile.append(am_standort(kunde))
    s = (" ".join(teile) + f" ein Batteriespeicher mit <b>{num(b['bat_kw'])} kW / {num(b['bat_kwh'])} kWh</b> über ein "
         f"ganzes Betriebsjahr ({jahr}) simuliert. "
         f"Der Speicher erwirtschaftet einen operativen Ertrag von <b>{chf(k.op_ertrag)}</b> pro Jahr. "
         f"Nach Abzug von Finanzierung, Unterhalt und Degradation verbleibt ein Nettogewinn von "
         f"<b>{chf(b['netto'])}</b> pro Jahr. Die Investition ist nach rund <b>{num(b['amortisationsdauer_j'], 1)} Jahren</b> "
         f"amortisiert. ")
    if top:
        s += f"Der grösste Teil des Ertrags stammt aus {namen[top[0][0]]} ({top[0][1]:.0f} %)"
        if len(top) > 1:
            rest = [f"{namen[n]} ({a:.0f} %)" for n, a in top[1:]]
            s += ", gefolgt von " + " und ".join(rest)
        s += "."
    if sz is not None:
        s += (f"<br/><br/>Die Grössenanalyse über {len(sz['df'])} Varianten zeigt: Die höchste Kapitalverzinsung "
              f"erzielt ein kleiner Speicher ({cfg_label(sz['best_roi'])}, {sz['best_roi'].roi_pct:.0f} %), den höchsten "
              f"absoluten Gewinn die Variante {cfg_label(sz['best_net'])} ({chf(sz['best_net'].netto_chf)}/Jahr). "
              f"Als ausgewogene Lösung bietet sich <b>{cfg_label(sz['ausgewogen'])}</b> an.")
    return s


def text_ertragsstruktur(b, k, kunde):
    """Neutraler Hinweis zur Ertragsstruktur (Titelseite)."""
    if k.srl <= 0:
        return None
    lo, _ = kunde.get("sensitivitaet_bereich_pct", [-50, 50])
    net_lo, pb_lo, _ = szenario(b, k, lo)
    s = (f"Mit {k.anteile['SRL']:.0f} % trägt die Vermarktung von Sekundärregelleistung den grössten Teil des "
         f"Ertrags. ")
    if net_lo > 0:
        s += (f"Auch bei SRL-Erlösen, die {abs(lo):.0f} % unter dem Niveau von 2025 liegen, bleibt der Speicher mit "
              f"{chf(net_lo)} pro Jahr wirtschaftlich. ")
    s += "Kapitel 05 zeigt die Wirkung unterschiedlicher Preisniveaus auf Gewinn und Amortisation."
    return s


# =============================================================================
# 8) Seiten
# =============================================================================
def seite_titel(d: Doc, b, k, sz, ch):
    c, kunde = d.c, d.kunde
    band = 88 * mm
    c.setFillColor(HexColor(C["dark"])); c.rect(0, H - band, W, band, fill=1, stroke=0)
    # dekorative Kreise
    c.saveState()
    pth = c.beginPath(); pth.rect(0, H - band, W, band); c.clipPath(pth, stroke=0, fill=0)
    for (cx, cy, r, a) in [(W - 25 * mm, H - 20 * mm, 70 * mm, 0.10), (W - 5 * mm, H - 95 * mm, 45 * mm, 0.08),
                           (W - 75 * mm, H - 110 * mm, 30 * mm, 0.06)]:
        c.setFillColor(HexColor(C["mid"])); c.setFillAlpha(a); c.circle(cx, cy, r, fill=1, stroke=0)
    c.restoreState()
    logo = ASSETS / "fleco_logo_rund.png"
    c.setFillColor(white); c.circle(W - M - 15 * mm, H - 27 * mm, 16.5 * mm, fill=1, stroke=0)
    if logo.exists():
        c.drawImage(ImageReader(str(logo)), W - M - 30 * mm, H - 42 * mm, 30 * mm, 30 * mm, mask="auto")
    c.setFillColor(HexColor(C["light"])); c.setFont("F-B", 9)
    c.drawString(M, H - 30 * mm, "WIRTSCHAFTLICHKEITSANALYSE")
    c.setFillColor(white); c.setFont("F-B", 32)
    c.drawString(M, H - 45 * mm, "Batteriespeicher")
    c.setFont("F", 14)
    c.drawString(M, H - 54 * mm, (f"Projekt {projekt(kunde)}  ·  " if projekt(kunde) else "")
                 + f"{num(b['bat_kw'])} kW / {num(b['bat_kwh'])} kWh")
    # Meta-Zeile: nur ausgefüllte Angaben, gleichmässig verteilt
    meta = [(lab, val) for lab, val in (("Auftraggeber", auftraggeber(kunde)), ("Projekt", projekt(kunde)),
                                        ("Standort", standort(kunde)), ("Datum", kunde.get("bericht_datum")))
            if val]
    breiten = {"Auftraggeber": 50, "Projekt": 36, "Standort": 52, "Datum": 30}
    x = M
    for lab, val in meta:
        c.setFont("F", 7.5); c.setFillColor(HexColor(C["light"])); c.drawString(x, H - 68 * mm, lab.upper())
        c.setFont("F-B", 10); c.setFillColor(white); c.drawString(x, H - 74 * mm, val)
        x += max(breiten[lab] * mm, pdfmetrics.stringWidth(val, "F-B", 10) + 8 * mm)

    y = H - band - 10 * mm
    y = d.h2("Auf einen Blick", y)
    gap = 4 * mm
    tw = (CW - 3 * gap) / 4
    tiles = [(num(b["netto"]), "CHF/J.", "Nettogewinn", C["dark"]),
             (num(b["amortisationsdauer_j"], 1), "Jahre", "Amortisation", C["dark"]),
             (num(b["kapitalverzinsung_pct"], 1), "%", "Kapitalverzinsung", C["dark"]),
             (num(k.op_ertrag), "CHF/J.", "Operativer Ertrag", C["green"])]
    for i, (v, u, l, f) in enumerate(tiles):
        d.kpi(M + i * (tw + gap), y, tw, 22 * mm, v, u, l, f)
    y -= 29 * mm
    y -= para(c, text_summary(b, k, sz, kunde), M, y, CW, S(9.6)) + 5 * mm
    y = d.h2("Ertrag nach Anwendung", y)
    y -= img(c, ch["donut"], M, y, CW) + 4 * mm
    r = text_ertragsstruktur(b, k, kunde)
    if r and y > 44 * mm:
        d.callout(M, y, CW, r, titel="Ertragsstruktur", farbe=C["mid"], fill=C["pale"])


def seite_ausgangslage(d: Doc, b, k, ch):
    c, kunde = d.c, d.kunde
    y = d.header(1, "Ausgangslage & System")
    # Auftrag & Objekt
    colw = (CW - 8 * mm) / 2
    adr = (kunde.get("kunde") or {}).get("adresse") or []
    kp = (kunde.get("kunde") or {}).get("kontaktperson")
    links = [("__h__", "Auftraggeber")]
    if auftraggeber(kunde):
        links.append(("Firma", auftraggeber(kunde)))
    if kp:
        links.append(("Ansprechperson", kp))
    if adr:
        links.append(("Adresse", ", ".join(adr)))
    rechts = [("__h__", "Objekt")]
    rechts += [(lab, v) for lab, v in (("Projekt", projekt(kunde)), ("Standort", standort(kunde))) if v]
    rechts += [(k_, v_) for k_, v_ in ((kunde.get("projekt") or {}).get("zusatz") or {}).items() if v_]
    if len(links) > 1 or len(rechts) > 1:   # Block nur zeigen, wenn es etwas anzuzeigen gibt
        h1 = d.tabelle(M, y, colw, links) if len(links) > 1 else 0
        h2 = d.tabelle(M + colw + 8 * mm, y, colw, rechts) if len(rechts) > 1 else 0
        y -= max(h1, h2) + 6 * mm

    txt = (f"Untersucht wird, wie sich ein Batteriespeicher {am_standort(kunde)} auf Energiekosten, "
           f"Lastspitzen und zusätzliche Markterlöse auswirkt. Grundlage ist der Lastgang mit einem Jahresverbrauch von "
           f"{num(b['last_mwh'], 1)} MWh sowie die bestehende PV-Anlage mit {num(b['pv_kwp'])} kWp. Verglichen wird "
           f"der Betrieb <b>ohne</b> und <b>mit</b> Batteriespeicher über den Zeitraum {b['zeitraum_von']} bis "
           f"{b['zeitraum_bis']}.")
    y -= para(c, txt, M, y, CW, S(9.4)) + 3 * mm
    y -= img(c, ch["schema"], M + 8 * mm, y, CW - 16 * mm) + 5 * mm

    ht, nt = quartalspreise(b.get("ht_preis")), quartalspreise(b.get("nt_preis"))
    left = [("__h__", "Standort & Verbrauch"),
            ("Jahresverbrauch", f"{num(b['last_mwh'], 1)} MWh"),
            ("PV-Anlage (DC)", f"{num(b['pv_kwp'], 1)} kWp"),
            ("PV-Abregelung bei negativen Preisen", "ja" if (b.get("pv_abschaltung_neg") or "").lower() == "ja" else "nein"),
            ("Referenzjahr", (b.get("zeitraum_von") or "")[-4:]),
            ("__h__", "Tarife"),
            ("Tarifmodell", "Hoch-/Niedertarif" if b.get("tarifschema") == "HT_NT" else (b.get("tarifschema") or "–")),
            ("Hochtarif-Zeiten", f"{b.get('ht_tage', '')}, {b.get('ht_fenster', '')}")]
    if ht and nt:
        left += [("__qh__", ("Energiepreis in Rp/kWh", ["Q1", "Q2", "Q3", "Q4"])),
                 ("__q__", ("Hochtarif", [num(v, 1) for v in ht])),
                 ("__q__", ("Niedertarif", [num(v, 1) for v in nt]))]
    else:
        left += [("Energiepreis HT", b.get("ht_preis") or "–"), ("Energiepreis NT", b.get("nt_preis") or "–")]
    left += [("Leistungspreis", f"{num(b['leistungspreis'], 2)} CHF/kW/Monat"),
             ("Rücklieferung PV / Batterie", f"{b.get('pv_schema', '')} / {b.get('bat_schema', '')}")]
    right = [("__h__", "Batteriespeicher"),
             ("Leistung", f"{num(b['bat_kw'])} kW"),
             ("Kapazität", f"{num(b['bat_kwh'])} kWh"),
             ("C-Rate", f"{b['bat_kw'] / b['bat_kwh']:.2g}C"),
             ("Wirkungsgrad (Zyklus)", f"{num(b['rte_pct'], 1)} %"),
             ("Nutzbarer Ladezustand", f"{num(b['soc_min'])} – {num(b['soc_max'])} %"),
             ("__h__", "Investition & Finanzierung"),
             ("Investition" + (" (Annahme)" if k.investition_abgeleitet else ""), f"ca. {chf(round(k.investition, -2))}"),
             ("Kapitalzins", f"{num(b['kapitalzins_pct'], 1)} % p.a."),
             ("Nutzungsdauer", f"{num(k.lebensdauer_j)} Jahre" if k.lebensdauer_j else "–")]
    d.tabelle(M, y, colw, left)
    d.tabelle(M + colw + 8 * mm, y, colw, right)


def seite_methodik(d: Doc, b, k):
    c = d.c
    y = d.header(2, "Methodik & Vermarktungsmöglichkeiten")
    schritte = [
        ("1", "Datengrundlage", f"Lastgang und PV-Produktion des Standorts, Hoch-/Niedertarif und Leistungspreis, "
                                f"Spotpreise für die Rücklieferung sowie SRL-Marktpreise {(b.get('zeitraum_von') or '')[-4:]}."),
        ("2", "Jahressimulation", "Der Batteriebetrieb wird über ein ganzes Jahr simuliert und dem Betrieb ohne "
                                  "Batterie gegenübergestellt. Die Differenz ergibt den Nutzen des Speichers."),
        ("3", "Bewertung", "Erträge je Anwendung, Kosten für Investition, Finanzierung, Unterhalt und Alterung "
                           "sowie Kennzahlen zu Rendite, Amortisation und Risiko."),
    ]
    bw = (CW - 2 * 5 * mm) / 3
    bh = 39 * mm
    for i, (n, t, txt) in enumerate(schritte):
        x = M + i * (bw + 5 * mm)
        rrect(c, x, y - bh, bw, bh, 3 * mm, fill=C["bg"])
        c.setFillColor(HexColor(C["dark"])); c.circle(x + 7 * mm, y - 7 * mm, 3.6 * mm, fill=1, stroke=0)
        c.setFont("F-B", 10); c.setFillColor(white); c.drawCentredString(x + 7 * mm, y - 8.3 * mm, n)
        c.setFont("F-B", 10); c.setFillColor(HexColor(C["ink"])); c.drawString(x + 13 * mm, y - 8.3 * mm, t)
        para(c, txt, x + 4 * mm, y - 13 * mm, bw - 8 * mm, S(8.3, C["grey"]))
        if i < 2:
            c.setFillColor(HexColor(C["mid"])); c.setFont("F-B", 14)
            c.drawCentredString(x + bw + 2.5 * mm, y - bh / 2 - 2 * mm, "›")
    y -= bh + 6 * mm
    txt = ("<b>Betriebsstrategie.</b> Die Batterie bedient zuerst die lokalen Anwendungen am Standort "
           "(Eigenverbrauch, Lastspitzen). Die dafür nicht benötigte Leistung (Residualflexibilität) wird als "
           f"Sekundärregelleistung vermarktet. Dabei wird der Ladezustand in einem Band von {num(b['soc_min'])} – {num(b['soc_max'])} % gehalten, "
           "damit die Batterie Regelabrufe in beide Richtungen jederzeit erbringen kann und schonend betrieben wird.")
    y -= para(c, txt, M, y, CW, S(9.4)) + 6 * mm
    y = d.h2("Untersuchte Anwendungen", y)
    besch = {
        "SRL": "Die Batterie stellt Swissgrid Regelleistung zur Stabilisierung des Stromnetzes bereit. Vergütet werden "
               "die vorgehaltene Leistung und die abgerufene Energie.",
        "Peak-Shaving": f"Die Batterie entlädt bei hoher Last und senkt so die monatliche Bezugsspitze, die über den "
                        f"Leistungspreis von {num(b['leistungspreis'], 2)} CHF/kW/Monat verrechnet wird.",
        "Eigenverbrauchsoptimierung": "Überschüssiger Solarstrom wird gespeichert und später vor Ort genutzt, statt ihn "
                                      "einzuspeisen und teureren Netzstrom zu beziehen.",
        "Einspeiseoptimierung": "Solarstrom wird bei tiefen Spotpreisen eingespeichert und zu Zeiten höherer Preise "
                                "ins Netz zurückgeliefert.",
        "Arbitrage": "Netzstrom wird zu günstigen Zeiten bezogen und zu teureren Zeiten zurückgespeist.",
    }
    order = sorted(k.module.keys(), key=lambda n: -k.module[n])
    cw2 = (CW - 5 * mm) / 2
    ch_ = 31 * mm
    for i, n in enumerate(order):
        col, row = i % 2, i // 2
        x = M + col * (cw2 + 5 * mm)
        yt = y - row * (ch_ + 4 * mm)
        rrect(c, x, yt - ch_, cw2, ch_, 3 * mm, fill="#FFFFFF", stroke=C["line"])
        rrect(c, x, yt - ch_, 1.4 * mm, ch_, 0.7 * mm, fill=MODUL_FARBEN[n])
        c.setFont("F-B", 9.8); c.setFillColor(HexColor(C["ink"])); c.drawString(x + 5 * mm, yt - 6.5 * mm, MODUL_LANG[n])
        c.setFont("F-B", 9.8); c.setFillColor(HexColor(C["dark"]))
        c.drawRightString(x + cw2 - 4 * mm, yt - 6.5 * mm, chf(k.module[n]))
        para(c, besch[n], x + 5 * mm, yt - 9.5 * mm, cw2 - 9 * mm, S(8.1, C["grey"]))
        # Anteilsbalken
        bx, bw_ = x + 5 * mm, cw2 - 9 * mm
        rrect(c, bx, yt - ch_ + 3.5 * mm, bw_, 1.6 * mm, 0.8 * mm, fill=C["line"])
        rrect(c, bx, yt - ch_ + 3.5 * mm, max(bw_ * k.anteile[n] / 100, 1.6 * mm), 1.6 * mm, 0.8 * mm,
              fill=MODUL_FARBEN[n])
        c.setFont("F", 7); c.setFillColor(HexColor(C["grey"]))
        c.drawRightString(x + cw2 - 4 * mm, yt - ch_ + 6.3 * mm, f"{pct(k.anteile[n])} des Ertrags")
    if len(order) % 2 == 1:
        i = len(order)
        x = M + (cw2 + 5 * mm)
        yt = y - (i // 2) * (ch_ + 4 * mm)
        rrect(c, x, yt - ch_, cw2, ch_, 3 * mm, fill=C["bg"])
        c.setFont("F-B", 9.8); c.setFillColor(HexColor(C["grey"])); c.drawString(x + 5 * mm, yt - 6.5 * mm,
                                                                                 "Nicht Teil dieser Analyse")
        para(c, d.kunde.get("texte", {}).get("nicht_bewertet") or
             "Weitere Regelenergieprodukte (Primär- und Tertiärregelleistung) sowie dynamische Netztarife wurden "
             "nicht bewertet und können zusätzliches Potenzial bieten.",
             x + 5 * mm, yt - 9.5 * mm, cw2 - 9 * mm, S(8.1, C["grey"]))


def seite_wirkung(d: Doc, b, ch, kunde):
    c = d.c
    y = d.header(3, "Wirkung am Standort")
    txt = ("Die Batterie verändert das Bezugsverhalten des Standorts deutlich: Mehr Solarstrom wird vor Ort genutzt, "
           "weniger Energie wird aus dem Netz bezogen oder eingespeist und die monatlichen Lastspitzen sinken.")
    y -= para(c, txt, M, y, CW, S(9.6)) + 5 * mm
    gap = 4 * mm
    tw = (CW - 3 * gap) / 4
    vals = [("Eigenverbrauchsquote", f"{num(b['evq_ohne'])}", f"{num(b['evq_mit'])}", "%"),
            ("Ø Netzbezug", num(b["netzbezug_ohne"], 1), num(b["netzbezug_mit"], 1), "kW"),
            ("Ø Netzeinspeisung", num(b["einspeisung_ohne"], 1), num(b["einspeisung_mit"], 1), "kW"),
            ("Ø Monatsspitze", num(b["spitze_avg_ohne"]), num(b["spitze_avg_mit"]), "kW")]
    for i, (l, o, m, u) in enumerate(vals):
        d.vergleich(M + i * (tw + gap), y, tw, 24 * mm, l, o, m, u, True)
    y -= 31 * mm
    y = d.h2("Monatliche Lastspitzen", y)
    if ch.get("spitzen"):
        y -= img(c, ch["spitzen"], M, y, CW) + 3 * mm
    red = (1 - b["spitze_avg_mit"] / b["spitze_avg_ohne"]) * 100
    txt = (f"Die höchste Bezugsleistung pro Monat sinkt im Mittel von {num(b['spitze_avg_ohne'])} kW auf "
           f"{num(b['spitze_avg_mit'])} kW (−{red:.0f} %), die Jahreshöchstspitze von {num(b['spitze_max_ohne'])} kW auf "
           f"{num(b['spitze_max_mit'])} kW. Bei einem Leistungspreis von {num(b['leistungspreis'], 2)} CHF/kW/Monat "
           f"entspricht dies einer Einsparung von <b>{chf(b['einnahmen']['Peak-Shaving'])}</b> pro Jahr.")
    y -= para(c, txt, M, y, CW, S(9.4)) + 6 * mm
    evd = b["evq_mit"] - b["evq_ohne"]
    txt = (f"<b>Eigenverbrauch.</b> Der Anteil des Solarstroms, der direkt am Standort genutzt wird, steigt um "
           f"{num(evd)} Prozentpunkte auf {num(b['evq_mit'])} %. Die eingesparten Netzbezugskosten betragen "
           f"<b>{chf(b['einnahmen']['Eigenverbrauchsoptimierung'])}</b> pro Jahr."
           + (" Der Effekt bleibt begrenzt, weil der Standort bereits ohne Batterie einen grossen Teil des "
              "Solarstroms selbst verbraucht." if b["evq_ohne"] >= 60 else ""))
    y -= para(c, txt, M, y, CW, S(9.4)) + 4 * mm
    if b.get("wert_pv_abschaltung"):
        para(c, f"Hinweis: Die Abregelung der PV-Anlage bei negativen Strompreisen hat bereits ohne Batterie einen Wert "
                f"von {chf(b['wert_pv_abschaltung'])} pro Jahr. Dieser Wert ist im Batterienutzen nicht enthalten.",
             M, y, CW, S(8, C["grey"]))


def seite_wirtschaft(d: Doc, b, k, ch):
    c = d.c
    y = d.header(4, "Wirtschaftlichkeit")
    txt = ("Die Grafik zeigt, wie sich der jährliche Nettogewinn zusammensetzt: links die Erträge je Anwendung, "
           "rechts die jährlichen Kosten für Unterhalt, Batteriealterung und Finanzierung der Investition.")
    y -= para(c, txt, M, y, CW, S(9.6)) + 3 * mm
    y -= img(c, ch["waterfall"], M, y, CW) + 5 * mm

    colw = (CW - 8 * mm) / 2
    items, klein = module_sortiert(k, schwelle_pct=0)
    rows = [("__h__", "Erträge pro Jahr")]
    rows += [(MODUL_LANG.get(n, n), chf(v)) for n, v in items]
    rows += [("**Operativer Ertrag", chf(k.op_ertrag))]
    rows2 = [("__h__", "Kosten pro Jahr"),
             ("Unterhalt", chf(-b["unterhalt"])), ("Batterie-Degradation", chf(-b["degradation"])),
             (f"Kapitalkosten ({num(b['kapitalzins_pct'], 0)} %)", chf(-b["kapitalkosten"])),
             ("Amortisation", chf(-b["amortisation"])),
             ("**Total Kosten", chf(-b["total_ausgaben"])),
             ("**Nettogewinn", chf(b["netto"]))]
    h1 = d.tabelle(M, y, colw, rows)
    h2 = d.tabelle(M + colw + 8 * mm, y, colw, rows2)
    y -= max(h1, h2) + 7 * mm

    y = d.h2("Rendite", y)
    gap = 4 * mm
    tw = (CW - 2 * gap) / 3
    d.kpi(M, y, tw, 24 * mm, num(b["amortisationsdauer_j"], 1), "Jahre", "Amortisationsdauer",
          hint="bei Investition aus Eigenmitteln")
    d.kpi(M + tw + gap, y, tw, 24 * mm, num(b["kapitalverzinsung_pct"], 1), "%", "Kapitalverzinsung",
          hint="Jahresertrag nach Unterhalt / Investition")
    d.kpi(M + 2 * (tw + gap), y, tw, 24 * mm, num(b["lcos_rp"], 1), "Rp/kWh", "Speicherkosten (LCOS)",
          farbe=C["green"], hint="Kosten je ausgespeicherter kWh")
    y -= 30 * mm
    txt = ("<b>Lesehilfe.</b> Der <i>Nettogewinn</i> berücksichtigt die Finanzierung der Investition über "
           f"{num(k.lebensdauer_j) if k.lebensdauer_j else 'die Nutzungsdauer'} Jahre zu {num(b['kapitalzins_pct'], 0)} % "
           "(Annuität). <i>Amortisationsdauer</i> und <i>Kapitalverzinsung</i> betrachten dagegen eine Investition aus "
           "Eigenmitteln. Die Amortisationsdauer gibt an, nach wie vielen Jahren der jährliche Ertrag nach Unterhalt "
           "und Alterung die Investition gedeckt hat. Die Kapitalverzinsung ist dieser Jahresertrag in Prozent der "
           "Investition.")
    if b.get("lcos_rp") and b["lcos_rp"] > 25:
        txt += (" Die <i>Speicherkosten</i> liegen über dem Strompreis. Das ist kein Widerspruch, denn der Grossteil des "
                "Nutzens entsteht nicht durch zeitlich verschobene Energie, sondern durch vorgehaltene Leistung (SRL) "
                "und gesenkte Lastspitzen.")
    para(c, txt, M, y, CW, S(8.4, C["grey"]))


def seite_risiko(d: Doc, b, k, ch):
    c, kunde = d.c, d.kunde
    lo, hi = kunde.get("sensitivitaet_bereich_pct", [-50, 50])
    y = d.header(5, "Sensitivität & Risiken")
    txt = ("Die Erträge aus Sekundärregelleistung richten sich nach der Entwicklung des Regelleistungsmarkts. "
           "Um die Robustheit der Investition einzuordnen, zeigt die folgende Auswertung Nettogewinn und "
           f"Amortisationsdauer für ein SRL-Preisniveau zwischen {sgn(lo)} % und {sgn(hi)} % gegenüber 2025.")
    y -= para(c, txt, M, y, CW, S(9.6)) + 4 * mm
    y -= img(c, ch["sens"], M, y, CW) + 6 * mm

    szen = kunde.get("sensitivitaet_szenarien_pct", [-50, -25, 0, 25, 50])
    gap = 3 * mm
    tw = (CW - (len(szen) - 1) * gap) / len(szen)
    h = 27 * mm
    for i, dlt in enumerate(szen):
        net, pb, roi = szenario(b, k, dlt)
        x = M + i * (tw + gap)
        basis = dlt == 0
        rrect(c, x, y - h, tw, h, 3 * mm, fill=C["dark"] if basis else C["pale"])
        hc = HexColor(C["light"]) if basis else HexColor(C["grey"])
        vc = white if basis else HexColor(C["dark"])
        c.setFont("F", 7.5); c.setFillColor(hc)
        c.drawString(x + 3.5 * mm, y - 6 * mm, "BASIS 2025" if basis else f"SRL-ERLÖSE {sgn(dlt)} %")
        c.setFont("F-B", 14); c.setFillColor(vc); c.drawString(x + 3.5 * mm, y - 13.5 * mm, num(net))
        c.setFont("F", 7.5); c.setFillColor(hc)
        c.drawString(x + 3.5 * mm, y - 17.5 * mm, "CHF Nettogewinn / Jahr")
        c.drawString(x + 3.5 * mm, y - 21.5 * mm, f"Amortisation {num(pb, 1)} J.")
        c.drawString(x + 3.5 * mm, y - 25 * mm, f"Verzinsung {num(roi, 1)} %")
    y -= h + 7 * mm

    net_lo, pb_lo, _ = szenario(b, k, lo)
    net_hi, pb_hi, _ = szenario(b, k, hi)
    if net_lo > 0:
        txt = (f"Der Speicher ist über die gesamte betrachtete Bandbreite wirtschaftlich: Auch bei einem Rückgang der "
               f"SRL-Erlöse um {abs(lo):.0f} % erzielt er einen Nettogewinn von <b>{chf(net_lo)}</b> pro Jahr und "
               f"amortisiert sich in {num(pb_lo, 1)} Jahren. Bei höheren Erlösen steigt der Nettogewinn auf bis zu "
               f"<b>{chf(net_hi)}</b> pro Jahr, die Amortisationsdauer sinkt auf {num(pb_hi, 1)} Jahre.")
    else:
        txt = (f"Innerhalb der betrachteten Bandbreite liegt der Nettogewinn zwischen {chf(net_lo)} und "
               f"{chf(net_hi)} pro Jahr, die Amortisationsdauer zwischen {num(pb_hi, 1)} und {num(pb_lo, 1)} Jahren.")
    y -= para(c, txt, M, y, CW, S(9.4)) + 6 * mm
    txt = ("<b>Weitere Einflussfaktoren.</b> Die Berechnung basiert auf einem historischen Jahr (Lastgang, Preise). "
           "Abweichungen ergeben sich insbesondere durch Veränderungen im Verbrauchsverhalten, bei Netz- und "
           "Energietarifen, durch Ausfallzeiten der Anlage sowie durch die tatsächliche Alterung der Batterie.")
    para(c, txt, M, y, CW, S(8.6, C["grey"]))


def seite_dimensionierung(d: Doc, b, sz, ch):
    c = d.c
    y = d.header(6, "Dimensionierung")
    txt = (f"Um die passende Batteriegrösse zu bestimmen, wurden {len(sz['df'])} Varianten mit "
           f"{len(sz['caps'])} Kapazitäten ({num(min(sz['caps']))} – {num(max(sz['caps']))} kWh) und "
           f"{len(sz['rates'])} Leistungsverhältnissen (C-Rate) gerechnet. Die C-Rate gibt an, wie gross die Leistung im "
           f"Verhältnis zur Kapazität ist: 1C bedeutet z. B. 300 kW bei 300 kWh.")
    y -= para(c, txt, M, y, CW, S(9.6)) + 2 * mm
    y -= img(c, ch["heat"], M, y, CW) + 2 * mm
    lg = (f"<b>R</b> = höchste Kapitalverzinsung ({cfg_label(sz['best_roi'])})   ·   "
          f"<b>A</b> = ausgewogen, höchster Gewinn bei mind. {sz['ziel_roi']:g} % Verzinsung "
          f"({cfg_label(sz['ausgewogen'])})   ·   <b>E</b> = höchster Nettogewinn ({cfg_label(sz['best_net'])})")
    y -= para(c, lg, M, y, CW, S(8, C["grey"])) + 5 * mm
    y = d.h2("Rendite und Gewinn im Vergleich", y)
    y -= img(c, ch["trade"], M, y, CW)


def kernaussagen(d: Doc, b, k, kunde, y):
    """Kernaussagen-Kacheln für Berichte ohne Sizing-Analyse."""
    c = d.c
    y = d.h2("Kernaussagen", y)
    lo, _ = kunde.get("sensitivitaet_bereich_pct", [-50, 50])
    net_lo, pb_lo, _ = szenario(b, k, lo)
    top = sorted(k.module.items(), key=lambda x: -x[1])[:2]
    items = [
        ("Wirtschaftlichkeit",
         f"Nettogewinn von {chf(b['netto'])} pro Jahr. Die Investition amortisiert sich in "
         f"{num(b['amortisationsdauer_j'], 1)} Jahren, die Kapitalverzinsung beträgt {num(b['kapitalverzinsung_pct'], 1)} %."),
        ("Ertragsquellen",
         f"Der operative Ertrag von {chf(k.op_ertrag)} pro Jahr stammt vor allem aus "
         f"{MODUL_LANG[top[0][0]]} ({k.anteile[top[0][0]]:.0f} %)"
         + (f" und {MODUL_LANG[top[1][0]]} ({k.anteile[top[1][0]]:.0f} %)." if len(top) > 1 else ".")),
    ]
    if b.get("spitze_avg_ohne") is not None:
        items.append(("Wirkung am Standort",
                      f"Die mittlere Monatsspitze sinkt von {num(b['spitze_avg_ohne'])} kW auf "
                      f"{num(b['spitze_avg_mit'])} kW. Die Eigenverbrauchsquote steigt von {num(b['evq_ohne'])} % auf "
                      f"{num(b['evq_mit'])} %."))
    if k.srl > 0:
        items.append(("Robustheit",
                      f"Auch bei SRL-Erlösen, die {abs(lo):.0f} % unter dem Niveau von 2025 liegen, beträgt der "
                      f"Nettogewinn {chf(net_lo)} pro Jahr bei einer Amortisationsdauer von {num(pb_lo, 1)} Jahren."
                      if net_lo > 0 else
                      f"Bei SRL-Erlösen, die {abs(lo):.0f} % unter dem Niveau von 2025 liegen, beträgt der "
                      f"Nettogewinn {chf(net_lo)} pro Jahr."))
    gap = 4 * mm
    cw2 = (CW - gap) / 2
    h = 30 * mm
    for i, (t, txt) in enumerate(items):
        x = M + (i % 2) * (cw2 + gap)
        yt = y - (i // 2) * (h + gap)
        rrect(c, x, yt - h, cw2, h, 3 * mm, fill=C["pale"])
        rrect(c, x, yt - h, 1.4 * mm, h, 0.7 * mm, fill=C["dark"])
        c.setFont("F-B", 10); c.setFillColor(HexColor(C["dark"])); c.drawString(x + 5 * mm, yt - 6.5 * mm, t)
        para(c, txt, x + 5 * mm, yt - 9.5 * mm, cw2 - 9 * mm, S(8.8))
    rows = (len(items) + 1) // 2
    return y - rows * (h + gap) - 4 * mm


def seite_empfehlung(d: Doc, b, k, sz, kunde):
    c = d.c
    nr = 7 if sz is not None else 6
    y = d.header(nr, "Empfehlung & Fazit" if sz is not None else "Fazit")
    df = sz["df"] if sz is not None else None
    if sz is not None:
        punkte = []
        if sz["dominante_crate"] is not None:
            punkte.append(f"<b>Leistung ist der stärkste Hebel.</b> Bei jeder Kapazität erzielt die Variante mit "
                          f"{sz['dominante_crate']:g}C den höchsten Gewinn. Die SRL-Erlöse skalieren vor allem mit der "
                          f"vorgehaltenen Leistung (kW). Zusätzliche Kapazität ohne zusätzliche Leistung bringt deshalb nur "
                          f"wenig Mehrertrag.")
        roi_small = df[df.kapazitaet_kwh == min(sz["caps"])].roi_pct.max()
        roi_big = df[df.kapazitaet_kwh == max(sz["caps"])].roi_pct.max()
        punkte.append(f"<b>Grösser heisst mehr Gewinn, aber tiefere Verzinsung.</b> Die Kapitalverzinsung sinkt mit "
                      f"zunehmender Grösse von {roi_small:.0f} % auf {roi_big:.0f} %. Der Grund: Schon ein kleiner Speicher "
                      f"deckt einen grossen Teil des Nutzens aus Eigenverbrauch und Peak-Shaving ab. Jede zusätzliche "
                      f"Kilowattstunde wird deshalb vor allem für die SRL-Vermarktung eingesetzt. Damit steigt mit der "
                      f"Grösse auch der Anteil des Ertrags, der von den Regelleistungspreisen abhängt.")
        if len(sz["negative"]):
            neg = sz["negative"]
            punkte.append(f"<b>Zu wenig Leistung lohnt sich nicht.</b> "
                          + (f"{len(neg)} Varianten mit tiefer C-Rate (z. B. {cfg_label(neg.iloc[0])}) erzielen "
                             if len(neg) > 1 else f"Die Variante {cfg_label(neg.iloc[0])} erzielt ")
                          + "keinen Nettogewinn.")
        if sz["rand_optimum"]:
            punkte.append("<b>Obergrenze.</b> Der Gewinn steigt bis zur grössten untersuchten Variante weiter an. "
                          "Die sinnvolle Obergrenze ergibt sich deshalb aus Investitionsbudget, Platz, "
                          "Netzanschluss und Risikobereitschaft.")
        for p in punkte:
            c.setFillColor(HexColor(C["mid"])); c.circle(M + 1.5 * mm, y - 2.3 * mm, 1.1 * mm, fill=1, stroke=0)
            y -= para(c, p, M + 5 * mm, y, CW - 5 * mm, S(9)) + 2 * mm
        y -= 2 * mm
        y = d.h2("Drei Varianten zur Auswahl", y)
        gap = 4 * mm
        tw = (CW - 2 * gap) / 3
        h = 49 * mm
        karten = [("R", "Rendite-optimiert", sz["best_roi"], "Kleine Investition mit der höchsten Verzinsung und dem geringsten absoluten Marktrisiko."),
                  ("A", "Ausgewogen", sz["ausgewogen"], f"Hoher Gewinn bei einer Verzinsung von mindestens {sz['ziel_roi']:g} %."),
                  ("E", "Ertrags-maximiert", sz["best_net"], "Höchster absoluter Gewinn bei der grössten Abhängigkeit von den SRL-Preisen.")]
        for i, (tag, titel, row, besch) in enumerate(karten):
            x = M + i * (tw + gap)
            hero = tag == "A"
            rrect(c, x, y - h, tw, h, 3 * mm, fill=C["dark"] if hero else C["bg"])
            fc = white if hero else HexColor(C["ink"])
            gc = HexColor(C["light"]) if hero else HexColor(C["grey"])
            c.setFillColor(HexColor(C["light"]) if hero else HexColor(C["dark"]))
            c.circle(x + 7 * mm, y - 7 * mm, 3.3 * mm, fill=1, stroke=0)
            c.setFont("F-B", 9); c.setFillColor(HexColor(C["dark"]) if hero else white)
            c.drawCentredString(x + 7 * mm, y - 8.2 * mm, tag)
            c.setFont("F-B", 10); c.setFillColor(fc); c.drawString(x + 12.5 * mm, y - 8.4 * mm, titel)
            c.setFont("F-B", 13); c.drawString(x + 4 * mm, y - 17 * mm, cfg_label(row))
            c.setFont("F", 8.5); c.setFillColor(gc)
            c.drawString(x + 4 * mm, y - 23.5 * mm, f"Nettogewinn")
            c.drawRightString(x + tw - 4 * mm, y - 23.5 * mm, f"{chf(row.netto_chf)}/J.")
            c.drawString(x + 4 * mm, y - 28.5 * mm, f"Kapitalverzinsung")
            c.drawRightString(x + tw - 4 * mm, y - 28.5 * mm, f"{row.roi_pct:.0f} %")
            c.drawString(x + 4 * mm, y - 33.5 * mm, f"Operativer Ertrag")
            c.drawRightString(x + tw - 4 * mm, y - 33.5 * mm, f"{chf(row.operativ_chf)}/J.")
            para(c, besch, x + 4 * mm, y - 36.5 * mm, tw - 8 * mm, S(7.6, C["light"] if hero else C["grey"]))
        y -= h + 5 * mm

    jahr = (b.get('zeitraum_von') or '')[-4:]
    if sz is None:
        y = kernaussagen(d, b, k, kunde, y)
    y = d.h2("Fazit", y)
    if sz is not None:
        fazit_auto = (
            f"Ein Batteriespeicher ist {am_standort(kunde)} unter den Marktbedingungen {jahr} wirtschaftlich "
            f"attraktiv. Die detailliert analysierte Variante ({num(b['bat_kw'])} kW / {num(b['bat_kwh'])} kWh) erzielt "
            f"einen Nettogewinn von {chf(b['netto'])} pro Jahr und amortisiert sich in "
            f"{num(b['amortisationsdauer_j'], 1)} Jahren. Entscheidend für die Wahl der Grösse ist die "
            f"Risikobereitschaft. Je grösser der Speicher, desto höher ist der mögliche Gewinn, aber auch die "
            f"Abhängigkeit von den Preisen für Regelleistung.")
    else:
        bewertung = "wirtschaftlich attraktiv" if b["netto"] > 0 else "unter den getroffenen Annahmen nicht wirtschaftlich"
        fazit_auto = (
            f"Ein Batteriespeicher mit {num(b['bat_kw'])} kW / {num(b['bat_kwh'])} kWh ist {am_standort(kunde)} "
            f"unter den Marktbedingungen {jahr} {bewertung}. Er erzielt einen Nettogewinn von {chf(b['netto'])} pro "
            f"Jahr und amortisiert sich in {num(b['amortisationsdauer_j'], 1)} Jahren. Neben dem finanziellen Ertrag "
            f"senkt der Speicher die Lastspitzen des Standorts und erhöht den Anteil des selbst genutzten Solarstroms.")
    fazit = kunde.get("texte", {}).get("fazit") or fazit_auto
    y -= para(c, fazit, M, y, CW, S(9.2)) + 4 * mm
    hinweis = (f"Alle Ergebnisse beruhen auf historischen Daten des Jahres {(b.get('zeitraum_von') or '')[-4:]} "
               "(Lastgang, Strom- und Regelleistungspreise). Sie zeigen, was der Speicher unter diesen Bedingungen "
               "erwirtschaftet hätte. Zukünftige Erträge können davon abweichen. Die Ergebnisse sind deshalb keine "
               "Garantie für die Wirtschaftlichkeit in den kommenden Jahren.")
    y -= d.callout(M, y, CW, hinweis, titel="Wichtiger Hinweis: Berechnung auf Basis historischer Daten",
                   farbe=C["amber"], fill=C["amber_pale"]) + 4 * mm
    return y


KONTAKT_H = 34 * mm
DISCLAIMER = ("<b>Haftungsausschluss.</b> Alle Berechnungen sind unverbindlich. Die Fleco Power AG übernimmt keine "
              "Verantwortung für deren Genauigkeit oder Vollständigkeit. Jegliche Haftung aus der Verwendung oder "
              "Nichtverwendung der in diesem Bericht präsentierten Informationen ist ausgeschlossen. Alle angenommenen "
              "Hypothesen werden im Bericht transparent dargestellt. Die Berechnungen beruhen auf Daten aus der Vergangenheit "
              "und stellen keine Garantie für zukünftige Erträge dar.")


def abschluss_block(d: Doc, kunde):
    """Kontaktbox + Haftungsausschluss, unten auf der Seite verankert."""
    c = d.c
    y0 = 20 * mm
    rrect(c, M, y0, CW, KONTAKT_H, 3 * mm, fill=C["pale"])
    il = ASSETS / "fleco_illustration.png"
    if il.exists():
        iw = 68 * mm
        img(c, il, M + CW - iw - 3 * mm, y0 + KONTAKT_H - 2.5 * mm, iw)
    c.setFont("F-B", 11); c.setFillColor(HexColor(C["dark"]))
    c.drawString(M + 6 * mm, y0 + KONTAKT_H - 8 * mm, "Bei Fragen stehen wir Ihnen jederzeit gerne zur Verfügung.")
    k = kunde.get("kontakt", {})
    l1 = [f"<b>{k.get('name', '')}</b>", k.get("funktion") or "", "Fleco Power AG"]
    l2 = [k.get("telefon", "+41 52 209 04 04"), k.get("email", "info@flecopower.ch"), "www.flecopower.ch"]
    para(c, "<br/>".join(z for z in l1 if z), M + 6 * mm, y0 + KONTAKT_H - 12 * mm, 45 * mm, S(8.8))
    para(c, "<br/>".join(l2), M + 52 * mm, y0 + KONTAKT_H - 12 * mm, 50 * mm, S(8.8))
    para(c, DISCLAIMER, M, y0 + KONTAKT_H + 3 * mm + 14 * mm, CW, S(7.2, C["grey"]))


ABSCHLUSS_RESERVE = 20 * mm + KONTAKT_H + 19 * mm


def seite_hinweise(d: Doc, kunde):
    d.header(d.kapitel + 1, "Hinweise & Kontakt")
    abschluss_block(d, kunde)


# =============================================================================
# 9) Main
# =============================================================================
def build(bericht, sweep, kunde_json, out):
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    import tempfile, shutil
    tmp = Path(tempfile.mkdtemp(prefix="fleco_charts_"))  # Grafiken nur temporär

    kunde = bereinige_kundendaten(json.loads(Path(kunde_json).read_text(encoding="utf-8")))
    b = parse_bericht(bericht)
    k = kennzahlen(b, kunde)
    df = parse_sweep(sweep) if sweep else None
    sz = sizing(df, kunde.get("ziel_kapitalverzinsung_pct", 15)) if df is not None else None
    kunde["_hat_sweep"] = sz is not None
    if sz is not None:
        sz["detail_konsistent"] = detail_konsistent(b, sz)
        if not sz["detail_konsistent"]:
            print("WARNUNG: Detailvariante passt nicht zum Sweep (siehe QA) -> wird im Trade-off-Chart nicht gezeigt.")

    qa = plausibilitaet(b, k, sz)
    # Fehlende Kundendaten (leer oder Vorlagen-Platzhalter) -> Bericht noch nicht versandfähig
    for pfad, wert in (("kunde.firma", auftraggeber(kunde)), ("projekt.bezeichnung", projekt(kunde)),
                       ("projekt.standort", standort(kunde)), ("bericht_datum", kunde.get("bericht_datum"))):
        if not wert:
            qa.append(f"[!! ] Kundendaten fehlen: {pfad} (im Bericht weggelassen)")
    print("\n".join(["Plausibilitätschecks:"] + qa))
    (out.with_name(out.stem + "_QA.txt")).write_text("\n".join(qa), encoding="utf-8")

    ch = {
        "donut": chart_donut(k, tmp / "donut.png", b),
        "waterfall": chart_waterfall(b, k, tmp / "waterfall.png"),
        "sens": chart_sensitivitaet(b, k, tmp / "sens.png", tuple(kunde.get("sensitivitaet_bereich_pct", [-50, 50]))),
        "schema": chart_schema(b, tmp / "schema.png"),
    }
    if kunde.get("monatsspitzen_ohne_kw") and kunde.get("monatsspitzen_mit_kw"):
        ch["spitzen"] = chart_lastspitzen(kunde["monatsspitzen_ohne_kw"], kunde["monatsspitzen_mit_kw"],
                                          tmp / "spitzen.png")
    if sz is not None:
        ch["heat"] = chart_heatmaps(sz, tmp / "heat.png")
        ch["trade"] = chart_tradeoff(sz, b, tmp / "trade.png")

    def render(target, total):
        d = Doc(target, kunde, b)
        d.total = total
        seiten = [lambda: seite_titel(d, b, k, sz, ch),
                  lambda: seite_ausgangslage(d, b, k, ch),
                  lambda: seite_methodik(d, b, k),
                  lambda: seite_wirkung(d, b, ch, kunde),
                  lambda: seite_wirtschaft(d, b, k, ch),
                  lambda: seite_risiko(d, b, k, ch)]
        if sz is not None:
            seiten.append(lambda: seite_dimensionierung(d, b, sz, ch))
        for f in seiten:
            d.new_page(); f(); d.footer()
        d.new_page()
        y_end = seite_empfehlung(d, b, k, sz, kunde)
        if y_end >= ABSCHLUSS_RESERVE:
            abschluss_block(d, kunde)
            d.footer()
        else:
            d.footer(); d.new_page(); seite_hinweise(d, kunde); d.footer()
        d.c.save()
        return d.page

    import io
    n = render(io.BytesIO(), 0)
    render(str(out), n)
    shutil.rmtree(tmp, ignore_errors=True)
    print(f"\nBericht erstellt: {out}")
    return out


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Fleco Kundenbericht Batteriespeicher")
    ap.add_argument("--bericht", required=True)
    ap.add_argument("--sweep", default=None)
    ap.add_argument("--kunde", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    build(a.bericht, a.sweep, a.kunde, a.out)
