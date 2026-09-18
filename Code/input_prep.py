"""
input_prep.py
=======================

Ergaenzt die fehlenden Zeitreihen im Batterie/PV-Analyse-Excel (input.xlsx)
und speichert das Ergebnis als input_lp.xlsx (fertig fuer den oemof.solph-Aufbau).

Was das Skript macht (siehe Funktionen unten):
  1. SwissIX Spotpreis (Zeitreihen!B), in dieser Prioritaet:
       a) aus einer lokalen CSV-Datei laden (Beats eigene, manuell gepflegte
          Datei mit historischen + kuenftigen Day-Ahead-Preisen in EUR/MWh,
          echte Europe/Zurich-Ortszeit mit Sommerzeit-Wechsel), Pfad:
          <Inputs-Ordner>/Input/ch_dayahead_2020_26_15min.csv (siehe
          SWISSIX_CSV_SUBDIR/SWISSIX_CSV_FILENAME) -- deckt die Datei den
          benoetigten Zeitraum nicht (vollstaendig) ab, wird nur die Luecke
          ueber b) gefuellt.
       b) fehlende Zeitschritte (oder die ganze Reihe, falls die Datei fehlt/
          leer ist) via ENTSO-E Transparency Platform API laden (Website).
       c) fehlt danach IMMER NOCH etwas: bereits von Hand in Zeitreihen!B
          eingetragene Werte verwenden.
     In allen drei Faellen: UTC/Europe-Zurich -> CET_FIXED konvertieren,
     EUR/MWh -> CHF/kWh konvertieren (Einheiten- UND Waehrungsumrechnung,
     via taeglichem EUR/CHF-Kurs).
  2. Bezugstarif (Zeitreihen!D) je nach Parameter!C5 (Tarifschema) berechnen:
     HT_NT (Wochentag/Uhrzeit-Schema aus Parameter!C6:C11, mit optional
     saisonalen HT/NT-Preisen [1/2/4 Werte komma-getrennt, siehe
     parse_saisonal_werte()] und optional mehreren HT-Zeitfenstern/Tag
     [Startzeit/Endzeit HT komma-getrennt]), SPOT (SwissIX +/- Aufschlag aus
     Parameter!C13), oder Zeitreihen (vorhandene historische Werte
     unveraendert uebernehmen).
     Rueckliefertarif PV (Zeitreihen!C) und Rueckliefertarif Batterie
     (Zeitreihen!K, NEU) werden UNABHAENGIG vom Bezugstarif nach je EIGENEM
     Schema berechnet (Parameter!C14 bzw. C19): Spot, Fixtarif, RMP oder
     RMP_Floor (RMP-Quartalswerte aus einer separaten Solar_RMP.xlsx) -- siehe
     resolve_rueckliefertarif(). Damit koennen PV und Batterie unterschiedlich
     vermarktet werden (z.B. PV per RMP_Floor, Batterie ohne Verguetung).
  3. PV-Produktionsprofil (Zeitreihen!E, "CH_PV_normiert_1kWp") optional via PVGIS
     laden -- braucht Standort/Ausrichtung, siehe fetch_pv_reference_profile().
  4. SRL-Preise (Zeitreihen!G/H) sind im Original-Excel bereits per Array-Formel
     aus Hilfen_Listen berechnet -- werden hier als fixe Werte "gebacken", damit
     die Endversion ohne Hilfsblatt-Referenzen auskommt.
  5. Speichert alles unter input_lp.xlsx (Parameter/Hilfen_Listen unveraendert,
     Zeitreihen ergaenzt).

WICHTIGER HINWEIS ZUR NETZWERK-ABHAENGIGKEIT:
  Schritt 1 (ENTSO-E) und Schritt 3 (PVGIS) brauchen echten Internetzugang.
  Dieses Skript wurde in einer Sandbox entwickelt, die nur auf Paket-Registries
  (PyPI etc.) zugreifen kann, nicht auf beliebige APIs. Die ENTSO-E-Anbindung ist
  daher gegen die dokumentierte API geschrieben, aber NICHT gegen eine echte
  Live-Antwort getestet worden. Bitte einen kurzen Testlauf mit einem kleinen
  Zeitraum (z.B. 2 Tage) machen, bevor das ganze Jahr geladen wird.

Benoetigte Pakete: pandas, openpyxl, requests
    pip install pandas openpyxl requests
"""

from __future__ import annotations

import datetime as dt
import os
import re
import warnings
from xml.etree import ElementTree as ET

import numpy as np
import pandas as pd
import requests
from openpyxl import load_workbook

# NEU (Beat, 18.9.2026: "wie werde ich das UserWarning los?"): openpyxl meldet
# beim Einlesen JEDER .xlsx-Datei mit Dropdown-Validierungen (Data Validation,
# z.B. Beats "Ja"/"Nein"-Felder in Parameter) diese harmlose Warnung -- die
# Validierung selbst bleibt beim reinen Lesen unberuehrt, nur openpyxls
# Faehigkeit, sie beim erneuten SPEICHERN (input_lp.xlsx) vollstaendig zu
# erhalten, ist eingeschraenkt. Das ist unkritisch (input_lp.xlsx ist eine
# reine Zwischen-Datei fuer battery_optimization.py, keine Datei, in die Beat
# spaeter noch von Hand etwas per Dropdown eintragen wuerde). Bewusst NUR
# diese eine, namentlich passende Meldung unterdrueckt (kein pauschales
# `ignore` aller UserWarnings), damit andere, tatsaechlich relevante
# Warnungen (z.B. Einheiten-Plausibilitaet, SwissIX-Fallback) weiterhin
# sichtbar bleiben.
warnings.filterwarnings(
    "ignore", message=r".*Data Validation extension is not supported.*", category=UserWarning
)

# --------------------------------------------------------------------------
# Konfiguration
# --------------------------------------------------------------------------

# Absoluter Datenordner statt relativer Pfade -- so muss die Excel-Datei NICHT
# im Git-Repo liegen. Default: derselbe Ordner wie dieses Skript (wo Inputs.xlsx
# bei euch tatsaechlich liegt); ueber die Umgebungsvariable BATTERIE_DATA_DIR
# auf einen beliebigen anderen absoluten Pfad umbiegbar, falls sich das
# spaeter aendert (z.B. `set BATTERIE_DATA_DIR=C:\Pfad\zu\Ordner` unter
# Windows, oder `export BATTERIE_DATA_DIR=/pfad/zu/ordner` unter Linux/Mac).
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.environ.get("BATTERIE_DATA_DIR", SCRIPT_DIR)

INPUT_PATH = os.path.join(DATA_DIR, "Inputs.xlsx")
OUTPUT_PATH = os.path.join(DATA_DIR, "input_lp.xlsx")

ENTSOE_API_KEY = os.environ.get(
    "ENTSOE_API_KEY", "73d39d94-710e-4a11-9160-9f8d98b6047e"
)
ENTSOE_BASE_URL = "https://web-api.tp.entsoe.eu/api"
# EIC-Code fuer die Schweizer Regelzone (Swissgrid)
ENTSOE_CH_DOMAIN = "10YCH-SWISSGRIDZ"

FX_API_URL = "https://api.frankfurter.app"  # kostenlos, kein Key noetig

# NEU: lokale SwissIX-Datei (Beats Wunsch) -- liegt in einem "Input"-
# Unterordner NEBEN der Inputs.xlsx (analog zum bereits bestehenden "Output"-
# Unterordner, siehe run_battery_analysis.py), also z.B.
#   <Inputs-Ordner>/Input/ch_dayahead_2020_26_15min.csv
# Im Unterschied zu "Output" ist dieser Ordner/diese Datei OPTIONAL -- fehlt
# sie, wird das nicht als Fehler behandelt, sondern loest einfach den
# naechsten Fallback aus (Website, siehe resolve_swissix()).
SWISSIX_CSV_SUBDIR = "Input"
SWISSIX_CSV_FILENAME = "ch_dayahead_2020_26_15min.csv"

# NEU (Beats Wunsch: RMP-Rueckliefertarif fuer PV/Batterie): RMP-Werte
# (CHF/kWh) je Jahr liegen in einer separaten Datei, analog zur SwissIX-CSV
# im selben "Input"-Unterordner, z.B.
#   <Inputs-Ordner>/Input/Solar_RMP.xlsx
# NEU (Beats Wunsch 18.9.2026: "fuer RMP und RMP mit Floor nicht den
# monatlichen RMP nehmen sondern den Quartalsweise, ich habe den im File
# ergaenzt"): Solar_RMP.xlsx enthaelt seither zusaetzlich zu den (weiterhin
# vorhandenen, aber nicht mehr gelesenen) Monatszeilen (1-12) auch vier
# Quartalszeilen ("Q1".."Q4", Q1=Jan-Maerz, Q2=Apr-Jun, Q3=Jul-Sep,
# Q4=Okt-Dez) -- siehe load_rmp_lookup()/build_rmp_series().
SOLAR_RMP_SUBDIR = "Input"
SOLAR_RMP_FILENAME = "Solar_RMP.xlsx"

# Das Excel zaehlt durchgehend 365*96 = 35040 Viertelstunden pro Jahr, ohne
# Sommerzeit-Sprung/-Dopplung zu beruecksichtigen (Kommentar im Sheet: "in CET").
# Wir behandeln "CET" hier deshalb bewusst als FESTEN Offset UTC+1 (nicht die
# echte, DST-wechselnde Europe/Zurich-Zeitzone) -- das entspricht exakt der
# Zaehlweise im Original-Excel und vermeidet Ambiguous-/Nonexistent-Zeit-Fehler
# beim Sommerzeitwechsel.
CET_FIXED = dt.timezone(dt.timedelta(hours=1))

# Spaltenlayout im Sheet "Zeitreihen" (1-indexiert wie in Excel)
COL_ZEIT = 1
COL_SWISSIX = 2
COL_RUECKLIEFER = 3  # NEU: das ist jetzt spezifisch der PV-Rueckliefertarif
COL_BEZUG = 4
COL_PV_NORMIERT = 5
COL_PV_PROFIL = 6
COL_SRL_NEG = 7
COL_SRL_POS = 8
# NEU (Beats Wunsch: eigener Rueckliefertarif fuer die Batterie, der vom
# PV-Rueckliefertarif abweichen kann): bisher gab es dafuer keine Spalte im
# Original-Excel -- wir legen deshalb eine neue Spalte K an (Header wird von
# main() automatisch gesetzt, falls noch nicht vorhanden). Spalte J
# ("Hilfsspalte") ist bereits belegt (Excel-interner XLOOKUP-Helfer fuer
# Spalte E) und wird hier bewusst NICHT wiederverwendet.
COL_RUECKLIEFER_BATTERIE = 11
COL_RUECKLIEFER_BATTERIE_HEADER = "Rückliefertarif_Batterie [CHF/kWh]"


# --------------------------------------------------------------------------
# 1. SwissIX / ENTSO-E
# --------------------------------------------------------------------------

def _strip_ns(tag: str) -> str:
    """Entfernt das XML-Namespace-Praefix von einem Tag-Namen."""
    return tag.split("}", 1)[-1] if "}" in tag else tag


def _parse_entsoe_day_ahead_xml(xml_bytes: bytes) -> pd.Series:
    """
    Parst ein ENTSO-E "Publication_MarketDocument" (Day-ahead prices, A44)
    und gibt eine pandas Series (UTC-indexiert, EUR/MWh) zurueck.
    """
    root = ET.fromstring(xml_bytes)

    # Fehlerantworten haben ein "Acknowledgement_MarketDocument" mit Reason-Text
    root_tag = _strip_ns(root.tag)
    if root_tag == "Acknowledgement_MarketDocument":
        reason = root.find(".//{*}Reason/{*}text")
        msg = reason.text if reason is not None else "unbekannter Fehler"
        raise RuntimeError(f"ENTSO-E API Fehlerantwort: {msg}")

    values = {}
    for ts in root.iter():
        if _strip_ns(ts.tag) != "TimeSeries":
            continue
        for period in ts:
            if _strip_ns(period.tag) != "Period":
                continue
            time_interval = period.find("{*}timeInterval")
            start_str = time_interval.find("{*}start").text
            resolution_str = period.find("{*}resolution").text

            start = pd.Timestamp(start_str, tz="UTC")
            resolution = pd.Timedelta(resolution_str.replace("PT", "").lower())
            # ISO 8601 Dauer wie "PT60M"/"PT15M" -> pandas Timedelta braucht
            # z.B. "60min"/"15min"; obiges .replace ist ein einfacher Spezialfall.
            match = re.match(r"PT(\d+)M", resolution_str)
            if match:
                resolution = pd.Timedelta(minutes=int(match.group(1)))

            for point in period.iter():
                if _strip_ns(point.tag) != "Point":
                    continue
                pos = int(point.find("{*}position").text)
                price = float(point.find("{*}price.amount").text)
                ts_utc = start + (pos - 1) * resolution
                values[ts_utc] = price

    if not values:
        raise RuntimeError(
            "Keine Preis-Punkte in der ENTSO-E-Antwort gefunden -- "
            "moeglich, dass fuer die Schweizer Zone (10YCH-SWISSGRIDZ) "
            "keine Day-Ahead-Preise auf ENTSO-E publiziert werden. "
            "Bitte pruefen (SwissIX wird von EPEX Spot separat berechnet "
            "und ist evtl. nicht 1:1 identisch mit den ENTSO-E-Daten)."
        )

    series = pd.Series(values).sort_index()
    series.index.name = "time_utc"
    return series


def fetch_entsoe_day_ahead_prices(
    start: pd.Timestamp, end: pd.Timestamp, api_key: str = ENTSOE_API_KEY
) -> pd.Series:
    """
    Laedt Day-Ahead-Preise (EUR/MWh) fuer die Schweiz von ENTSO-E, chunked
    nach Monat (ENTSO-E limitiert die Anfragegroesse). Rueckgabe: Series,
    UTC-indexiert.
    """
    chunks = []
    cur = start
    while cur < end:
        chunk_end = min(cur + pd.DateOffset(months=1), end)
        params = {
            "documentType": "A44",
            "in_Domain": ENTSOE_CH_DOMAIN,
            "out_Domain": ENTSOE_CH_DOMAIN,
            "periodStart": cur.strftime("%Y%m%d%H%M"),
            "periodEnd": chunk_end.strftime("%Y%m%d%H%M"),
            "securityToken": api_key,
        }
        resp = requests.get(ENTSOE_BASE_URL, params=params, timeout=60)
        resp.raise_for_status()
        chunks.append(_parse_entsoe_day_ahead_xml(resp.content))
        cur = chunk_end

    combined = pd.concat(chunks).sort_index()
    # Monatsgrenzen ueberlappen sich haeufig um genau einen Zeitpunkt (der
    # Endzeitpunkt eines Chunks ist zugleich der Startzeitpunkt des naechsten)
    # -- das erzeugt doppelte Indexwerte, die .reindex() weiter unten zum
    # Absturz bringen wuerden ("cannot reindex on an axis with duplicate
    # labels"). Daher hier dedupliziseren, bevor wir den Index weiterreichen.
    combined = combined[~combined.index.duplicated(keep="first")]
    return combined


def fetch_eur_chf_fx_rates(start_date: str, end_date: str) -> pd.Series:
    """Taegliche EUR/CHF-Kurse (Frankfurter/ECB-Referenz), noetig fuer die
    Umrechnung von EUR/MWh (ENTSO-E) nach CHF/kWh."""
    url = f"{FX_API_URL}/{start_date}..{end_date}"
    resp = requests.get(url, params={"from": "EUR", "to": "CHF"}, timeout=30)
    resp.raise_for_status()
    data = resp.json()["rates"]
    fx = pd.Series({pd.Timestamp(d): v["CHF"] for d, v in data.items()}).sort_index()
    fx = fx[~fx.index.duplicated(keep="first")]
    return fx


def build_swissix_chf_per_kwh(
    target_index: pd.DatetimeIndex, api_key: str = ENTSOE_API_KEY
) -> pd.Series:
    """
    Baut die SwissIX-Preisreihe in CHF/kWh auf dem Zielindex (Europe/Zurich,
    15-Min-Aufloesung) auf: ENTSO-E (EUR/MWh, UTC) laden, auf Zielindex
    ausrichten (ffill fuer stuendliche Ausgangsdaten), dann mit dem
    Tages-EUR/CHF-Kurs in CHF/kWh umrechnen.
    """
    start_utc = target_index[0].tz_convert("UTC")
    end_utc = target_index[-1].tz_convert("UTC") + pd.Timedelta(minutes=15)

    prices_eur_mwh = fetch_entsoe_day_ahead_prices(start_utc, end_utc, api_key)
    prices_eur_mwh.index = prices_eur_mwh.index.tz_convert(CET_FIXED)

    # Auf den Zielindex ausrichten (falls ENTSO-E nur stuendlich liefert,
    # werden die 4 Viertelstunden je Stunde mit demselben Preis befuellt)
    aligned = prices_eur_mwh.reindex(target_index, method="ffill")

    fx = fetch_eur_chf_fx_rates(
        target_index[0].strftime("%Y-%m-%d"), target_index[-1].strftime("%Y-%m-%d")
    )
    fx_daily = fx.reindex(target_index.normalize().tz_localize(None), method="ffill")
    fx_daily.index = target_index

    price_chf_kwh = aligned * fx_daily.values / 1000.0
    price_chf_kwh.name = "SwissIX[CHF/kWh]"
    return price_chf_kwh


def load_swissix_csv_chf_per_kwh(
    csv_path: str, target_index: pd.DatetimeIndex
) -> pd.Series | None:
    """
    Laedt SwissIX/Day-Ahead-Spotpreise aus einer lokalen CSV-Datei (Beats
    eigene, manuell gepflegte Datei mit historischen + kuenftigen Preisen)
    und rechnet sie von EUR/MWh nach CHF/kWh um -- inhaltlich dasselbe wie
    build_swissix_chf_per_kwh() oben, nur mit der Datei statt der ENTSO-E-
    API als Rohdatenquelle.

    Erwartetes Format: Spalten "timestamp" (NAIV, aber echte Europe/Zurich-
    ORTSZEIT mit Sommerzeit-Wechsel -- verifiziert an Beats Datei: am
    29.3. springt die Uhr von 01:45 direkt auf 03:00 [02:00-02:45 fehlen],
    am 25.10. erscheint 02:00-02:45 zweimal hintereinander, genau wie bei
    einer echten Zuercher Uhr -- NICHT der feste CET_FIXED-Offset ohne
    Sommerzeit, den der Rest dieses Moduls fuer die Excel-Zeitspalte
    verwendet, siehe Modul-Header) und "price_eur_mwh".

    WICHTIG zur Umrechnung auf `target_index`: `target_index` ist intern
    unter der Fixed-CET-Konvention aufgebaut (jede Zeile = ein bestimmter
    ABSOLUTER Zeitpunkt, dargestellt mit konstantem UTC+1-Offset). Die
    CSV-Zeitstempel werden hier zuerst korrekt (mit Sommerzeit) nach
    Europe/Zurich lokalisiert -- das ergibt fuer jede Zeile den korrekten
    ABSOLUTEN Zeitpunkt -- und danach auf dieselbe Fixed-CET-Darstellung
    umgerechnet. Der Abgleich mit `target_index` (via .reindex()) erfolgt
    dann ueber den absoluten Zeitpunkt, nicht ueber die angezeigte Uhrzeit
    -- exakt dasselbe Prinzip wie beim UTC-Abgleich in
    build_swissix_chf_per_kwh() oben.

    Rueckgabe:
      - None, wenn die Datei fehlt, leer ist, oder nach dem Ausrichten auf
        `target_index` KEINEN einzigen Wert liefert (= fuer den angefragten
        Zeitraum komplett "leer") -- der Aufrufer soll dann komplett auf
        die Website (ENTSO-E) ausweichen.
      - Eine Series auf `target_index`, ggf. mit einzelnen NaN-Luecken (z.B.
        wenn der Szenario-Zeitraum ueber das Ende der Datei hinausreicht),
        falls zumindest TEILWEISE Daten vorhanden sind.

    Wirft ValueError, wenn die Datei zwar existiert, aber strukturell nicht
    lesbar ist (falsche Spalten, nicht eindeutig interpretierbare
    Zeitstempel) -- das ist ein echter Datei-Fehler, kein simples "leer",
    und soll dem Nutzer sichtbar gemeldet werden statt es zu uebergehen.
    """
    if not os.path.isfile(csv_path):
        return None

    try:
        df = pd.read_csv(csv_path)
    except Exception as exc:  # noqa: BLE001
        raise ValueError(f"Datei konnte nicht gelesen werden ({exc})") from exc

    if df.empty:
        return None

    missing_cols = {"timestamp", "price_eur_mwh"} - set(df.columns)
    if missing_cols:
        raise ValueError(
            f"erwartete Spalten fehlen ({sorted(missing_cols)}) -- "
            "erwartet werden 'timestamp' (Europe/Zurich-Ortszeit, naiv) "
            "und 'price_eur_mwh'."
        )

    try:
        naive = pd.DatetimeIndex(pd.to_datetime(df["timestamp"]))
        local_index = naive.tz_localize("Europe/Zurich", ambiguous="infer")
    except Exception as exc:  # noqa: BLE001
        raise ValueError(
            f"Zeitstempel in Spalte 'timestamp' nicht eindeutig als "
            f"Europe/Zurich-Ortszeit interpretierbar ({exc}) -- bitte "
            "pruefen, ob die Datei lueckenlos und chronologisch sortiert "
            "ist (v.a. rund um die Sommerzeit-Umstellungen im Maerz/Oktober)."
        ) from exc

    prices_eur_mwh = pd.Series(
        df["price_eur_mwh"].to_numpy(dtype=float), index=local_index.tz_convert(CET_FIXED)
    )
    prices_eur_mwh = prices_eur_mwh[~prices_eur_mwh.index.duplicated(keep="first")].sort_index()

    aligned_eur_mwh = prices_eur_mwh.reindex(target_index)
    if aligned_eur_mwh.isna().all():
        return None

    covered = aligned_eur_mwh.dropna()
    fx = fetch_eur_chf_fx_rates(
        covered.index[0].strftime("%Y-%m-%d"), covered.index[-1].strftime("%Y-%m-%d")
    )
    fx_daily = fx.reindex(target_index.normalize().tz_localize(None), method="ffill")
    fx_daily.index = target_index

    price_chf_kwh = aligned_eur_mwh * fx_daily.values / 1000.0
    price_chf_kwh.name = "SwissIX[CHF/kWh] (aus Datei)"
    return price_chf_kwh


def resolve_swissix(
    index: pd.DatetimeIndex,
    csv_path: str,
    fetch_swissix: bool,
    existing_swissix: pd.Series,
    api_key: str = ENTSOE_API_KEY,
) -> tuple[pd.Series | None, str | None]:
    """
    Ermittelt die SwissIX-Preisreihe (CHF/kWh) auf `index`, nach Beats
    gewuenschter Prioritaet:
      1. Lokale CSV-Datei (`csv_path`, siehe load_swissix_csv_chf_per_kwh()).
      2. Fehlt die Datei/deckt sie den Zeitraum nicht vollstaendig ab: die
         fehlenden Zeitschritte automatisch von der ENTSO-E-Website
         nachladen (sofern `fetch_swissix=True`).
      3. Fehlt danach IMMER NOCH etwas: auf bereits von Hand in Zeitreihen!B
         eingetragene Werte zurueckfallen (nur fuer die Zeitschritte, die
         dort tatsaechlich befuellt sind).

    Gibt (swissix_oder_None, quelle_beschreibung_oder_None) zurueck.
    `swissix` ist `None`, wenn WEDER Datei NOCH Website NOCH manueller
    Fallback auch nur einen einzigen Wert liefern konnte -- build_tariffs()
    wirft in diesem Fall fuer Tarifschema=SPOT einen klaren Fehler.
    """
    swissix: pd.Series | None = None
    source_notes: list[str] = []

    # --- 1. Lokale CSV-Datei ------------------------------------------
    try:
        csv_series = load_swissix_csv_chf_per_kwh(csv_path, index)
    except ValueError as exc:
        print(f"  WARNUNG: SwissIX-Datei {csv_path} nicht nutzbar: {exc}")
        csv_series = None
    except Exception as exc:  # noqa: BLE001 -- z.B. FX-Kurs-Abruf fehlgeschlagen
        print(f"  WARNUNG: SwissIX-Datei {csv_path} konnte nicht verarbeitet werden ({exc}).")
        csv_series = None

    if csv_series is None:
        print(
            f"  SwissIX-Datei {csv_path} nicht gefunden/leer/nutzbar fuer "
            "den benoetigten Zeitraum -- weiche auf Website (ENTSO-E) aus."
        )
    else:
        n_missing = int(csv_series.isna().sum())
        if n_missing == 0:
            print(f"  SwissIX vollstaendig aus Datei geladen: {csv_path}")
            source_notes.append(f"Datei ({SWISSIX_CSV_FILENAME})")
        else:
            print(
                f"  SwissIX-Datei deckt {len(csv_series) - n_missing} von "
                f"{len(csv_series)} Zeitschritten ab -- {n_missing} fehlen "
                "(z.B. Zeitraum ueber das Dateiende hinaus). Versuche "
                "Luecke ueber die Website zu fuellen ..."
            )
            source_notes.append(f"Datei ({SWISSIX_CSV_FILENAME}, {n_missing} Zeitschritte fehlen)")
        swissix = csv_series

    # --- 2. Website (ENTSO-E) -- fuer komplett fehlende ODER Rest-Luecken --
    missing_index = index if swissix is None else index[swissix.isna()]
    if len(missing_index) > 0:
        if fetch_swissix:
            try:
                print(f"Lade {len(missing_index)} SwissIX-Zeitschritt(e) von ENTSO-E (Website) ...")
                website_series = build_swissix_chf_per_kwh(missing_index, api_key)
                if swissix is None:
                    swissix = pd.Series(index=index, dtype=float)
                swissix.loc[website_series.index] = website_series.values
                print("  -> Website-Fetch erfolgreich.")
                if not source_notes:
                    source_notes.append("automatisch (ENTSO-E API)")
                elif swissix.isna().sum() < len(missing_index):
                    source_notes.append("Rest via ENTSO-E API")
            except Exception as exc:  # noqa: BLE001
                warnings.warn(
                    f"SwissIX-Website-Fetch fehlgeschlagen ({exc}). Bitte pruefen: (1) Kann "
                    "diese URL direkt im Browser auf demselben Rechner geoeffnet "
                    "werden? https://web-api.tp.entsoe.eu/api?documentType=A44&"
                    "in_Domain=10YCH-SWISSGRIDZ&out_Domain=10YCH-SWISSGRIDZ&"
                    "periodStart=202501010000&periodEnd=202501020000&"
                    "securityToken=<key> -- zeigt der Browser eine XML-Antwort von "
                    "ENTSO-E (auch eine Fehlermeldung zaehlt), liegt es an den "
                    "Abfrage-Parametern; zeigt er eine generische "
                    "Fehler-/Blockseite, blockiert vermutlich ein Firmen-Proxy/"
                    "-Firewall den Zugriff auf web-api.tp.entsoe.eu. (2) Der "
                    "verwendete API-Key ist aktuell ein geteilter Default-Key im "
                    "Code -- ein eigener, kostenloser ENTSO-E-Account-Key (siehe "
                    "transparencyplatform.zendesk.com) via Umgebungsvariable "
                    "ENTSOE_API_KEY ist zuverlaessiger."
                )
        else:
            print("  SwissIX-Website-Fetch uebersprungen (fetch_swissix=False).")

    # --- 3. Manueller Fallback: bereits in Zeitreihen!B vorhandene Werte --
    remaining_missing = index if swissix is None else index[swissix.isna()]
    if len(remaining_missing) > 0:
        manual_fill = existing_swissix.reindex(remaining_missing).dropna()
        if len(manual_fill) > 0:
            print(
                "  -> verwende zusaetzlich bereits in Zeitreihen!B vorhandene "
                f"Werte fuer {len(manual_fill)} verbleibende Zeitschritt(e)."
            )
            if swissix is None:
                swissix = pd.Series(index=index, dtype=float)
            swissix.loc[manual_fill.index] = manual_fill.values
            source_notes.append("manuell (Zeitreihen im Inputfile)")

    if swissix is not None and swissix.notna().any():
        quelle = " + ".join(source_notes) if source_notes else "unbekannt"
        return swissix, quelle
    return None, None


# --------------------------------------------------------------------------
# 2. HT/NT- bzw. SPOT- bzw. Zeitreihen-Tarif (Bezug), sowie separates
#    Rueckliefertarif-Schema fuer PV und Batterie (Beats Erweiterung:
#    "PV mit Spotpreis, Bezug per HT/NT" + Quartals-/Saison-Preise +
#    mehrere HT-Fenster/Tag + RMP-Rueckliefertarif)
# --------------------------------------------------------------------------

def parse_saisonal_werte(raw, feld_name: str) -> list[float]:
    """
    Parst eine Parameter-Zelle mit 1, 2 oder 4 (komma-getrennten) Preisen:
      - 1 Wert  -> ganzjaehrig derselbe Preis.
      - 2 Werte -> [Winterhalbjahr (Okt-Maer), Sommerhalbjahr (Apr-Sep)].
      - 4 Werte -> [Q1 (Jan-Maer), Q2 (Apr-Jun), Q3 (Jul-Sep), Q4 (Okt-Dez)].
    Leerzeichen um die Kommas werden toleriert (Beats Beispiel: "0.2159,0.1965, 0.2159,0.1965").
    """
    if raw is None:
        raise ValueError(f"{feld_name} ist leer -- bitte 1, 2 oder 4 komma-getrennte Preise eintragen.")
    if isinstance(raw, (int, float)):
        return [float(raw)]
    parts = [p.strip() for p in str(raw).split(",")]
    try:
        werte = [float(p) for p in parts]
    except ValueError as exc:
        raise ValueError(
            f"{feld_name} = {raw!r} enthaelt einen nicht-numerischen Wert -- "
            "erwartet werden 1, 2 oder 4 komma-getrennte Zahlen (z.B. '0.15' "
            "oder '0.20,0.15' oder '0.22,0.18,0.16,0.20')."
        ) from exc
    if len(werte) not in (1, 2, 4):
        raise ValueError(
            f"{feld_name} = {raw!r} hat {len(werte)} komma-getrennte Werte -- "
            "erwartet werden genau 1 (ganzjaehrig), 2 (Winter/Sommer) oder "
            "4 (Q1/Q2/Q3/Q4)."
        )
    return werte


def saisonal_werte_zu_serie(werte: list[float], index: pd.DatetimeIndex) -> pd.Series:
    """Bildet eine Liste von 1/2/4 Werten (siehe parse_saisonal_werte()) auf
    eine Zeitreihe ab. Winterhalbjahr = Okt-Maer, Sommerhalbjahr = Apr-Sep
    (Schweizer Netz-/Energietarif-Konvention). Quartale = Kalenderquartale."""
    n = len(werte)
    if n == 1:
        return pd.Series(werte[0], index=index, dtype=float)
    month = index.month
    if n == 2:
        ist_winter = month.isin([10, 11, 12, 1, 2, 3])
        vals = pd.Series(ist_winter, index=index).map({True: werte[0], False: werte[1]})
        return vals.astype(float)
    # n == 4: Quartal aus dem Monat ableiten (Jan-Maer=Q1 ... Okt-Dez=Q4)
    quartal_idx = (month - 1) // 3  # 0..3
    vals = pd.Series(quartal_idx, index=index).map({i: werte[i] for i in range(4)})
    return vals.astype(float)


def parse_zeitfenster_liste(raw, feld_name: str) -> list[dt.time]:
    """Parst eine Startzeit/Endzeit-HT-Zelle mit einem oder mehreren (komma-
    getrennten) HH:MM-Zeitpunkten (Beats Wunsch: manchmal 2 HT-Fenster/Tag,
    z.B. Startzeit='07:00, 16:00'). Excel liefert einzelne Zeitwerte oft
    bereits als datetime.time/datetime.datetime -- wird direkt uebernommen."""
    if raw is None:
        raise ValueError(f"{feld_name} ist leer.")
    if isinstance(raw, dt.time):
        return [raw]
    if isinstance(raw, dt.datetime):
        return [raw.time()]
    result = []
    for teil in str(raw).split(","):
        teil = teil.strip()
        try:
            h, m = teil.split(":")
            result.append(dt.time(int(h), int(m)))
        except Exception as exc:  # noqa: BLE001
            raise ValueError(
                f"{feld_name} enthaelt {teil!r} -- erwartet wird ein HH:MM-Zeitwert "
                f"(ggf. mehrere, komma-getrennt fuer mehrere Fenster/Tag): {raw!r}."
            ) from exc
    return result


def build_ht_nt_bezugstarif(
    index: pd.DatetimeIndex,
    ht_werte: list[float],
    nt_werte: list[float],
    tage_ht: int,
    start_ht_liste: list[dt.time],
    end_ht_liste: list[dt.time],
) -> pd.Series:
    """
    Berechnet den Bezugstarif je Zeitschritt nach HT/NT-Schema.
    tage_ht: Anzahl Tage/Woche mit HT, gezaehlt ab Montag (z.B. 5 -> Mo-Fr HT).
    HT gilt an diesen Tagen in JEDEM der Fenster [start_ht_liste[i], end_ht_liste[i]),
    sonst NT (Beats Wunsch: teils 2 HT-Fenster/Tag). HT-/NT-Preis koennen ihrerseits
    saisonal variieren (1/2/4 Werte, siehe saisonal_werte_zu_serie()).
    """
    if len(start_ht_liste) != len(end_ht_liste):
        raise ValueError(
            f"Startzeit HT hat {len(start_ht_liste)} Zeitfenster, Endzeit HT "
            f"hat {len(end_ht_liste)} -- die Anzahl muss uebereinstimmen."
        )
    is_ht_day = index.weekday < tage_ht
    times = index.time
    is_ht_time = np.zeros(len(index), dtype=bool)
    for start, end in zip(start_ht_liste, end_ht_liste):
        is_ht_time |= (times >= start) & (times < end)
    is_ht = is_ht_day & is_ht_time

    ht_serie = saisonal_werte_zu_serie(ht_werte, index)
    nt_serie = saisonal_werte_zu_serie(nt_werte, index)
    bezug = pd.Series(np.where(is_ht, ht_serie.values, nt_serie.values), index=index)
    bezug.name = "Bezugstarif [CHF/kWh]"
    return bezug


def load_rmp_lookup(path: str) -> dict[tuple[int, int], float]:
    """
    Liest Beats "Solar_RMP.xlsx" (Sheet 'RMP_Solar'): Kopfzeile (Zeile 2) =
    Jahre. Darunter stehen Zeilenweise sowohl Monatswerte (Spalte A = 1-12)
    als auch, NEU (Beats Wunsch 18.9.2026: "fuer RMP und RMP mit Floor nicht
    den monatlichen RMP nehmen sondern den Quartalsweise, ich habe den im
    File ergaenzt"), Quartalswerte (Spalte A = "Q1".."Q4", Text). Der
    RMP-Rueckliefertarif nutzt seither NUR NOCH die Quartals-Zeilen (Q1=Jan-
    Maerz, Q2=Apr-Jun, Q3=Jul-Sep, Q4=Okt-Dez) -- die Monatszeilen bleiben
    zwar in der Datei stehen (evtl. fuer andere Auswertungen), werden hier
    aber bewusst NICHT mehr gelesen. Nicht jedes Jahr deckt jedes Quartal ab
    (z.B. weil die Datenreihe erst waehrend eines Jahres beginnt) -- nur
    tatsaechlich befuellte Zellen landen im Ergebnis-dict.

    Gibt {(jahr, quartal): preis_chf_kwh} zurueck (quartal in 1..4). Wirft
    FileNotFoundError, falls `path` nicht existiert (RMP-Datei ist -- anders
    als die SwissIX-CSV -- nicht "optional": wird sie fuer
    Parameter!C14/C19='RMP'/'RMP_Floor' gebraucht aber fehlt, ist das ein
    echter Fehler, kein stiller Fallback).
    """
    if not os.path.isfile(path):
        raise FileNotFoundError(
            f"Solar_RMP.xlsx nicht gefunden unter {path!r} -- wird fuer das "
            "Rueckliefertarif-Schema RMP/RMP_Floor benoetigt. Bitte Datei an "
            "diesem Pfad bereitstellen oder ein anderes Schema waehlen."
        )
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb["RMP_Solar"] if "RMP_Solar" in wb.sheetnames else wb.active

    header = [c.value for c in next(ws.iter_rows(min_row=2, max_row=2))]
    jahr_spalten = {
        idx: int(jahr) for idx, jahr in enumerate(header) if isinstance(jahr, (int, float))
    }
    if not jahr_spalten:
        raise ValueError(
            f"Solar_RMP.xlsx ({path!r}): in Zeile 2 wurden keine Jahres-Spalten "
            "gefunden -- erwartetes Format: Spalte A=Monat (1-12) bzw. "
            "Quartal ('Q1'..'Q4'), Zeile 2 je weitere Spalte ein Jahr (z.B. "
            "2024, 2025, ...)."
        )

    lookup: dict[tuple[int, int], float] = {}
    quartal_je_label = {"Q1": 1, "Q2": 2, "Q3": 3, "Q4": 4}
    for row in ws.iter_rows(min_row=3):
        label_wert = row[0].value
        if not isinstance(label_wert, str):
            continue
        quartal = quartal_je_label.get(label_wert.strip().upper())
        if quartal is None:
            continue
        for idx, jahr in jahr_spalten.items():
            if idx >= len(row):
                continue
            preis = row[idx].value
            if isinstance(preis, (int, float)):
                lookup[(jahr, quartal)] = float(preis)
    if not lookup:
        raise ValueError(
            f"Solar_RMP.xlsx ({path!r}): keine Quartalszeilen ('Q1'..'Q4' in "
            "Spalte A) mit Werten gefunden -- seit Beats Umstellung auf "
            "Quartals-RMP wird nur noch dieses Format gelesen (die frueheren "
            "Monatszeilen 1-12 reichen alleine nicht mehr)."
        )
    return lookup


def build_rmp_series(index: pd.DatetimeIndex, rmp_lookup: dict, label: str) -> pd.Series:
    """Baut aus dem {(jahr,quartal): preis}-Lookup (siehe load_rmp_lookup())
    eine Zeitreihe auf `index` -- jeder Zeitschritt wird ueber sein
    Kalenderquartal (Q1=Jan-Maerz, Q2=Apr-Jun, Q3=Jul-Sep, Q4=Okt-Dez, NEU
    statt bisher ueber den Monat) nachgeschlagen; alle Zeitschritte
    desselben Quartals/Jahres erhalten denselben Wert. Fehlt fuer ein im
    Szenario vorkommendes Quartal/Jahr ein RMP-Wert, wird ein klarer Fehler
    geworfen (kein stilles Extrapolieren/Nullsetzen -- ein falscher
    RMP-Wert wuerde direkt die Wirtschaftlichkeitsrechnung verfaelschen)."""
    jahr_quartal = [(jahr, (monat - 1) // 3 + 1) for jahr, monat in zip(index.year, index.month)]
    fehlende = sorted({(j, q) for (j, q) in jahr_quartal if (j, q) not in rmp_lookup})
    if fehlende:
        fehlende_str = ", ".join(f"Q{q}/{j}" for j, q in fehlende)
        raise ValueError(
            f"Rueckliefertarif_{label}=RMP(_Floor) braucht Quartals-RMP-Werte "
            f"aus Solar_RMP.xlsx (Zeilen 'Q1'..'Q4') fuer folgende Quartale, "
            f"die dort fehlen: {fehlende_str}. Bitte Solar_RMP.xlsx ergaenzen "
            "oder fuer diesen Zeitraum ein anderes Schema waehlen."
        )
    werte = [rmp_lookup[(j, q)] for j, q in jahr_quartal]
    return pd.Series(werte, index=index)


def resolve_rueckliefertarif(
    schema,
    index: pd.DatetimeIndex,
    swissix: pd.Series | None,
    spot_abschlag_lieferung: float,
    hkn_werte: list[float],
    floor_value: float,
    fixtarif_value: float,
    rmp_lookup: dict | None,
    label: str,
) -> pd.Series:
    """
    Berechnet den Rueckliefertarif (PV oder Batterie, je nach `label`) nach
    dem gewaehlten Schema (Parameter!C14 bzw. C19):
      - "Spot":      SwissIX +/- Abschlag_Lieferung, OHNE HKN (Beats Vorgabe:
                     "Spot ist Spot ohne HKN").
      - "Fixtarif":  fester Preis + HKN.
      - "RMP":       RMP-Quartalswert (aus Solar_RMP.xlsx, Zeilen "Q1".."Q4",
                     NEU seit 18.9.2026 statt zuvor Monatswert) + HKN.
      - "RMP_Floor": max(RMP-Quartalswert, Floor) + HKN.
    HKN kann seinerseits saisonal variieren (1/2/4 Werte, wie HT/NT).
    """
    normalisiert = str(schema).strip() if schema is not None else ""

    if normalisiert == "Spot":
        if swissix is None:
            raise ValueError(
                f"Rueckliefertarif_{label}=Spot braucht die SwissIX-Zeitreihe "
                "(Spalte B), die aber nicht ermittelt werden konnte (siehe "
                "Warnung(en) weiter oben zu SwissIX-CSV/Website/manuellem "
                "Fallback)."
            )
        # Gleiche |swissix|-Korrektur wie beim Bezugstarif (siehe dortiger
        # Kommentar) -- verhindert, dass ein negativer Spotpreis die Wirkung
        # des Abschlags umkehrt.
        return swissix - swissix.abs() * (spot_abschlag_lieferung / 100.0)

    hkn_serie = saisonal_werte_zu_serie(hkn_werte, index)

    if normalisiert == "Fixtarif":
        return pd.Series(fixtarif_value, index=index, dtype=float) + hkn_serie

    if normalisiert == "RMP":
        if rmp_lookup is None:
            raise ValueError(
                f"Rueckliefertarif_{label}=RMP braucht Solar_RMP.xlsx, die "
                "aber nicht gefunden/geladen werden konnte (siehe Warnung "
                "weiter oben)."
            )
        return build_rmp_series(index, rmp_lookup, label) + hkn_serie

    if normalisiert == "RMP_Floor":
        if rmp_lookup is None:
            raise ValueError(
                f"Rueckliefertarif_{label}=RMP_Floor braucht Solar_RMP.xlsx, "
                "die aber nicht gefunden/geladen werden konnte (siehe Warnung "
                "weiter oben)."
            )
        rmp_serie = build_rmp_series(index, rmp_lookup, label)
        return rmp_serie.clip(lower=floor_value) + hkn_serie

    raise ValueError(
        f"Unbekanntes Rueckliefertarif_{label}-Schema: {schema!r} -- erwartet "
        "wird eines von 'Spot', 'Fixtarif', 'RMP', 'RMP_Floor'."
    )


def build_tariffs(
    index: pd.DatetimeIndex,
    params: dict,
    swissix: pd.Series | None,
    existing_bezug: pd.Series,
    rmp_lookup: dict | None = None,
) -> tuple[pd.Series, pd.Series, pd.Series]:
    """
    Gibt (rueckliefertarif_pv, rueckliefertarif_batterie, bezugstarif) als
    Series auf `index` zurueck.
      - Bezugstarif: nach Parameter!C5 (Tarifschema) -- HT_NT, SPOT oder
        Zeitreihen (unveraendert wie bisher, betrifft NUR den Netzbezug).
      - Rueckliefertarif PV/Batterie: je EIGENEM Schema (Parameter!C14 bzw.
        C19: Spot/Fixtarif/RMP/RMP_Floor), UNABHAENGIG vom Bezugsschema
        (Beats Erweiterung: z.B. Bezug=HT_NT, PV-Rueckliefertarif=RMP_Floor).
        Ist beides auf "Spot"/"SPOT" gesetzt, entspricht das dem alten
        Alles-Spot-Verhalten (daher der Excel-Kommentar "Bei SPOT wird auch
        der Rueckliefertarif zu SPOT" -- keine gesonderte Kopplung noetig,
        ergibt sich von selbst aus den unabhaengigen Berechnungen).
    """
    schema = params["tarifschema"]

    if schema == "HT_NT":
        ht_werte = parse_saisonal_werte(params["ht_raw"], "Parameter!C6 (HT)")
        nt_werte = parse_saisonal_werte(params["nt_raw"], "Parameter!C7 (NT)")
        start_ht_liste = parse_zeitfenster_liste(params["start_ht_raw"], "Parameter!C10 (Startzeit HT)")
        end_ht_liste = parse_zeitfenster_liste(params["end_ht_raw"], "Parameter!C11 (Endzeit HT)")
        bezug = build_ht_nt_bezugstarif(
            index, ht_werte, nt_werte, params["tage_ht"], start_ht_liste, end_ht_liste
        )

    elif schema == "SPOT":
        if swissix is None:
            raise ValueError(
                "Tarifschema=SPOT braucht die SwissIX-Zeitreihe (Spalte B) -- "
                "diese ist weder aus der lokalen CSV-Datei "
                f"(<Inputs-Ordner>/{SWISSIX_CSV_SUBDIR}/{SWISSIX_CSV_FILENAME}) "
                "noch automatisch von der ENTSO-E-Website (Fetch "
                "fehlgeschlagen oder uebersprungen, siehe Warnung weiter oben) "
                "noch von Hand in Zeitreihen!B vorhanden. Entweder (a) die "
                "CSV-Datei an obigem Pfad bereitstellen/ergaenzen, (b) die "
                "Netzwerk-/Key-Probleme fuer den Website-Fetch beheben (siehe "
                "Warnung oben) und erneut laufen lassen, oder (c) Tagespreise "
                "von Hand in Zeitreihen!B eintragen und danach erneut laufen "
                "lassen."
            )
        # WICHTIG (Bugfix, siehe Bugfix-Log): Auf-/Abschlag ueber |swissix|
        # statt ueber den signierten Spotpreis selbst. Mit der urspruenglichen
        # Formel (bezug = swissix * (1 + aufschlag)) kehrt sich bei NEGATIVEM
        # Spotpreis die Wirkung um: ein "Aufschlag" macht den Bezug dann
        # GUENSTIGER statt teurer. Mit |swissix| wirkt der Aufschlag IMMER in
        # die gleiche wirtschaftliche Richtung wie bei positivem Spotpreis.
        if params["spot_aufschlag_bezug"] < 0:
            warnings.warn(
                "Parameter!C13 (Aufschlag Bezug) ist negativ -- das "
                "widerspricht der ueblichen Bedeutung als Zuschlag. Bitte "
                "pruefen, ob das so gewollt ist."
            )
        bezug = swissix + swissix.abs() * (params["spot_aufschlag_bezug"] / 100.0)

    elif schema == "Zeitreihen":
        if existing_bezug.isna().all():
            warnings.warn(
                "Tarifschema=Zeitreihen erwartet bereits vorhandene historische "
                "Werte in Spalte D -- diese sind aber leer. Bitte den "
                "historischen Bezugstarif von Hand ergaenzen; das Skript kann "
                "diesen nicht herleiten."
            )
        bezug = existing_bezug

    else:
        raise ValueError(f"Unbekanntes Tarifschema: {schema!r}")

    if params["spot_abschlag_lieferung"] < 0:
        warnings.warn(
            "Parameter!C12 (Abschlag Lieferung) ist negativ -- das "
            "widerspricht der ueblichen Bedeutung als Abschlag und betrifft "
            "jedes Rueckliefertarif-Schema, das 'Spot' waehlt. Bitte pruefen, "
            "ob das so gewollt ist."
        )

    rueckliefer_pv = resolve_rueckliefertarif(
        schema=params["rueckliefer_pv_schema"],
        index=index,
        swissix=swissix,
        spot_abschlag_lieferung=params["spot_abschlag_lieferung"],
        hkn_werte=parse_saisonal_werte(params["rueckliefer_pv_hkn_raw"], "Parameter!C15 (HKN PV)"),
        floor_value=params["rueckliefer_pv_floor"],
        fixtarif_value=params["rueckliefer_pv_fixtarif"],
        rmp_lookup=rmp_lookup,
        label="PV",
    )
    rueckliefer_batterie = resolve_rueckliefertarif(
        schema=params["rueckliefer_batterie_schema"],
        index=index,
        swissix=swissix,
        spot_abschlag_lieferung=params["spot_abschlag_lieferung"],
        hkn_werte=parse_saisonal_werte(params["rueckliefer_batterie_hkn_raw"], "Parameter!C20 (HKN Batterie)"),
        floor_value=params["rueckliefer_batterie_floor"],
        fixtarif_value=params["rueckliefer_batterie_fixtarif"],
        rmp_lookup=rmp_lookup,
        label="Batterie",
    )
    return rueckliefer_pv, rueckliefer_batterie, bezug


# --------------------------------------------------------------------------
# 3. PV-Referenzprofil (optional, braucht Standortangaben)
# --------------------------------------------------------------------------

def fetch_pv_reference_profile(
    index: pd.DatetimeIndex,
    lat: float,
    lon: float,
    tilt: float = 30,
    azimuth: float = 0,
) -> pd.Series:
    """
    Laedt ein normiertes PV-Erzeugungsprofil (kW pro 1 kWp) via PVGIS
    (EU Joint Research Centre, kostenlos, kein Key noetig) fuer den
    angegebenen Standort. NICHT automatisch aufgerufen -- braucht lat/lon,
    die noch nicht bekannt sind (siehe Kommentar "-> mail Urs/Astrid" im
    Excel). Erst nutzen, wenn Standort/Ausrichtung final feststehen.
    """
    url = "https://re.jrc.ec.europa.eu/api/v5_2/seriescalc"
    params = {
        "lat": lat,
        "lon": lon,
        "startyear": index[0].year,
        "endyear": index[-1].year,
        "pvcalculation": 1,
        "peakpower": 1,  # normiert auf 1 kWp
        "loss": 14,  # Systemverluste in %, Standardannahme
        "angle": tilt,
        "aspect": azimuth,
        "outputformat": "json",
    }
    resp = requests.get(url, params=params, timeout=60)
    resp.raise_for_status()
    hourly = resp.json()["outputs"]["hourly"]
    s = pd.Series(
        {pd.Timestamp(h["time"], format="%Y%m%d:%H%M", utc=True): h["P"] / 1000.0
         for h in hourly}
    ).sort_index()
    s.index = s.index.tz_convert(CET_FIXED)
    return s.reindex(index, method="ffill")


def bake_formula_cells(ws_write, ws_values) -> int:
    """Ersetzt in `ws_write` (Workbook mit data_only=False, das gespeichert
    wird) jede FORMEL-Zelle durch ihren gecachten Wert aus `ws_values`
    (dasselbe Sheet, aber aus einem mit data_only=True geladenen Workbook).

    HINTERGRUND (echter Bug, gefunden ueber Beats Crash "TypeError: bad
    operand type for unary -: 'NoneType'" bei Parameter!C26): openpyxl
    verliert beim Speichern eines mit data_only=False geladenen Workbooks
    den gecachten Ergebniswert JEDER Formel-Zelle -- es wird nur die
    Formel-Zeichenkette geschrieben, kein Ergebnis. Oeffnet man die
    gespeicherte Datei danach mit data_only=True (wie battery_optimization.py
    es fuer Parameter!C* tut), liefert eine solche Zelle `None`, GARANTIERT,
    unabhaengig vom Inhalt der Formel. Bisher war das nur fuer die
    Zeitreihen-Spalten A/G/H bekannt und dort einzeln abgefangen (siehe
    Kommentar oben bei Spalte A) -- Beats neue Formel in Parameter!C26
    (`=C20*C25*0.01`, Unterhaltkosten aus Kapitalkosten) sowie die
    vorbestehende Formel in Parameter!C29 (`=C20*C28/C27`,
    Durchsatzbegrenzung/Jahr) waren davon NICHT erfasst.

    Fix: GENERISCH fuers ganze Sheet -- jede Zelle mit `data_type == "f"`
    (Formel) wird durch ihren gecachten Wert aus `ws_values` ersetzt, egal
    welche Formel/Zelle das betrifft. Das schuetzt auch vor zukuenftigen
    Formeln, die Beat im Parameter-Sheet ergaenzt. Gibt die Anzahl
    ersetzter Zellen zurueck (fuer eine Konsolen-Meldung).
    """
    n_baked = 0
    for row in ws_write.iter_rows():
        for cell in row:
            if cell.data_type == "f":
                cached = ws_values[cell.coordinate].value
                cell.value = cached
                n_baked += 1
    return n_baked


# --------------------------------------------------------------------------
# 4. Parameter einlesen
# --------------------------------------------------------------------------

def _check_label(ws_param, row: int, expected) -> None:
    """Prueft, dass Parameter!A<row> wie erwartet lautet (Gross-/
    Kleinschreibung und Leerzeichen egal). Schuetzt vor STILLEN
    Fehlinterpretationen, falls Beat das Parameter-Sheet wieder umbaut
    (Zeilen verschoben/eingefuegt -- ist in diesem Projekt schon mehrfach
    passiert) -- lieber ein klarer Fehler hier als falsche Zahlen in der
    Optimierung."""
    actual = ws_param[f"A{row}"].value
    actual_norm = str(actual).strip().lower() if actual is not None else ""
    erwartet = [expected] if isinstance(expected, str) else list(expected)
    if actual_norm not in [e.strip().lower() for e in erwartet]:
        raise RuntimeError(
            f"Parameter!A{row} = {actual!r}, erwartet {erwartet!r}. Das "
            "Parameter-Sheet-Layout hat sich vermutlich veraendert (Zeilen "
            "verschoben/eingefuegt) -- bitte die Zeilennummern in "
            "input_prep.py::read_params() (und battery_optimization.py::"
            "read_inputs()) an das neue Layout anpassen."
        )


def read_params(ws_param) -> dict:
    _check_label(ws_param, 2, "Beginn Analysezeitraum")
    _check_label(ws_param, 3, "Ende Analysezeitraum")
    _check_label(ws_param, 5, "Tarifschema")
    _check_label(ws_param, 6, "HT")
    _check_label(ws_param, 7, "NT")
    _check_label(ws_param, 8, "Netznutzung Leistung")
    _check_label(ws_param, 9, "Tage HT")
    _check_label(ws_param, 10, "Startzeit HT")
    _check_label(ws_param, 11, "Endzeit HT")
    _check_label(ws_param, 12, "Spot_Abschlag_Lieferung")
    _check_label(ws_param, 13, "Spot_Aufschlag_Bezug")
    _check_label(ws_param, 14, "Rücklieferung_PV")
    _check_label(ws_param, 15, "HKN")
    _check_label(ws_param, 16, "Floor")
    _check_label(ws_param, 17, "Fixtarif")
    _check_label(ws_param, 19, "Rücklieferung_Batterie")
    _check_label(ws_param, 20, "HKN")
    _check_label(ws_param, 21, "Floor")
    _check_label(ws_param, 22, "Fixtarif")
    _check_label(ws_param, 25, "DC Leistung")
    _check_label(ws_param, 26, "Produktions-Input")

    return {
        "beginn": ws_param["C2"].value,
        "ende": ws_param["C3"].value,
        "tarifschema": ws_param["C5"].value,
        "ht_raw": ws_param["C6"].value,
        "nt_raw": ws_param["C7"].value,
        "netznutzung_leistung": ws_param["C8"].value,
        "tage_ht": int(ws_param["C9"].value),
        "start_ht_raw": ws_param["C10"].value,
        "end_ht_raw": ws_param["C11"].value,
        "spot_abschlag_lieferung": ws_param["C12"].value,
        "spot_aufschlag_bezug": ws_param["C13"].value,
        "rueckliefer_pv_schema": ws_param["C14"].value,
        "rueckliefer_pv_hkn_raw": ws_param["C15"].value,
        "rueckliefer_pv_floor": ws_param["C16"].value,
        "rueckliefer_pv_fixtarif": ws_param["C17"].value,
        "rueckliefer_batterie_schema": ws_param["C19"].value,
        "rueckliefer_batterie_hkn_raw": ws_param["C20"].value,
        "rueckliefer_batterie_floor": ws_param["C21"].value,
        "rueckliefer_batterie_fixtarif": ws_param["C22"].value,
        "dc_leistung": ws_param["C25"].value,
        "produktions_input": ws_param["C26"].value,
    }


# --------------------------------------------------------------------------
# 5. Hauptablauf
# --------------------------------------------------------------------------

def main(
    input_path: str = INPUT_PATH,
    output_path: str = OUTPUT_PATH,
    fetch_swissix: bool = True,
    fetch_pv: bool = False,
    pv_lat: float | None = None,
    pv_lon: float | None = None,
    swissix_csv_path: str | None = None,
    rmp_path: str | None = None,
):
    print(f"Lese {input_path} ...")
    wb_values = load_workbook(input_path, data_only=True)   # fuer gecachte Werte (G/H)
    wb = load_workbook(input_path, data_only=False)          # zum Beschreiben/Speichern

    ws_param = wb["Parameter"]
    ws_param_vals = wb_values["Parameter"]
    ws_zeit = wb["Zeitreihen"]
    ws_zeit_vals = wb_values["Zeitreihen"]

    # WICHTIG (Bugfix, siehe bake_formula_cells()): jede Formel-Zelle im
    # Parameter-Sheet (z.B. Beats neue Unterhaltkosten-Formel C26 oder die
    # Durchsatzbegrenzung C29) wird HIER in einen festen, gecachten Wert
    # umgewandelt -- sonst geht der Wert beim Speichern von input_lp.xlsx
    # verloren und battery_optimization.py liest spaeter `None` (Absturz).
    n_baked_param = bake_formula_cells(ws_param, ws_param_vals)
    if n_baked_param:
        print(
            f"  {n_baked_param} Formel-Zelle(n) im Sheet 'Parameter' als "
            "feste Werte uebernommen (verhindert Wertverlust beim Speichern)."
        )
    if "Hilfen_Listen" in wb.sheetnames:
        n_baked_hilfen = bake_formula_cells(wb["Hilfen_Listen"], wb_values["Hilfen_Listen"])
        if n_baked_hilfen:
            print(f"  {n_baked_hilfen} Formel-Zelle(n) im Sheet 'Hilfen_Listen' als feste Werte uebernommen.")

    # NEU (Bugfix, gefunden ueber Beats Absturz mit "Inputs_BAT_PV_ohneNetzladen.xlsx"):
    # dieselbe Formel-Wertverlust-Falle wie bei Parameter/Hilfen_Listen oben
    # betraf bisher NUR Zeitreihen!G/H (SRL-Preise, siehe Schritt 4 weiter
    # unten, dort einzeln abgefangen) -- nicht aber andere Zeitreihen-Spalten.
    # Beat hatte Zeitreihen!E (CH_PV_normiert_1kWp) neu als FORMEL befuellt
    # (statt fester Werte) -- ohne generisches Baken ging deren gecachter Wert
    # beim Speichern verloren, battery_optimization.py las anschliessend
    # `None` und stuerzte mit "Zeitreihen!E ist leer" ab, OBWOHL die Zelle im
    # Original-Excel sichtbar einen Wert zeigte. Fix: das GESAMTE
    # Zeitreihen-Sheet generisch baken (wie Parameter/Hilfen_Listen oben) --
    # deckt damit nicht nur Spalte E ab, sondern automatisch auch jede
    # zukuenftige Formel, die Beat in irgendeiner Zeitreihen-Spalte ergaenzt.
    # Spalte A (Zeit) wird weiter unten ohnehin komplett durch ein frisches,
    # driftfreies Zeitraster ersetzt -- das Baken hier stoert dort nicht.
    n_baked_zeit = bake_formula_cells(ws_zeit, ws_zeit_vals)
    if n_baked_zeit:
        print(f"  {n_baked_zeit} Formel-Zelle(n) im Sheet 'Zeitreihen' als feste Werte uebernommen.")

    params = read_params(ws_param)
    n_rows = ws_zeit.max_row

    zeit_raw = [ws_zeit_vals.cell(row=r, column=COL_ZEIT).value for r in range(2, n_rows + 1)]

    # NEU (gefunden ueber Beats "Inputs_BKW_Q2.xlsx"-Absturz beim SwissIX-
    # Fetch): manche Szenario-Dateien fuellen bewusst nur einen TEIL des
    # 35040-Zeilen-Jahres-Templates mit echten Zeitstempeln (z.B. nur Q2 2026,
    # Spalte A danach leer) -- andere Spalten (Last, SRL-Preise) koennen
    # trotzdem ueber das GESAMTE Template hinweg Werte haben (z.B. weil sie in
    # der Vorlage per Formel befuellt sind, unabhaengig vom Szenario-Zeitraum).
    # Bisher wurde IMMER die volle Blattlaenge (ws_zeit.max_row) als Anzahl
    # Zeitschritte angenommen -- bei so einer Teil-Jahr-Datei extrapolierte
    # der driftfreie Zeitraster (siehe Bugfix gleich unten) dadurch weit ueber
    # die tatsaechlich vorhandenen Daten hinaus (hier: ein volles Jahr ab dem
    # 1.4.2026, also bis weit ins naechste Jahr) -- und der anschliessende
    # SwissIX/ENTSO-E-Fetch versuchte dann Datumsbereiche abzufragen, die noch
    # gar nicht existieren ("No matching data found", da ENTSO-E nur
    # Vergangenheit + max. 1 Tag voraus liefert). Fix: nur die FUEHRENDEN,
    # LUECKENLOS befuellten Zeit-Zellen zaehlen -- die erste leere Zeit-Zelle
    # markiert das Ende der tatsaechlichen Szenario-Laenge. Bei vollstaendig
    # befuellten (Voll-Jahr-)Dateien aendert das nichts (n_valid == n_rows-1).
    n_valid = next((i for i, v in enumerate(zeit_raw) if v is None), len(zeit_raw))
    if n_valid < len(zeit_raw):
        print(
            f"  HINWEIS: Zeitreihen!A ist nur fuer {n_valid} von {len(zeit_raw)} "
            f"Zeilen befuellt (erste leere Zeit-Zelle in Zeile {n_valid + 2}) -- "
            "Szenario wird auf diese Laenge begrenzt. Andere Spalten (Last, "
            "SRL-Preise etc.) koennen ueber diesen Zeitraum hinaus noch Werte "
            "enthalten -- diese werden dann NICHT verwendet. Falls ein "
            "vollstaendiges Jahr gewuenscht ist, bitte Spalte A (Zeit) "
            "vollstaendig befuellen."
        )
        zeit_raw = zeit_raw[:n_valid]
        n_rows = n_valid + 1  # +1, da range(2, n_rows+1) ab Zeile 2 zaehlt

    # WICHTIGER BUGFIX (gefunden ueber Beats Meldung "Q2/Q3/Q4 fehlen im
    # Netzbezug-Tagesprofil-Chart", verifiziert an seiner echten Ergebnis-
    # Datei): Excel speichert Datum/Zeit intern als Fliesskomma-Seriennummer
    # (Tage seit 1900). Werden 35040 Viertelstunden-Zeitpunkte per Formel
    # hochgezaehlt, akkumuliert sich winzige Fliesskomma-Rundung -- am
    # Jahresende liegt der gecachte Zeitstempel schon ein paar Millisekunden
    # VOR dem eigentlichen 15-Minuten-Takt (z.B. "23:59:59.993" statt
    # "00:00:00.000", verifiziert: Drift waechst monoton von 0ms am 1.1. auf
    # ~7ms am 31.12.). Fuer die Chronologie/Sortierung ist das folgenlos,
    # ABER `strftime("%H:%M")` (siehe
    # output_create._build_quarterly_avg_profile_generic()) SCHNEIDET die
    # Sekunden ab, statt zu runden -- ein paar Millisekunden zu frueh faellt
    # dadurch in die FALSCHE Minute. Ab ca. Ende Februar (wenn die Drift zum
    # ersten Mal eine volle Millisekunde unter :00 faellt) kippen dadurch die
    # betroffenen 96 Viertelstunden-Bins jedes Tages komplett auf
    # "HH:(MM-1):59.99x" -- die eigentlichen "HH:MM:00"-Bins bleiben fuer den
    # gesamten Rest des Jahres LEER (0 von N Werten), was Beats Symptom exakt
    # erklaert (Q1 grösstenteils ok, Q2/Q3/Q4 komplett betroffen).
    #
    # Fix: statt der (potenziell driftenden) gecachten Excel-Werte einen
    # GARANTIERT exakten 15-Minuten-Raster ab dem ersten Zeitstempel neu
    # erzeugen -- eliminiert die Drift vollstaendig, unabhaengig davon, wie
    # stark sie in der Quelle bereits ist oder in Zukunft waere.
    zeit = pd.date_range(start=zeit_raw[0], periods=len(zeit_raw), freq="15min")
    index = pd.DatetimeIndex(zeit).tz_localize(CET_FIXED)

    # Spalte A war im Original teils eine Array-Formel (Anker in A2, Rest als
    # gespiegelte Werte). openpyxl verliert beim Speichern den gecachten Wert
    # der Formel-Ankerzelle -- deshalb hier explizit mit dem oben erzeugten,
    # driftfreien 15-Minuten-Raster ueberschreiben, damit Spalte A garantiert
    # vollstaendig, formelfrei UND exakt im Output steht.
    for r, t in zip(range(2, n_rows + 1), zeit):
        ws_zeit.cell(row=r, column=COL_ZEIT).value = t.to_pydatetime()

    existing_bezug = pd.Series(
        [ws_zeit_vals.cell(row=r, column=COL_BEZUG).value for r in range(2, n_rows + 1)],
        index=index, dtype=float,
    )
    # NEU (Bugfix, siehe unten): bereits vorhandene SwissIX-Werte in Spalte B
    # VOR dem Fetch-Versuch einlesen -- dient als LETZTER manueller Fallback
    # in resolve_swissix() unten (Stufe 3), falls weder die CSV-Datei noch
    # die Website etwas liefern.
    existing_swissix = pd.Series(
        [ws_zeit_vals.cell(row=r, column=COL_SWISSIX).value for r in range(2, n_rows + 1)],
        index=index, dtype=float,
    )

    # --- 1. SwissIX --------------------------------------------------
    # NEU (Beats Wunsch): SwissIX jetzt in dieser Prioritaet ermitteln --
    # (a) lokale CSV-Datei, (b) Website (ENTSO-E) fuer das, was die Datei
    # nicht abdeckt, (c) bereits von Hand in Zeitreihen!B eingetragene Werte
    # fuer das, was danach immer noch fehlt. Siehe resolve_swissix() weiter
    # oben fuer die Details; hier nur der Pfad-Aufbau + der Aufruf.
    if swissix_csv_path is None:
        swissix_csv_path = os.path.join(
            os.path.dirname(os.path.abspath(input_path)), SWISSIX_CSV_SUBDIR, SWISSIX_CSV_FILENAME
        )
    swissix, swissix_quelle = resolve_swissix(index, swissix_csv_path, fetch_swissix, existing_swissix)
    if swissix is not None:
        for r, v in zip(range(2, n_rows + 1), swissix.values):
            if pd.notna(v):
                ws_zeit.cell(row=r, column=COL_SWISSIX).value = float(v)

    # NEU (Beats Wunsch, siehe Hinweis zu Inputs_BAT_standalone.xlsx): im
    # Report soll transparent sein, woher die SwissIX-Preise stammen -- dafuer
    # in einer bisher ungenutzten Zelle (Parameter!E5, neben der Tarifschema-
    # Zeile) vermerkt. battery_optimization.read_inputs() liest diese Zelle
    # als params["swissix_quelle"] und output_create.build_input_summary_rows()
    # zeigt sie auf der Eingabeparameter-Seite an.
    if swissix_quelle is not None:
        ws_param["E5"] = swissix_quelle

    # NEU (kritischer Plausibilitaets-Check, gefunden bei Beats
    # Inputs_BAT_standalone.xlsx-Lauf): die von Hand eingetragenen SwissIX-
    # Werte dort liegen im Bereich ~85-115 -- das ist der TYPISCHE Massstab
    # fuer EUR/MWh (rohe ENTSO-E-Day-Ahead-Preise), NICHT fuer CHF/kWh (wo
    # selbst Preis-Spitzen fast immer < 1-2 liegen). Ungeprueft fuehrt das zu
    # einem ~1000x zu hohen Bezugs-/Rueckliefertarif und (verifiziert an
    # genau diesem Fall) zu einem absurden Arbitrage-Ertrag von > 22 Mio.
    # CHF/Jahr in einem 1 MWh/1 MW-System. Deshalb hier eine harte, nicht zu
    # uebersehende Warnung, unabhaengig davon ob die Werte von der API kamen
    # oder von Hand eingetragen wurden.
    if swissix is not None and swissix.notna().any():
        swissix_median_abs = float(swissix.abs().median())
        if swissix_median_abs > 3.0:
            warnings.warn(
                f"SwissIX-Werte (Zeitreihen!B) haben einen unplausibel hohen "
                f"Median von {swissix_median_abs:,.1f} fuer eine Einheit "
                "CHF/kWh -- typische Strompreise liegen auch bei Spitzen "
                "fast immer unter 1-2 CHF/kWh. Das sieht nach EUR/MWh bzw. "
                "CHF/MWh statt CHF/kWh aus (Faktor ~1000 zu hoch) -- bitte "
                "pruefen, ob die Werte in Zeitreihen!B durch 1000 geteilt "
                "werden muessen. Mit den aktuellen Werten wird der SPOT-"
                "Bezugs-/Rueckliefertarif (und damit die gesamte "
                "Wirtschaftlichkeitsrechnung) um denselben Faktor verzerrt."
            )

    # --- 2. Tarife -----------------------------------------------------
    # NEU (Beats Wunsch: RMP-Rueckliefertarif fuer PV/Batterie): Solar_RMP.xlsx
    # nur laden, wenn sie tatsaechlich am erwarteten Pfad liegt -- ist sie
    # nicht noetig (kein RMP/RMP_Floor-Schema gewaehlt), soll ihr Fehlen kein
    # Fehler sein. Wird sie GEBRAUCHT aber fehlt, wirft resolve_rueckliefertarif()
    # weiter unten einen klaren Fehler (rmp_lookup=None).
    if rmp_path is None:
        rmp_path = os.path.join(
            os.path.dirname(os.path.abspath(input_path)), SOLAR_RMP_SUBDIR, SOLAR_RMP_FILENAME
        )
    if os.path.isfile(rmp_path):
        print(f"Lade RMP-Quartalswerte aus {rmp_path} ...")
        rmp_lookup = load_rmp_lookup(rmp_path)
    else:
        print(f"  Solar_RMP.xlsx nicht gefunden unter {rmp_path} -- wird nur gebraucht, falls RMP/RMP_Floor gewaehlt ist.")
        rmp_lookup = None

    print("Berechne Rueckliefer-/Bezugstarif ...")
    ruecklieferung_pv, ruecklieferung_batterie, bezug = build_tariffs(
        index, params, swissix, existing_bezug, rmp_lookup
    )
    for r, v in zip(range(2, n_rows + 1), ruecklieferung_pv.values):
        if pd.notna(v):
            ws_zeit.cell(row=r, column=COL_RUECKLIEFER).value = float(v)
    for r, v in zip(range(2, n_rows + 1), bezug.values):
        if pd.notna(v):
            ws_zeit.cell(row=r, column=COL_BEZUG).value = float(v)
    # NEU: eigene Spalte K fuer den Batterie-Rueckliefertarif (siehe
    # COL_RUECKLIEFER_BATTERIE weiter oben) -- Header wird hier automatisch
    # gesetzt, falls noch nicht vorhanden (Beats Original-Excel hat diese
    # Spalte noch nicht).
    if ws_zeit.cell(row=1, column=COL_RUECKLIEFER_BATTERIE).value != COL_RUECKLIEFER_BATTERIE_HEADER:
        ws_zeit.cell(row=1, column=COL_RUECKLIEFER_BATTERIE).value = COL_RUECKLIEFER_BATTERIE_HEADER
    for r, v in zip(range(2, n_rows + 1), ruecklieferung_batterie.values):
        if pd.notna(v):
            ws_zeit.cell(row=r, column=COL_RUECKLIEFER_BATTERIE).value = float(v)

    # --- 3. PV-Profil (optional) ----------------------------------------
    if fetch_pv and params["produktions_input"] == "normiert":
        if pv_lat is None or pv_lon is None:
            warnings.warn(
                "fetch_pv=True aber pv_lat/pv_lon nicht angegeben -- "
                "PV-Profil (Spalte E) bleibt leer."
            )
        else:
            try:
                print("Lade PV-Referenzprofil von PVGIS ...")
                pv_profile = fetch_pv_reference_profile(index, pv_lat, pv_lon)
                for r, v in zip(range(2, n_rows + 1), pv_profile.values):
                    ws_zeit.cell(row=r, column=COL_PV_NORMIERT).value = float(v)
                print("  -> PV-Profil erfolgreich ergaenzt.")
            except Exception as exc:  # noqa: BLE001
                warnings.warn(f"PVGIS-Fetch fehlgeschlagen ({exc}).")
    else:
        print(
            "PV-Referenzprofil (Spalte E) NICHT ergaenzt -- Standort/Ausrichtung "
            "noch nicht bekannt (siehe Excel-Kommentar 'mail Urs/Astrid'). "
            "main(fetch_pv=True, pv_lat=..., pv_lon=...) aufrufen, sobald bekannt."
        )

    # --- 4. SRL-Preise "backen" (bereits im Original per Formel berechnet) --
    print("Uebernehme SRL-Preise (bereits vorhanden) als feste Werte ...")
    for r in range(2, n_rows + 1):
        for col in (COL_SRL_NEG, COL_SRL_POS):
            cached = ws_zeit_vals.cell(row=r, column=col).value
            if cached is not None:
                ws_zeit.cell(row=r, column=col).value = cached

    # --- 5. Speichern ----------------------------------------------------
    print(f"Speichere {output_path} ...")
    wb.save(output_path)
    print("Fertig.")


if __name__ == "__main__":
    main()