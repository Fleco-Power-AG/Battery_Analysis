# -*- coding: utf-8 -*-
"""
pv_standort.py
==============
Berechnet ein PV-Erzeugungsprofil (kW) für einen konkreten Standort aus GEMESSENEN
Strahlungs- und Temperaturwerten von MeteoSwiss (SwissMetNet, Open Government Data).

Ablauf
------
1. PLZ -> Koordinaten (swisstopo-Suchdienst). Alternativ direkt "Breite, Länge".
2. Nächstgelegene SwissMetNet-Station mit Globalstrahlung (gre000z0).
3. 10-Minuten-Messwerte (Globalstrahlung, Lufttemperatur) für den Analysezeitraum
   laden und in "03 Stammdaten/MeteoSwiss" zwischenspeichern.
4. Pro Teilfläche (bis 4, je kWp / Neigung / Ausrichtung), mit pvlib:
   Sonnenstand (NREL SPA) -> Aufteilung Direkt/Diffus (Erbs) -> Strahlung auf die
   Modulebene (Perez) -> Winkelverluste (ASHRAE) -> Modultemperatur (Faiman, mit
   gemessenem Wind) -> DC-Leistung (PVWatts) -> Systemverluste -> Wechselrichter
   (PVWatts-Wirkungsgradkurve, Begrenzung auf die AC-Leistung).
   Ist pvlib nicht installiert, rechnet eine vereinfachte numpy-Variante
   (NOAA, Erbs, Hay-Davies) mit Warnung weiter.
5. Summe aller Flächen, optional auf die Wechselrichter-AC-Leistung begrenzt,
   als 15-Minuten-Mittelwerte auf das Zeitraster der Analyse.

Abhängigkeiten: numpy, pandas, pvlib (conda-forge).

Konventionen
------------
Ausrichtung (Azimut der Modulfläche): Süd = 0°, Ost = -90°, West = +90°, Nord = 180°.
Neigung: 0° = horizontal, 90° = senkrecht.
MeteoSwiss-Zeitstempel sind UTC und bezeichnen das ENDE des 10-Minuten-Intervalls.
"""

from __future__ import annotations

import io
import json
import math
import os
import time
import urllib.parse
import urllib.request
from dataclasses import dataclass

import numpy as np
import pandas as pd

STAC_COLLECTION = "https://data.geo.admin.ch/api/stac/v1/collections/ch.meteoschweiz.ogd-smn"
DATA_BASE = "https://data.geo.admin.ch/ch.meteoschweiz.ogd-smn"
GEO_SEARCH = "https://api3.geo.admin.ch/rest/services/api/SearchServer"
PARAM_GHI = "gre000z0"      # Globalstrahlung, 10-Min-Mittel [W/m²]
PARAM_TEMP = "tre200s0"     # Lufttemperatur 2 m [°C]
PARAM_WIND = "fkl010z0"     # Windgeschwindigkeit skalar, 10-Min-Mittel [m/s]

STANDARD_VERLUSTE = 0.14    # Kabel, Verschmutzung, Mismatch, Toleranzen (PVWatts-Standard)
WECHSELRICHTER_ETA = 0.965
TEMP_KOEFF = -0.004         # 1/K, kristalline Module
ALBEDO = 0.2
TIMEOUT = 60


@dataclass
class Flaeche:
    kwp: float
    neigung: float
    ausrichtung: float      # Süd = 0, Ost = -90, West = +90


# =============================================================================
# Download-Helfer
# =============================================================================
def _get(url: str) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": "FlecoPower-Batterieanalyse"})
    with urllib.request.urlopen(req, timeout=TIMEOUT) as r:
        return r.read()


def _csv_lesen(raw: bytes, **kw) -> pd.DataFrame:
    for enc in ("utf-8", "cp1252", "latin-1"):
        try:
            return pd.read_csv(io.BytesIO(raw), sep=";", encoding=enc, **kw)
        except UnicodeDecodeError:
            continue
    raise ValueError("CSV konnte nicht gelesen werden.")


def _cache(cache_dir: str, name: str, url_fn, max_alter_h: float | None) -> bytes:
    """Datei aus dem Cache lesen oder neu laden (max_alter_h=None: nie neu laden)."""
    os.makedirs(cache_dir, exist_ok=True)
    pfad = os.path.join(cache_dir, name)
    if os.path.isfile(pfad):
        alter_h = (time.time() - os.path.getmtime(pfad)) / 3600
        if max_alter_h is None or alter_h <= max_alter_h:
            with open(pfad, "rb") as f:
                return f.read()
    try:
        raw = _get(url_fn())
    except Exception as exc:
        if os.path.isfile(pfad):
            print(f"  WARNUNG: {name} konnte nicht aktualisiert werden ({exc}), verwende Cache.")
            with open(pfad, "rb") as f:
                return f.read()
        raise
    with open(pfad, "wb") as f:
        f.write(raw)
    return raw


def _stac_assets(url: str) -> dict:
    try:
        return {k: v.get("href") for k, v in json.loads(_get(url)).get("assets", {}).items()}
    except Exception:
        return {}


# =============================================================================
# 1) Standort
# =============================================================================
def standort_koordinaten(eingabe) -> tuple[float, float, str]:
    """PLZ (z.B. 8556) oder 'Breite, Länge' -> (lat, lon, Bezeichnung)."""
    txt = str(eingabe).strip()
    if "," in txt or ";" in txt:
        teile = [t.strip() for t in txt.replace(";", ",").split(",")]
        if len(teile) == 2:
            lat, lon = float(teile[0]), float(teile[1])
            return lat, lon, f"{lat:.4f}, {lon:.4f}"
    plz = txt.split(".")[0]          # Excel liefert ggf. 8556.0
    q = urllib.parse.urlencode({"searchText": plz, "type": "locations", "sr": "4326", "limit": 20})
    daten = json.loads(_get(f"{GEO_SEARCH}?{q}"))
    treffer = daten.get("results", [])
    plz_treffer = [t for t in treffer if t.get("attrs", {}).get("origin") == "zipcode"] or treffer
    if not plz_treffer:
        raise ValueError(f"PLZ {plz} wurde nicht gefunden.")
    a = plz_treffer[0]["attrs"]
    label = str(a.get("label", plz)).replace("<b>", "").replace("</b>", "")
    return float(a["lat"]), float(a["lon"]), label


# =============================================================================
# 2) Station
# =============================================================================
def _haversine_km(lat1, lon1, lat2, lon2):
    r = 6371.0
    p1, p2 = np.radians(lat1), np.radians(lat2)
    dp, dl = p2 - p1, np.radians(lon2 - lon1)
    a = np.sin(dp / 2) ** 2 + np.cos(p1) * np.cos(p2) * np.sin(dl / 2) ** 2
    return 2 * r * np.arcsin(np.sqrt(a))


def _meta(cache_dir: str, wort: str) -> pd.DataFrame:
    name = f"ogd-smn_meta_{wort}.csv"

    def url():
        assets = _stac_assets(STAC_COLLECTION)
        for k, href in assets.items():
            if wort in k:
                return href
        return f"{DATA_BASE}/{name}"
    return _csv_lesen(_cache(cache_dir, name, url, max_alter_h=24 * 30))


def stationen_mit_strahlung(cache_dir: str) -> pd.DataFrame:
    st = _meta(cache_dir, "stations")
    st.columns = [c.strip().lower() for c in st.columns]
    try:
        inv = _meta(cache_dir, "datainventory")
        inv.columns = [c.strip().lower() for c in inv.columns]
        mit = inv[inv["parameter_shortname"].astype(str).str.strip() == PARAM_GHI]["station_abbr"]
        st = st[st["station_abbr"].isin(set(mit.astype(str).str.strip()))]
    except Exception:
        pass   # ohne Inventar: Strahlung wird beim Laden der Daten geprüft
    return st


def naechste_stationen(lat: float, lon: float, cache_dir: str, n: int = 5) -> pd.DataFrame:
    st = stationen_mit_strahlung(cache_dir).copy()
    st["dist_km"] = _haversine_km(lat, lon, st["station_coordinates_wgs84_lat"].astype(float),
                                  st["station_coordinates_wgs84_lon"].astype(float))
    return st.sort_values("dist_km").head(n)


# =============================================================================
# 3) Messdaten
# =============================================================================
def _dateien_fuer_zeitraum(start_utc: pd.Timestamp, ende_utc: pd.Timestamp) -> list[str]:
    jahr_jetzt = pd.Timestamp.utcnow().year
    teile = []
    for dekade in range(start_utc.year // 10 * 10, ende_utc.year // 10 * 10 + 1, 10):
        teile.append(f"historical_{dekade}-{dekade + 9}")
    if ende_utc.year >= jahr_jetzt:
        teile.append("recent")
    return teile


def messdaten(abbr: str, start_utc: pd.Timestamp, ende_utc: pd.Timestamp, cache_dir: str) -> pd.DataFrame:
    """10-Min-Messwerte (ghi W/m², temp °C, wind m/s), Index = Intervall-ENDE in UTC."""
    a = abbr.lower()
    teile_df = []
    assets = None
    for teil in _dateien_fuer_zeitraum(start_utc, ende_utc):
        name = f"ogd-smn_{a}_t_{teil}.csv"

        def url(name=name):
            nonlocal assets
            if assets is None:
                assets = _stac_assets(f"{STAC_COLLECTION}/items/{a}")
            return assets.get(name) or f"{DATA_BASE}/{a}/{name}"
        max_alter = 20 if teil == "recent" else 24 * 60   # recent: täglich, historisch: alle 60 Tage
        try:
            raw = _cache(cache_dir, name, url, max_alter_h=max_alter)
        except Exception as exc:
            print(f"  Hinweis: {name} nicht verfügbar ({exc}).")
            continue
        df = _csv_lesen(raw, usecols=lambda c: c.strip() in ("reference_timestamp", PARAM_GHI, PARAM_TEMP,
                                                             PARAM_WIND), dtype=str)
        df.columns = [c.strip() for c in df.columns]
        if PARAM_GHI not in df.columns:
            raise ValueError(f"Station {abbr.upper()} liefert keine Globalstrahlung ({PARAM_GHI}).")
        df["zeit"] = pd.to_datetime(df["reference_timestamp"], format="%d.%m.%Y %H:%M", utc=True)
        out = pd.DataFrame({
            "ghi": pd.to_numeric(df[PARAM_GHI], errors="coerce").values,
            "temp": pd.to_numeric(df[PARAM_TEMP], errors="coerce").values if PARAM_TEMP in df else np.nan,
            "wind": pd.to_numeric(df[PARAM_WIND], errors="coerce").values if PARAM_WIND in df else np.nan,
        }, index=df["zeit"])
        teile_df.append(out)
    if not teile_df:
        raise RuntimeError(f"Keine Messdaten für Station {abbr.upper()} erhalten.")
    d = pd.concat(teile_df)
    d = d[~d.index.duplicated(keep="last")].sort_index()
    d = d[(d.index > start_utc) & (d.index <= ende_utc + pd.Timedelta(minutes=10))]
    return d


# =============================================================================
# 4) Physik
# =============================================================================
def sonnenstand(zeit_utc: pd.DatetimeIndex, lat: float, lon: float):
    """Zenit- und Azimutwinkel [°] (Azimut: Nord=0, Ost=90, Süd=180) nach NOAA."""
    doy = zeit_utc.dayofyear.values
    std = zeit_utc.hour.values + zeit_utc.minute.values / 60 + zeit_utc.second.values / 3600
    g = 2 * np.pi / 365 * (doy - 1 + (std - 12) / 24)
    eqt = 229.18 * (0.000075 + 0.001868 * np.cos(g) - 0.032077 * np.sin(g)
                    - 0.014615 * np.cos(2 * g) - 0.040849 * np.sin(2 * g))
    dekl = (0.006918 - 0.399912 * np.cos(g) + 0.070257 * np.sin(g) - 0.006758 * np.cos(2 * g)
            + 0.000907 * np.sin(2 * g) - 0.002697 * np.cos(3 * g) + 0.00148 * np.sin(3 * g))
    wahre_zeit = std * 60 + eqt + 4 * lon
    stundenwinkel = np.radians(wahre_zeit / 4 - 180)
    phi = np.radians(lat)
    cos_z = np.sin(phi) * np.sin(dekl) + np.cos(phi) * np.cos(dekl) * np.cos(stundenwinkel)
    cos_z = np.clip(cos_z, -1, 1)
    zen = np.arccos(cos_z)
    az = np.degrees(np.arctan2(np.sin(stundenwinkel),
                               np.cos(stundenwinkel) * np.sin(phi) - np.tan(dekl) * np.cos(phi))) + 180
    return np.degrees(zen), az % 360, doy


def extraterrestrisch(doy):
    return 1367.0 * (1 + 0.033 * np.cos(2 * np.pi * doy / 365))


def erbs(ghi, zen, doy):
    """Globalstrahlung -> Direktnormal (DNI) und Diffus (DHI)."""
    cos_z = np.cos(np.radians(zen))
    i0h = extraterrestrisch(doy) * np.maximum(cos_z, 0.065)
    kt = np.clip(np.where(zen < 87, ghi / i0h, 0), 0, 1)
    df = np.where(kt <= 0.22, 1 - 0.09 * kt,
                  np.where(kt <= 0.8, 0.9511 - 0.1604 * kt + 4.388 * kt ** 2 - 16.638 * kt ** 3 + 12.336 * kt ** 4,
                           0.165))
    dhi = df * ghi
    dni = np.where(zen < 87, (ghi - dhi) / np.maximum(cos_z, 0.065), 0)
    return np.clip(dni, 0, 1100), np.clip(dhi, 0, None)


def poa(ghi, dni, dhi, zen, az_sonne, doy, neigung, ausrichtung):
    """Strahlung auf die Modulebene (Hay-Davies) inkl. Winkelverlusten des Direktanteils."""
    beta = np.radians(neigung)
    az_flaeche = np.radians((ausrichtung + 180) % 360)      # Süd=0 -> 180° (Nord=0-System)
    z, a = np.radians(zen), np.radians(az_sonne)
    cos_aoi = np.cos(z) * np.cos(beta) + np.sin(z) * np.sin(beta) * np.cos(a - az_flaeche)
    cos_aoi_pos = np.clip(cos_aoi, 0, 1)
    sonne_oben = zen < 90
    rb = np.where(sonne_oben, cos_aoi_pos / np.maximum(np.cos(z), 0.087), 0)
    ai = np.clip(dni / extraterrestrisch(doy), 0, 1)
    diffus = dhi * (ai * rb + (1 - ai) * (1 + np.cos(beta)) / 2)
    boden = ghi * ALBEDO * (1 - np.cos(beta)) / 2
    iam = np.clip(1 - 0.05 * (1 / np.maximum(cos_aoi_pos, 1e-3) - 1), 0, 1)   # ASHRAE, b0 = 0.05
    direkt = np.where(sonne_oben, dni * cos_aoi_pos * iam, 0)
    return np.clip(direkt + diffus + boden, 0, None)


def pvlib_verfuegbar() -> bool:
    try:
        import pvlib  # noqa: F401
        return True
    except ImportError:
        return False


def pv_leistung_kw(mess: pd.DataFrame, lat: float, lon: float, flaechen: list[Flaeche],
                   verluste: float = STANDARD_VERLUSTE, ac_max_kw: float | None = None,
                   hoehe_m: float | None = None) -> pd.Series:
    """AC-Leistung [kW] je 10-Min-Messintervall (Index = Intervall-Ende, UTC).
    Rechnet mit pvlib, sonst mit der vereinfachten numpy-Variante."""
    if pvlib_verfuegbar():
        return pv_leistung_kw_pvlib(mess, lat, lon, flaechen, verluste, ac_max_kw, hoehe_m)
    print("  WARNUNG: pvlib nicht installiert -> vereinfachte Berechnung (Hay-Davies statt Perez). "
          "Installation: conda install -c conda-forge pvlib")
    return pv_leistung_kw_numpy(mess, lat, lon, flaechen, verluste, ac_max_kw)


def pv_leistung_kw_pvlib(mess, lat, lon, flaechen, verluste, ac_max_kw, hoehe_m):
    import pvlib

    t = mess.index - pd.Timedelta(minutes=5)                 # Intervallmitte, UTC
    ghi = pd.Series(np.clip(mess["ghi"].values, 0, None), index=t)
    temp = pd.Series(mess["temp"].values, index=t).fillna(10.0)
    wind = (pd.Series(mess["wind"].values, index=t) if "wind" in mess else pd.Series(np.nan, index=t)).fillna(1.0)

    sp = pvlib.solarposition.get_solarposition(t, lat, lon, altitude=hoehe_m)
    zerl = pvlib.irradiance.erbs(ghi, sp["zenith"], t)
    dni, dhi = zerl["dni"].fillna(0), zerl["dhi"].fillna(0)
    dni_extra = pvlib.irradiance.get_extra_radiation(t)
    airmass = pvlib.atmosphere.get_relative_airmass(sp["apparent_zenith"])

    p_dc = pd.Series(0.0, index=t)
    for f in flaechen:
        if not f.kwp:
            continue
        az = (f.ausrichtung + 180) % 360                       # Süd=0 -> pvlib Süd=180
        poa = pvlib.irradiance.get_total_irradiance(
            f.neigung, az, sp["apparent_zenith"], sp["azimuth"], dni, ghi, dhi,
            dni_extra=dni_extra, airmass=airmass, albedo=ALBEDO, model="perez")
        aoi = pvlib.irradiance.aoi(f.neigung, az, sp["apparent_zenith"], sp["azimuth"])
        iam = pvlib.iam.ashrae(aoi, b=0.05)
        eff = (poa["poa_direct"].fillna(0) * iam.fillna(0) + poa["poa_diffuse"].fillna(0)).clip(lower=0)
        t_zelle = pvlib.temperature.faiman(poa["poa_global"].fillna(0), temp, wind)
        p_dc += pvlib.pvsystem.pvwatts_dc(eff, t_zelle, f.kwp, TEMP_KOEFF).clip(lower=0).fillna(0)

    p_dc *= (1 - verluste)
    kwp = sum(f.kwp for f in flaechen if f.kwp)
    # Wechselrichter (PVWatts): pdc0 = DC-Eingangsgrenze -> AC-Maximum = eta_nom * pdc0
    pdc0 = (ac_max_kw / WECHSELRICHTER_ETA) if ac_max_kw else kwp * 2.0
    p_ac = pvlib.inverter.pvwatts(p_dc, pdc0, eta_inv_nom=WECHSELRICHTER_ETA).clip(lower=0).fillna(0)
    return pd.Series(p_ac.values, index=mess.index)


def pv_leistung_kw_numpy(mess: pd.DataFrame, lat: float, lon: float, flaechen: list[Flaeche],
                         verluste: float = STANDARD_VERLUSTE, ac_max_kw: float | None = None) -> pd.Series:
    """Vereinfachte Variante ohne pvlib (Notfall)."""
    t_mitte = mess.index - pd.Timedelta(minutes=5)          # Sonnenstand in Intervallmitte
    zen, az, doy = sonnenstand(t_mitte, lat, lon)
    ghi = np.clip(mess["ghi"].values, 0, None)
    temp = mess["temp"].values if "temp" in mess else np.full(len(mess), 10.0)
    temp = np.where(np.isnan(temp), 10.0, temp)
    dni, dhi = erbs(ghi, zen, doy)
    gesamt = np.zeros(len(mess))
    for f in flaechen:
        if not f.kwp:
            continue
        e = poa(ghi, dni, dhi, zen, az, doy, f.neigung, f.ausrichtung)
        t_zelle = temp + e * (45 - 20) / 800                  # NOCT 45 °C
        p_dc = f.kwp * e / 1000 * (1 + TEMP_KOEFF * (t_zelle - 25))
        gesamt += np.clip(p_dc, 0, None)
    p_ac = gesamt * (1 - verluste) * WECHSELRICHTER_ETA
    if ac_max_kw:
        p_ac = np.minimum(p_ac, ac_max_kw)
    return pd.Series(p_ac, index=mess.index)


def auf_raster(p10: pd.Series, index: pd.DatetimeIndex) -> pd.Series:
    """10-Min-Werte (Index = Intervall-Ende) -> Mittelwerte auf das 15-Min-Raster der Analyse
    (Index-Zeitstempel = Intervall-Beginn)."""
    # jeder 10-Min-Wert gilt für die beiden 5-Min-Abschnitte davor
    s5 = p10.copy()
    s5.index = s5.index - pd.Timedelta(minutes=10)            # -> Intervall-Beginn
    s5 = s5.resample("5min").ffill(limit=1)
    ziel = index.tz_convert("UTC")
    p15 = s5.resample("15min").mean()
    return p15.reindex(ziel).values


# =============================================================================
# 5) Gesamtablauf
# =============================================================================
def berechne_standortprofil(standort, flaechen: list[Flaeche], index: pd.DatetimeIndex, cache_dir: str,
                            verluste: float | None = None, ac_max_kw: float | None = None):
    """Liefert (kW-Zeitreihe auf `index`, Info-Dict für Report/Konsole)."""
    verluste = STANDARD_VERLUSTE if verluste in (None, "") else float(verluste)
    lat, lon, label = standort_koordinaten(standort)
    print(f"  PV-Standort: {label} ({lat:.4f}, {lon:.4f})")
    start = index[0].tz_convert("UTC")
    ende = (index[-1] + pd.Timedelta(minutes=15)).tz_convert("UTC")

    letzter_fehler = None
    for _, st in naechste_stationen(lat, lon, cache_dir).iterrows():
        abbr = str(st["station_abbr"]).strip()
        try:
            mess = messdaten(abbr, start, ende, cache_dir)
        except Exception as exc:
            letzter_fehler = exc
            print(f"  Station {abbr} übersprungen: {exc}")
            continue
        soll = int((ende - start) / pd.Timedelta(minutes=10))
        abdeckung = mess["ghi"].notna().sum() / max(soll, 1)
        if abdeckung < 0.9:
            letzter_fehler = RuntimeError(f"Station {abbr}: nur {abdeckung:.0%} des Zeitraums mit Strahlungsdaten")
            print(f"  Station {abbr} übersprungen: nur {abdeckung:.0%} Datenabdeckung.")
            continue
        # kleine Lücken (Messausfälle) linear füllen, längere als 0 (konservativ)
        voll = pd.date_range(start + pd.Timedelta(minutes=10), ende, freq="10min")
        mess = mess.reindex(voll)
        mess["ghi"] = mess["ghi"].interpolate(limit=6).fillna(0)
        mess["temp"] = mess["temp"].interpolate(limit=36)
        if "wind" in mess:
            mess["wind"] = mess["wind"].interpolate(limit=36)
        hoehe = pd.to_numeric(st.get("station_height_masl"), errors="coerce")
        p10 = pv_leistung_kw(mess, lat, lon, flaechen, verluste, ac_max_kw,
                             None if pd.isna(hoehe) else float(hoehe))
        profil = pd.Series(auf_raster(p10, index), index=index).fillna(0.0)
        kwp = sum(f.kwp for f in flaechen if f.kwp)
        jahre = len(index) * 0.25 / 8760
        spez = profil.sum() * 0.25 / kwp / jahre if kwp and jahre else float("nan")
        info = dict(
            station=abbr, station_name=str(st.get("station_name", abbr)), dist_km=float(st["dist_km"]),
            station_hoehe=st.get("station_height_masl"), standort=label, lat=lat, lon=lon, kwp=kwp,
            abdeckung=abdeckung, spez_ertrag=spez, verluste=verluste, ac_max_kw=ac_max_kw,
            modell="pvlib" if pvlib_verfuegbar() else "vereinfacht",
        )
        print(f"  MeteoSwiss-Station {abbr} ({info['station_name']}), {info['dist_km']:.1f} km entfernt, "
              f"Datenabdeckung {abdeckung:.0%}")
        print(f"  Spezifischer Ertrag: {spez:,.0f} kWh/kWp pro Jahr (hochgerechnet)")
        return profil, info
    raise RuntimeError(f"Keine geeignete MeteoSwiss-Station gefunden ({letzter_fehler}).")
