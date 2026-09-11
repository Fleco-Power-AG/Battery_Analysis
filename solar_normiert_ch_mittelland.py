"""
PV-Normprofil (auf 1 kWp normiert) fürs Schweizer Mittelland, 15-Minuten-Auflösung, 1 Jahr.

Datenquelle: PVGIS (Photovoltaic Geographical Information System) der EU-Kommission
(Joint Research Centre) - kostenlos, ohne API-Key, öffentlich zugänglich unter
https://re.jrc.ec.europa.eu/api/v5_3/seriescalc

Vorgehen:
1. Für mehrere repräsentative Standorte im Mittelland (Bern, Zürich, Luzern,
   Solothurn, Aarau, Fribourg) wird via PVGIS ein PV-Ertragsprofil für 1 kWp
   installierte Leistung abgerufen (stündliche Auflösung, das ist die feinste
   Auflösung die PVGIS liefert).
2. Es wird über mehrere Jahre (z.B. 2016-2020) und über alle Standorte gemittelt,
   um ein "typisches" Jahr (Durchschnittsprognose) zu erhalten - dabei wird pro
   Kalendertag/Uhrzeit gemittelt (Schaltjahr-Tag 29.02. wird verworfen, damit ein
   sauberes 365-Tage-Jahr entsteht).
3. Das resultierende stündliche Profil wird linear auf 15-Minuten-Werte interpoliert
   (PVGIS liefert nur stündliche Werte - echte 15-Min-Messwerte gibt es dort nicht).
4. Export als CSV mit den Spalten "Zeit" und "CH_PV_normiert_1kWp [kW]" - passend
   zum bestehenden Spaltennamen in eurem Inputs.xlsx (Sheet "Zeitreihen", Spalte E).

WICHTIG:
- Dieses Skript braucht Internetzugriff auf re.jrc.ec.europa.eu. In manchen
  Firmennetzwerken/Sandboxes ist das gesperrt - dann lokal (eigener Laptop) ausführen.
- "Normiert auf 1 kWp" heisst: die Werte sind die PV-Leistung [kW], die eine Anlage
  mit exakt 1 kWp installierter Leistung produzieren würde (Wertebereich 0 bis ~1,
  je nach Systemverlusten meist < 1). Für eine reale Anlage mit z.B. 10 kWp einfach
  mit 10 multiplizieren.

Benötigte Pakete: requests, pandas, numpy
    pip install requests pandas numpy
"""

import time
import numpy as np
import pandas as pd
import requests

PVGIS_URL = "https://re.jrc.ec.europa.eu/api/v5_3/seriescalc"

# Repräsentative Standorte im Schweizer Mittelland (bewusst Flachland, keine Alpen/Voralpen)
STANDORTE = {
    "Bern":      (46.95, 7.44),
    "Zuerich":   (47.37, 8.54),
    "Luzern":    (47.05, 8.31),
    "Solothurn": (47.21, 7.54),
    "Aarau":     (47.39, 8.05),
    "Fribourg":  (46.80, 7.15),
}

# Referenzjahre für die Durchschnittsbildung (mehrere Jahre = robusterer Durchschnitt,
# da Wetter/Strahlung von Jahr zu Jahr schwankt). PVGIS-Datenbank deckt i.d.R. bis
# vor 1-2 Jahre ab; bei Bedarf anpassen.
START_JAHR = 2016
END_JAHR = 2020

# PV-Systemannahmen (typisches Schweizer Aufdach-System)
PEAKPOWER_KWP = 1.0     # normiert auf 1 kWp
SYSTEM_VERLUST_PROZENT = 14.0  # typische Gesamtverluste (Wechselrichter, Kabel, Verschmutzung, Temperatur)
NEIGUNG_GRAD = 30.0     # Dachneigung
AUSRICHTUNG_GRAD = 0.0  # 0 = Süden (PVGIS-Konvention: 0=Süd, -90=Ost, 90=West)

# Referenzjahr für die Ausgabe-Zeitreihe (nicht-Schaltjahr, damit sauber 365 Tage x 96 = 35'040 Zeilen)
OUTPUT_JAHR = 2025


def hole_pvgis_serie(lat, lon, start_jahr, end_jahr, session):
    """Ruft die stündliche PV-Ertragsserie für einen Standort von PVGIS ab.
    Gibt eine Series zurück, indiziert mit (Monat, Tag, Stunde), Werte in kW (für 1 kWp)."""
    params = {
        "lat": lat,
        "lon": lon,
        "pvcalculation": 1,
        "peakpower": PEAKPOWER_KWP,
        "loss": SYSTEM_VERLUST_PROZENT,
        "angle": NEIGUNG_GRAD,
        "aspect": AUSRICHTUNG_GRAD,
        "mountingplace": "building",
        "startyear": start_jahr,
        "endyear": end_jahr,
        "outputformat": "json",
        "usehorizon": 1,
    }
    resp = session.get(PVGIS_URL, params=params, timeout=60)
    resp.raise_for_status()
    data = resp.json()

    hourly = data["outputs"]["hourly"]
    df = pd.DataFrame(hourly)
    # PVGIS liefert "time" im Format "YYYYMMDD:HHMM" (UTC) und "P" = PV-Leistung in W
    df["timestamp"] = pd.to_datetime(df["time"], format="%Y%m%d:%H%M", utc=True)
    df["Monat"] = df["timestamp"].dt.month
    df["Tag"] = df["timestamp"].dt.day
    df["Stunde"] = df["timestamp"].dt.hour
    df["P_kW"] = df["P"].astype(float) / 1000.0  # W -> kW (für 1 kWp installiert)

    # 29. Februar rauswerfen, damit über alle Jahre sauber auf 365 Tage gemittelt werden kann
    df = df[~((df["Monat"] == 2) & (df["Tag"] == 29))]

    return df.set_index(["Monat", "Tag", "Stunde"])["P_kW"]


def baue_durchschnittsprofil():
    """Holt für alle Standorte/Jahre die PVGIS-Serie und mittelt alles zu einem
    einzigen typischen stündlichen Jahresprofil (Monat, Tag, Stunde) -> kW."""
    session = requests.Session()
    alle_serien = []

    for name, (lat, lon) in STANDORTE.items():
        print(f"Lade PVGIS-Daten für {name} ({lat}, {lon}) ...")
        serie = hole_pvgis_serie(lat, lon, START_JAHR, END_JAHR, session)
        alle_serien.append(serie)
        time.sleep(1)  # kurze Pause, um die PVGIS-API nicht zu überlasten

    # Alle Serien (verschiedene Standorte, je mehrere Jahre) zusammen mitteln
    kombiniert = pd.concat(alle_serien, axis=0)
    durchschnitt = kombiniert.groupby(level=["Monat", "Tag", "Stunde"]).mean()
    durchschnitt = durchschnitt.sort_index()
    return durchschnitt


def erzeuge_15min_zeitreihe(stunden_profil, jahr):
    """Baut aus dem stündlichen (Monat,Tag,Stunde)-Durchschnittsprofil eine
    vollständige 15-Minuten-Zeitreihe für das gegebene Kalenderjahr."""
    start = pd.Timestamp(year=jahr, month=1, day=1, hour=0, minute=0)
    end = pd.Timestamp(year=jahr + 1, month=1, day=1, hour=0, minute=0)

    stunden_index = pd.date_range(start, end, freq="1h", inclusive="left")
    werte = [
        stunden_profil.loc[(ts.month, ts.day, ts.hour)]
        for ts in stunden_index
    ]
    stunden_serie = pd.Series(werte, index=stunden_index)

    # Zusätzlichen Punkt am Jahresende anhängen (= erster Wert, für saubere Interpolation
    # bis exakt Mitternacht des Folgejahres), dann linear auf 15 Minuten interpolieren.
    stunden_serie.loc[end] = stunden_serie.iloc[0]
    viertelstunden_index = pd.date_range(start, end, freq="15min", inclusive="left")
    serie_15min = stunden_serie.reindex(stunden_serie.index.union(viertelstunden_index))
    serie_15min = serie_15min.interpolate(method="time")
    serie_15min = serie_15min.reindex(viertelstunden_index)

    # Kleine negative Interpolationsartefakte (z.B. kurz vor Sonnenaufgang) auf 0 setzen
    serie_15min = serie_15min.clip(lower=0)

    return serie_15min


def main():
    stunden_profil = baue_durchschnittsprofil()
    serie_15min = erzeuge_15min_zeitreihe(stunden_profil, OUTPUT_JAHR)

    out = pd.DataFrame({
        "Zeit": serie_15min.index,
        "CH_PV_normiert_1kWp [kW]": serie_15min.values,
    })

    out_path = "PV_Mittelland_normiert_1kWp_15min.csv"
    out.to_csv(out_path, index=False, sep=";", decimal=",", date_format="%d.%m.%Y %H:%M")
    print(f"Fertig: {len(out)} Zeilen gespeichert in {out_path}")
    print(f"Jahressumme (spezifischer Ertrag): {serie_15min.sum() / 4:.1f} kWh/kWp")


if __name__ == "__main__":
    main()