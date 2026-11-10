# -*- coding: utf-8 -*-
"""
parameter_layout.py -- Parameter-Zellen über die BEZEICHNUNG in Spalte A finden.

Der Code adressiert Parameter historisch über feste Zellen (z.B. "C29" = Kapazität).
Seit Zeilen im Parameter-Blatt eingefügt werden (z.B. Netznutzung & Abgaben ab Zeile 14),
stimmen diese Adressen nicht mehr für jede Datei. `zelle(ws, "C29")` übersetzt die
historische Adresse in die tatsächliche Zeile der aktuellen Datei:

    ws[zelle(ws, "C29")].value      # Kapazität, egal in welcher Zeile sie steht

Funktioniert für alte Dateien (Originallayout) und neue Dateien (mit eingefügten Zeilen).
Wird eine Bezeichnung nicht gefunden, bleibt die Adresse unverändert.
"""

from __future__ import annotations

import re

# Originallayout: Zeile -> Bezeichnung in Spalte A (Stand vor dem Einfügen von Zeilen)
ALT_LABELS = {
    1: "Parameter", 2: "Beginn Analysezeitraum", 3: "Ende Analysezeitraum", 4: "Energie-Tarif",
    5: "Tarifschema", 6: "HT", 7: "NT", 8: "Netznutzung Leistung", 9: "Tage HT", 10: "Startzeit HT",
    11: "Endzeit HT", 12: "Spot_Abschlag_Lieferung", 13: "Spot_Aufschlag_Bezug", 14: "Rücklieferung_PV",
    19: "Rücklieferung_Batterie", 24: "PV", 25: "DC Leistung", 26: "Produktions-Input",
    27: "Abschaltung bei Negativpreisen", 28: "Batterie", 29: "Kapazität", 30: "Leistung",
    31: "Wirkungsgrad Transformator", 32: "Ladewirkungsgard", 33: "Entladewirkungsgard",
    34: "Investitionskosten pro kWh", 35: "Unterhaltkosten", 36: "Lebenszeit", 37: "Vollzyklen Lebensdauer",
    38: "Durchsatzbegrenzung/jahr", 39: "Ladegrenze SOC", 40: "Entladegrenze SOC", 41: "Netto Kosten",
    42: "Netz", 43: "Maximale Einspeisung", 44: "Maximaler Bezug", 45: "Maximale Einspeisung Batterie",
    46: "Maximaler Bezug Batterie", 48: "SRL", 49: "SRL-Leistung", 50: "SRL-Leistung-Preise",
    51: "SRL-Energie", 52: "SRL-Kundenfaktor", 53: "Last", 54: "Lastgang",
}
# Doppelte Bezeichnungen (HKN/Floor/Fixtarif bei PV und Batterie): relativ zum Abschnittstitel
RELATIV = {
    15: ("Rücklieferung_PV", 1), 16: ("Rücklieferung_PV", 2), 17: ("Rücklieferung_PV", 3),
    20: ("Rücklieferung_Batterie", 1), 21: ("Rücklieferung_Batterie", 2), 22: ("Rücklieferung_Batterie", 3),
}

_CACHE: dict[int, dict[str, int]] = {}


def _norm(txt) -> str:
    return " ".join(str(txt or "").lower().split())


def _zeilen(ws) -> dict[str, int]:
    key = id(ws)
    if key not in _CACHE:
        z = {}
        for r in range(1, (ws.max_row or 1) + 1):
            v = ws.cell(row=r, column=1).value
            if v is not None and _norm(v) not in z:
                z[_norm(v)] = r
        _CACHE[key] = z
    return _CACHE[key]


def zeile(ws, alte_zeile: int) -> int:
    """Tatsächliche Zeile zur historischen Zeilennummer."""
    z = _zeilen(ws)
    if alte_zeile in RELATIV:
        anker, off = RELATIV[alte_zeile]
        r = z.get(_norm(anker))
        return r + off if r else alte_zeile
    label = ALT_LABELS.get(alte_zeile)
    if label:
        return z.get(_norm(label), alte_zeile)
    return alte_zeile


def zelle(ws, adresse: str) -> str:
    """Historische Adresse (z.B. 'C29') -> Adresse in der aktuellen Datei."""
    m = re.fullmatch(r"([A-Z]+)(\d+)", adresse)
    if not m:
        return adresse
    return f"{m.group(1)}{zeile(ws, int(m.group(2)))}"


def cache_leeren():
    _CACHE.clear()
