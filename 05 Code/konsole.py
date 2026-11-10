# -*- coding: utf-8 -*-
"""
konsole.py -- schlanke Konsolenausgabe für Batterieanalyse und Sizing-Sweep.

Standard ist der KOMPAKTE Modus: In der Konsole erscheinen nur Schritte,
Fortschritt und Warnungen/Fehler. Alles andere (Solver-Log, Detailmeldungen)
landet vollständig in einer Log-Datei im Output-Ordner.

Vollständige Ausgabe wie früher:   set BATTERIE_KONSOLE=voll
"""

from __future__ import annotations

import contextlib
import io
import os
import sys
import time
import warnings

VOLL = os.environ.get("BATTERIE_KONSOLE", "kompakt").strip().lower() == "voll"

_ZEILE_OFFEN = False   # steht der Cursor nach "[n/x] Text ..." noch in derselben Zeile?

# Zeilen mit diesen Wörtern erscheinen auch im kompakten Modus in der Konsole
WICHTIG = ("WARNUNG", "FEHLER", "ERROR", "Error", "Traceback", "ACHTUNG")


class _Filter(io.TextIOBase):
    """Schreibt alles ins Log und reicht nur wichtige Zeilen an die Konsole weiter."""

    def __init__(self, konsole, log):
        self.konsole, self.log, self._puffer = konsole, log, ""

    def write(self, text):
        if self.log:
            self.log.write(text)
        self._puffer += text
        while "\n" in self._puffer:
            zeile, self._puffer = self._puffer.split("\n", 1)
            if any(w in zeile for w in WICHTIG):
                global _ZEILE_OFFEN
                if _ZEILE_OFFEN:
                    self.konsole.write("\n")
                    _ZEILE_OFFEN = False
                self.konsole.write("      " + zeile.strip() + "\n")
                self.konsole.flush()
        return len(text)

    def flush(self):
        if self.log:
            self.log.flush()


@contextlib.contextmanager
def leise(logpfad: str | None = None):
    """Leitet stdout/stderr ins Log um (kompakter Modus). Im vollen Modus: keine Änderung."""
    if VOLL:
        yield
        return
    log = open(logpfad, "a", encoding="utf-8", errors="replace") if logpfad else None
    konsole = sys.__stdout__ or sys.stdout
    f = _Filter(konsole, log)
    try:
        with contextlib.redirect_stdout(f), contextlib.redirect_stderr(f), warnings.catch_warnings():
            warnings.simplefilter("ignore")
            yield
    finally:
        f.flush()
        if log:
            log.close()


def dauer(sekunden: float) -> str:
    s = int(round(sekunden))
    if s < 60:
        return f"{s} s"
    h, rest = divmod(s, 3600)
    m, s = divmod(rest, 60)
    return f"{h}:{m:02d}:{s:02d}" if h else f"{m}:{s:02d} min"


@contextlib.contextmanager
def schritt(nr: int, gesamt: int, text: str, logpfad: str | None = None):
    """'[1/4] Text ... fertig (12 s)' -- Details während des Schritts ins Log."""
    global _ZEILE_OFFEN
    t0 = time.time()
    print(f"  [{nr}/{gesamt}] {text} ...", end="", flush=True)
    _ZEILE_OFFEN = not VOLL
    try:
        with leise(logpfad):
            yield
    except Exception:
        print(f" FEHLER nach {dauer(time.time() - t0)}", flush=True)
        _ZEILE_OFFEN = False
        raise
    prefix = " " if _ZEILE_OFFEN else "        "
    _ZEILE_OFFEN = False
    print(f"{prefix}fertig ({dauer(time.time() - t0)})", flush=True)


def zahl(x: float, dez: int = 0) -> str:
    """Schweizer Tausendertrennung: 12'345"""
    return f"{x:,.{dez}f}".replace(",", "'")


def log_neu(logpfad: str, titel: str) -> str:
    """Log-Datei für einen neuen Lauf anlegen (überschreibt das alte Log)."""
    with open(logpfad, "w", encoding="utf-8") as f:
        f.write(f"{titel}\n{time.strftime('%d.%m.%Y %H:%M:%S')}\n{'=' * 70}\n")
    return logpfad
