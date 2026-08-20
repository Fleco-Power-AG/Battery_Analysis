"""
run_battery_analysis.py
========================

Einstiegspunkt, der den kompletten Ablauf orchestriert:
  1. Input-Excel einlesen und fehlende Daten ergaenzen (input_prep.py)
  2. oemof.solph-Optimierung (battery_optimization.py)
  3. Ergebnis schreiben (Format/Inhalt wird spaeter noch verfeinert -- aktuell
     ein einfacher Excel-Export der Zeitreihen + Modul-Ertraege)

WICHTIGES PRINZIP: Der CODE ist ortsunabhaengig. Dieses Skript (und
input_prep.py / battery_optimization.py) koennen in einem beliebigen
Ordner liegen (z.B. wo auch immer `git clone` das Repo hinlegt) -- das spielt
keine Rolle. Massgeblich ist NUR der Pfad der Inputs.xlsx, den der Nutzer
angibt: `input_lp.xlsx` und das Ergebnis landen automatisch im selben Ordner
wie die von ihm gewaehlte Inputs.xlsx, nicht im Code-Ordner.

Aufruf-Varianten:
  1. Mit explizitem Pfad (z.B. aus einer anderen Person/einem Skript heraus):
       python run_battery_analysis.py "C:\\Pfad\\zu\\Inputs.xlsx"
  2. Ohne Pfad, interaktiv (z.B. Doppelklick via run_battery_analysis.bat):
       python run_battery_analysis.py
     -> oeffnet einen Datei-Auswahl-Dialog, in dem die Inputs.xlsx ausgewaehlt
        werden kann. Kein Pfad-Tippen noetig.
  3. Per Drag & Drop: die Inputs.xlsx auf das Icon von run_battery_analysis.bat
     ziehen -- Windows uebergibt den Pfad dann automatisch als Argument.

Jeder Fleco-Nutzer kann also das Repo klonen, sein eigenes Inputs.xlsx an
einem beliebigen Ort (z.B. sein eigener SharePoint-/Dokumente-Ordner) haben
und dieses Skript darauf loslassen, ohne den Code anzufassen.
"""

from __future__ import annotations

import argparse
import os
import sys

import input_prep
import battery_optimization


def pick_file_dialog() -> str:
    """Oeffnet einen Datei-Auswahl-Dialog (funktioniert unter Windows/Mac/
    Linux mit Standard-Python -- tkinter ist Teil der Standardbibliothek)."""
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError as exc:
        raise RuntimeError(
            "Kein Pfad angegeben und der Datei-Dialog (tkinter) ist in dieser "
            "Python-Installation nicht verfuegbar. Bitte den Pfad direkt als "
            "Argument angeben:\n"
            "  python run_battery_analysis.py <Pfad-zu-Inputs.xlsx>"
        ) from exc

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askopenfilename(
        title="Inputs.xlsx auswaehlen",
        filetypes=[("Excel-Dateien", "*.xlsx"), ("Alle Dateien", "*.*")],
    )
    root.destroy()
    if not path:
        raise SystemExit("Keine Datei ausgewaehlt -- Abbruch.")
    return path


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Batterieanalyse: Input aufbereiten + oemof.solph-Optimierung."
    )
    parser.add_argument(
        "input_file", nargs="?", default=None,
        help="Pfad zur Inputs.xlsx. Falls weggelassen: Datei-Auswahl-Dialog.",
    )
    parser.add_argument("--solver", default="cbc", help="Solver (Default: cbc)")
    parser.add_argument(
        "--kein-swissix-fetch", action="store_true",
        help="ENTSO-E-Abfrage ueberspringen (z.B. ohne Internetzugang testen).",
    )
    args = parser.parse_args(argv)

    input_path = args.input_file or pick_file_dialog()
    input_path = os.path.abspath(input_path)

    if not os.path.isfile(input_path):
        raise SystemExit(f"Datei nicht gefunden: {input_path}")

    # Alles landet im selben Ordner wie die vom Nutzer gewaehlte Inputs.xlsx --
    # NICHT im Ordner, in dem dieser Code liegt.
    data_dir = os.path.dirname(input_path)
    input_lp_path = os.path.join(data_dir, "input_lp.xlsx")
    result_path = os.path.join(data_dir, "ergebnis.xlsx")

    print("=" * 70)
    print("Batterieanalyse")
    print(f"  Input:      {input_path}")
    print(f"  input_lp:   {input_lp_path}")
    print(f"  Ergebnis:   {result_path}")
    print("=" * 70)

    print("\n=== Schritt 1/2: Input aufbereiten (input_prep) ===")
    input_prep.main(
        input_path=input_path,
        output_path=input_lp_path,
        fetch_swissix=not args.kein_swissix_fetch,
    )

    print("\n=== Schritt 2/2: Optimierung (battery_optimization) ===")
    module, ts, anteile = battery_optimization.main(
        input_path=input_lp_path, solver=args.solver
    )

    # Ergebnis-Report: Format/Inhalt wird spaeter noch definiert -- vorerst ein
    # einfacher Excel-Export der Zeitreihen + Modul-Ertraege, damit man sofort
    # etwas in der Hand hat.
    import pandas as pd

    with pd.ExcelWriter(result_path) as writer:
        ts.join(anteile).to_excel(writer, sheet_name="Zeitreihen")
        pd.Series(module, name="CHF/Jahr").to_frame().to_excel(writer, sheet_name="Modul-Ertraege")

    print(f"\nFertig. Ergebnis gespeichert unter: {result_path}")
    return module, ts, anteile


if __name__ == "__main__":
    main()