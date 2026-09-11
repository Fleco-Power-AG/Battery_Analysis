"""
run_battery_analysis.py
========================

Einstiegspunkt, der den kompletten Ablauf orchestriert:
  1. Input-Excel einlesen und fehlende Daten ergaenzen (input_prep.py)
  2. oemof.solph-Optimierung (battery_optimization.py)
  3. Ergebnis schreiben als "Output-Kontrolle" (output_create.py): Zeitreihen-
     Tabelle (Energiepreis, PV, Batterieleistung, SOC, SRL, Ertragsspalten je
     Zeitschritt) + Modul-Ertrags-Zusammenfassung.
  4. PDF-Report erstellen (output_create.build_pdf_report()): Wirtschaftlich-
     keit, Renditebetrachtung, sowie Monats-/Quartals-Charts.

WICHTIGES PRINZIP: Der CODE ist ortsunabhaengig. Dieses Skript (und
input_prep.py / battery_optimization.py) koennen in einem beliebigen
Ordner liegen (z.B. wo auch immer `git clone` das Repo hinlegt) -- das spielt
keine Rolle. Massgeblich ist NUR der Pfad der Inputs.xlsx, den der Nutzer
angibt: der PDF-Report landet automatisch im selben Ordner wie die von ihm
gewaehlte Inputs.xlsx, nicht im Code-Ordner. Alle Dateinamen werden vom Namen
der gewaehlten Inputs-Datei abgeleitet (siehe
output_create.derive_input_lp_filename()/derive_result_filename()/
derive_pdf_filename()), damit sich bei mehreren Szenarien im selben Ordner
(z.B. "Inputs_default.xlsx", "batterie_4.xlsx") jede Zwischen-/Ergebnisdatei
eindeutig ihrem Input zuordnen laesst und nichts ueberschrieben wird.

NEU (Beats Wunsch, danach korrigiert): NUR der PDF-Report bleibt direkt im
Inputs-Ordner. ALLES ANDERE -- die aufbereitete Zwischen-Datei
(`input_lp_<name>.xlsx`), das Ergebnis-Excel (`Ergebnis_<name>.xlsx`) UND das
Solver-Log (`.log`, siehe battery_optimization.main() -- landet automatisch
mit, da es sich vom Ordner der input_lp-Datei ableitet) -- landet in einem
"Output"-Unterordner des Inputs-Ordners (Inputs-Ordner + "\\Output"). Dieser
Unterordner wird BEWUSST NICHT automatisch angelegt -- existiert er nicht,
bricht das Skript mit einer klaren Meldung ab, BEVOR die (potenziell lange
laufende) Optimierung startet.

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
import output_create


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
    parser.add_argument(
        "--solver", default="cbc",
        help="Solver (Default: cbc). Standard-Setup: Conda-Umgebung mit "
             "'conda install -c conda-forge coincbc', dann diese Umgebung "
             "aktivieren, bevor das Skript laeuft. Alternativen z.B. "
             "'appsi_highs' (via 'pip install highspy', ohne Conda nutzbar).",
    )
    parser.add_argument(
        "--kein-swissix-fetch", action="store_true",
        help="ENTSO-E-Abfrage ueberspringen (z.B. ohne Internetzugang testen).",
    )
    args = parser.parse_args(argv)

    input_path = args.input_file or pick_file_dialog()
    input_path = os.path.abspath(input_path)

    if not os.path.isfile(input_path):
        raise SystemExit(f"Datei nicht gefunden: {input_path}")

    # Der PDF-Report landet im selben Ordner wie die vom Nutzer gewaehlte
    # Inputs.xlsx -- NICHT im Ordner, in dem dieser Code liegt.
    data_dir = os.path.dirname(input_path)

    # NEU (Beats Wunsch, KORRIGIERT -- urspruenglich nur fuer input_lp gedacht,
    # gilt aber fuer ALLES ausser dem PDF-Report): die input_lp-Zwischen-
    # Datei, das Ergebnis-Excel UND das Solver-Log (siehe battery_optimization.
    # main(), das Logfile leitet sich vom Ordner von input_path == input_lp_path
    # ab und landet dadurch automatisch hier mit) landen in einem "Output"-
    # Unterordner des Inputs-Ordners. NUR der PDF-Report bleibt im Inputs-
    # Ordner selbst. Der Output-Ordner wird bewusst NICHT automatisch angelegt
    # (kein os.makedirs). Fehlt er, bricht das Skript HIER ab (vor Schritt
    # 1/2), statt erst nach der (potenziell lange laufenden) Optimierung beim
    # Speichern zu scheitern.
    output_dir = os.path.join(data_dir, "Output")
    if not os.path.isdir(output_dir):
        raise SystemExit(
            f"Output-Ordner nicht gefunden: {output_dir}\n"
            "input_lp-Datei, Ergebnis-Excel und Solver-Log werden nur in "
            "einen bereits vorhandenen 'Output'-Unterordner (neben der "
            "Inputs.xlsx) geschrieben -- dieser Ordner wird bewusst NICHT "
            "automatisch angelegt. Bitte den Ordner 'Output' manuell in\n"
            f"  {data_dir}\n"
            "erstellen und das Skript erneut starten."
        )

    # input_lp/Ergebnis/Bericht-Dateinamen alle vom Input-Dateinamen abgeleitet,
    # damit man bei mehreren Szenarien im selben Ordner (z.B. "Inputs_default.xlsx",
    # "batterie_4.xlsx") nicht jedes Mal dieselben Dateien ueberschreibt UND jede
    # Zwischen-/Ergebnisdatei eindeutig ihrem Input zuordenbar bleibt:
    # -> "input_lp_default.xlsx", "Ergebnis_default.xlsx", "Bericht_default.pdf"
    # (siehe output_create.derive_input_lp_filename()/derive_result_filename()/
    # derive_pdf_filename()).
    input_lp_path = os.path.join(output_dir, output_create.derive_input_lp_filename(input_path))
    result_path = os.path.join(output_dir, output_create.derive_result_filename(input_path))
    pdf_path = os.path.join(data_dir, output_create.derive_pdf_filename(input_path))

    print("=" * 70)
    print("Batterieanalyse")
    print(f"  Input:      {input_path}")
    print(f"  input_lp:   {input_lp_path}")
    print(f"  Ergebnis:   {result_path}")
    print(f"  PDF-Report: {pdf_path}")
    print("=" * 70)

    print("\n=== Schritt 1/2: Input aufbereiten (input_prep) ===")
    input_prep.main(
        input_path=input_path,
        output_path=input_lp_path,
        fetch_swissix=not args.kein_swissix_fetch,
    )

    print("\n=== Schritt 2/2: Optimierung (battery_optimization) ===")
    module, ts, anteile, zeitreihen, params, pv_profile, last_profile = battery_optimization.main(
        input_path=input_lp_path, solver=args.solver
    )

    print("\n=== Output-Kontrolle schreiben (output_create) ===")
    zeitreihen_tabelle, modul_tabelle = output_create.save_output_excel(
        result_path, module, ts, anteile, zeitreihen, params, pv_profile, last_profile,
    )

    print("\n=== PDF-Report erstellen (output_create) ===")
    # entladeenergie_kwh mitgeben, damit die Degradationskosten (siehe
    # battery_optimization.DEGRADATIONSKOSTEN_CHF_PRO_KWH), die schon IN der
    # LP-Zielfunktion beruecksichtigt sind, auch im PDF-Report als eigene
    # Ausgaben-Zeile auftauchen (sonst wuerden "Total Einnahmen - Total
    # Ausgaben" nicht mit dem tatsaechlichen Optimierungsergebnis uebereinstimmen).
    entladeenergie_kwh = float(ts["dis"].sum() * battery_optimization.DT_HOURS)
    kapitalkosten = battery_optimization.capital_costs(params, entladeenergie_kwh=entladeenergie_kwh)
    output_create.build_pdf_report(
        pdf_path, module, ts, anteile, zeitreihen, params, pv_profile, last_profile, kapitalkosten,
    )

    print(f"\nFertig. Ergebnis gespeichert unter: {result_path}")
    print(f"PDF-Report gespeichert unter: {pdf_path}")
    return module, ts, anteile, zeitreihen_tabelle, modul_tabelle


if __name__ == "__main__":
    main()