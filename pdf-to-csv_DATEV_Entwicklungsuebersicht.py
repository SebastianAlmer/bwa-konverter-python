#!/usr/bin/env python3
"""
DATEV BWA Jahresentwicklungsuebersicht (eine Seite) -> CSV.

Usage:
  python pdf-to-csv_DATEV_Entwicklungsuebersicht.py --pdf "input/BWA 2025.12.pdf" --out output/jahresentwicklung_2025_12.csv
  python pdf-to-csv_DATEV_Entwicklungsuebersicht.py --batch --input-dir input --output-dir output

- Liest genau die angegebene Seite (1-basiert) aus der PDF.
- Ohne --page wird die Seite mit "Entwicklungsuebersicht" gesucht.
- Erwartet 13 Monats-Spalten im Kopf (z. B. "Sep/2024" oder "März 24") und haelt das DE-Zahlenformat bei.
- Trenner ist Semikolon; Ausgabe nutzt UTF-8 mit BOM, damit Excel Umlaute korrekt oeffnet.
- Mit --batch werden alle PDFs im Input-Ordner verarbeitet.
- Optional: Struktur-CSV kann zur Pruefung der Zeilenanzahl genutzt werden.
- Standardmaessig wird zusaetzlich eine .xlsx geschrieben (mit --no-excel deaktivierbar).
- Ohne Parameter laeuft ein Batch von input nach output.
"""
import argparse
import csv
import re
import sys
from pathlib import Path

import pdfplumber

MONTHS_RE = re.compile(
    r"\b(?:Jan(?:uar)?|Feb(?:ruar)?|M(?:ärz?|aerz?|rz|ar)|Apr(?:il)?|Mai|Jun(?:i)?|"
    r"Jul(?:i)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Okt(?:ober)?|Nov(?:ember)?|Dez(?:ember)?)\.?"
    r"\s*(?:[/\.\-]\s*|\s+)\d{2,4}\b"
)
DE_NUMBER_RE = re.compile(r"-?\d{1,3}(?:\.\d{3})*,\d{2}|-?0,00")
ENTWICKLUNGSUEBERSICHT_TERMS = ("entwicklungsübersicht", "entwicklungsuebersicht")
SEPARATOR_CLEAN_RE = re.compile(r"\s*([/.\-])\s*")

SECTION_BREAK_AFTER = {
    "Aktivierte Eigenleistungen",
    "Gesamtleistung",
    "Material-/Wareneinkauf",
    "Rohertrag",
    "So. betr. Erlöse",
    "Betrieblicher Rohertrag",
    "Gesamtkosten",
    "Betriebsergebnis",
    "Neutraler Aufwand",
    "Neutraler Ertrag",
    "Kontenklasse unbesetzt",
    "Ergebnis vor Steuern",
    "Steuern Einkommen u. Ertrag",
    "Vorläufiges Ergebnis",
}
COST_LABELS = {
    "Personalkosten",
    "Raumkosten",
    "Betriebliche Steuern",
    "Versicherungen/Beiträge",
    "Besondere Kosten",
    "Fahrzeugkosten (ohne Steuer)",
    "Werbe-/Reisekosten",
    "Kosten Warenabgabe",
    "Abschreibungen",
    "Reparatur/Instandhaltung",
    "Sonstige Kosten",
    "Gesamtkosten",
}


def normalize_month_token(token: str):
    token = " ".join(token.split())
    token = SEPARATOR_CLEAN_RE.sub(r"\1", token)
    return token


def safe_print(message: str):
    try:
        print(message)
    except UnicodeEncodeError:
        encoding = sys.stdout.encoding or "utf-8"
        safe_message = message.encode(encoding, errors="backslashreplace").decode(
            encoding, errors="ignore"
        )
        print(safe_message)


def extract_month_tokens(text: str):
    months = [m.group(0) for m in MONTHS_RE.finditer(text)]
    seen = set()
    ordered = []
    for m in months:
        normalized = normalize_month_token(m)
        if normalized not in seen:
            seen.add(normalized)
            ordered.append(normalized)
    return ordered


def detect_month_header(text: str):
    for raw_line in text.splitlines():
        if "Bezeichnung" not in raw_line:
            continue
        ordered = extract_month_tokens(raw_line)
        if len(ordered) >= 13:
            return ordered[-13:]

    ordered = extract_month_tokens(text)
    if len(ordered) >= 13:
        return ordered[-13:]
    return None


def parse_rows_from_text(text: str, months):
    rows = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if line == "Kostenarten:":
            rows.append((line, [""] * len(months)))
            continue
        nums = DE_NUMBER_RE.findall(line)
        if len(nums) >= len(months):
            values = nums[-len(months):]
            first_num = values[0]
            pos = line.find(first_num)
            label = line[:pos].strip()
            rows.append((label, values))
    return rows


def ensure_kostenarten(rows, months):
    labels = [label for label, _ in rows]
    if "Kostenarten:" in labels:
        return rows
    first_cost_idx = next((i for i, (label, _) in enumerate(rows) if label in COST_LABELS), None)
    if first_cost_idx is None:
        return rows
    return rows[:first_cost_idx] + [("Kostenarten:", [""] * len(months))] + rows[first_cost_idx:]


def insert_section_breaks(rows, months_count):
    out = [("", [""] * months_count)]  # blank row after header
    for label, values in rows:
        out.append((label, values))
        if label in SECTION_BREAK_AFTER:
            out.append(("", [""] * months_count))
    out.append(("", [""] * months_count))  # trailing blank row
    return out


def compress_blank_rows(rows):
    cleaned = []
    prev_blank = False
    for label, values in rows:
        is_blank = not label and (not values or all(v == "" for v in values))
        if is_blank and prev_blank:
            continue
        cleaned.append((label, values))
        prev_blank = is_blank
    return cleaned


def build_output_table(header, final_rows, structure_numbers):
    if structure_numbers is not None and len(structure_numbers) != len(final_rows):
        raise RuntimeError(
            "Struktur hat "
            f"{len(structure_numbers)} Zeilen, Ergebnis hat {len(final_rows)} Zeilen."
        )

    columns = ["Bezeichnung"] + header
    rows = []
    for label, values in final_rows:
        if not label and values and all(v == "" for v in values):
            rows.append([""] + [""] * len(header))
        else:
            rows.append([label] + values)
    return columns, rows


def write_csv_table(columns, rows, out_path: Path):
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(columns)
        for row in rows:
            writer.writerow(row)


def write_excel_table(columns, rows, out_path: Path):
    try:
        import pandas as pd
    except ImportError as exc:
        raise RuntimeError("Pandas ist nicht installiert. Excel-Ausgabe nicht moeglich.") from exc

    out_path.parent.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(rows, columns=columns)
    df.to_excel(out_path, index=False)


def find_entwicklungsuebersicht_page(pdf):
    for index, page in enumerate(pdf.pages, start=1):
        page_text = page.extract_text() or ""
        lowered = page_text.casefold()
        if any(term in lowered for term in ENTWICKLUNGSUEBERSICHT_TERMS):
            return index, page_text
    return None, ""


def convert_page_to_csv(
    pdf_path: Path,
    page_number: int | None,
    out_path: Path,
    structure_numbers: list[str] | None = None,
    write_excel_file: bool = False,
    excel_path: Path | None = None,
):
    with pdfplumber.open(pdf_path) as pdf:
        if page_number is None:
            page_number, page_text = find_entwicklungsuebersicht_page(pdf)
            if page_number is None:
                raise RuntimeError("Keine Seite mit 'Entwicklungsuebersicht' in der PDF gefunden.")
        else:
            if page_number < 1:
                raise ValueError("page_number muss 1-basiert sein")
            if page_number > len(pdf.pages):
                raise ValueError(f"PDF hat nur {len(pdf.pages)} Seiten, angefragt wurde {page_number}.")
            page_text = pdf.pages[page_number - 1].extract_text() or ""

    header = detect_month_header(page_text)
    if not header:
        raise RuntimeError("Konnte keine 13 Monats-Spalten auf der Seite erkennen.")

    rows = parse_rows_from_text(page_text, header)
    rows = ensure_kostenarten(rows, header)

    seen = set()
    uniq_rows = []
    for label, values in rows:
        key = (label, tuple(values))
        if key not in seen:
            seen.add(key)
            uniq_rows.append((label, values))

    final_rows = compress_blank_rows(insert_section_breaks(uniq_rows, len(header)))
    columns, rows = build_output_table(header, final_rows, structure_numbers)
    write_csv_table(columns, rows, out_path)
    if write_excel_file:
        excel_target = excel_path or out_path.with_suffix(".xlsx")
        write_excel_table(columns, rows, excel_target)

    return out_path


def build_output_paths(out_dir: Path, pdf_path: Path, excel_dir: Path | None, write_excel_file: bool):
    csv_path = out_dir / f"{pdf_path.stem}_jahresentwicklung.csv"
    excel_path = None
    if write_excel_file:
        target_dir = excel_dir or out_dir
        excel_path = target_dir / f"{pdf_path.stem}_jahresentwicklung.xlsx"
    return csv_path, excel_path


def load_structure_numbers(structure_path: Path):
    with structure_path.open(encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f, delimiter=";")
        header = next(reader, None)
        if not header or len(header) < 2:
            raise RuntimeError(f"Ungueltige Struktur-CSV: {structure_path}")
        numbers = []
        for row in reader:
            if not row:
                continue
            numbers.append(row[0].strip())
    return numbers


def convert_batch(
    input_dir: Path,
    out_dir: Path,
    page_number: int | None,
    write_excel_file: bool,
    excel_dir: Path | None,
    structure_numbers: list[str] | None = None,
):
    pdf_paths = sorted(input_dir.glob("*.pdf"))
    if not pdf_paths:
        raise RuntimeError(f"Keine PDF-Dateien in {input_dir} gefunden.")

    out_dir.mkdir(parents=True, exist_ok=True)
    if write_excel_file and excel_dir is not None:
        excel_dir.mkdir(parents=True, exist_ok=True)
    written = []
    skipped = []
    for pdf_path in pdf_paths:
        out_path, excel_path = build_output_paths(out_dir, pdf_path, excel_dir, write_excel_file)
        if out_path.exists() and (not write_excel_file or (excel_path and excel_path.exists())):
            safe_print(f"INFO: Uebersprungen (bereits vorhanden): {out_path}")
            continue
        try:
            convert_page_to_csv(
                pdf_path,
                page_number,
                out_path,
                structure_numbers,
                write_excel_file,
                excel_path,
            )
            written.append(out_path)
            safe_print(f"Geschrieben: {out_path}")
            if write_excel_file and excel_path is not None:
                safe_print(f"Geschrieben: {excel_path}")
        except Exception as exc:
            skipped.append((pdf_path, str(exc)))
            safe_print(f"Uebersprungen: {pdf_path} ({exc})")
    return written, skipped


def pick_default_pdf():
    candidates = sorted(Path("input").glob("*.pdf"))
    for p in candidates:
        if "bwa" in p.stem.lower():
            return p
    return candidates[0] if candidates else None


def main():
    parser = argparse.ArgumentParser(
        description="DATEV BWA Jahresentwicklungsuebersicht einer Seite in CSV umwandeln"
    )
    parser.add_argument("--pdf", type=Path, default=None, help="Pfad zur PDF-Datei")
    parser.add_argument(
        "--batch",
        action="store_true",
        help="Alle PDFs im Input-Ordner verarbeiten",
    )
    parser.add_argument(
        "--input-dir",
        type=Path,
        default=Path("input"),
        help="Input-Ordner fuer --batch",
    )
    parser.add_argument(
        "--page",
        type=int,
        default=None,
        help="1-basierte Seitennummer; ohne Angabe wird nach 'Entwicklungsuebersicht' gesucht",
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=Path("output") / "jahresentwicklung.csv",
        help="Pfad zur Ausgabe-CSV",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("output"),
        help="Ausgabe-Ordner fuer --batch",
    )
    parser.add_argument(
        "--excel-dir",
        type=Path,
        default=None,
        help="Excel-Ausgabeordner fuer --batch (Standard: output-dir)",
    )
    parser.add_argument(
        "--structure",
        type=Path,
        default=None,
        help="Pfad zur Struktur-CSV fuer die Zeilenanzahl-Pruefung",
    )
    parser.add_argument(
        "--excel",
        type=Path,
        default=None,
        help="Pfad zur Excel-Ausgabe (Einzelmodus)",
    )
    parser.add_argument(
        "--no-excel",
        action="store_true",
        help="Excel-Ausgabe deaktivieren",
    )
    raw_args = sys.argv[1:]
    if not raw_args:
        raw_args = ["--batch"]
    args = parser.parse_args(raw_args)

    structure_path = None
    if args.structure is not None:
        structure_path = args.structure
        if not structure_path.exists():
            parser.error(f"Struktur-CSV nicht gefunden: {structure_path}")
    else:
        default_structure = Path("DATEV Struktur") / "BWA Export Datei -leer -.csv"
        if default_structure.exists():
            structure_path = default_structure

    structure_numbers = None
    if structure_path is not None:
        structure_numbers = load_structure_numbers(structure_path)

    write_excel_file = not args.no_excel

    if args.batch:
        input_dir = args.input_dir
        if args.pdf is not None:
            if args.pdf.is_dir():
                input_dir = args.pdf
            else:
                parser.error("--batch erwartet einen Ordner; --pdf zeigt auf eine Datei.")
        if not input_dir.exists():
            parser.error(f"Input-Ordner nicht gefunden: {input_dir}")
        _, skipped = convert_batch(
            input_dir,
            args.output_dir,
            args.page,
            write_excel_file,
            args.excel_dir,
            structure_numbers,
        )
        if skipped:
            safe_print(f"{len(skipped)} PDFs uebersprungen.")
        return

    pdf_path = args.pdf or pick_default_pdf()
    if pdf_path is None:
        parser.error("Keine PDF angegeben und keine Datei in ./input gefunden.")
    if not pdf_path.exists():
        parser.error(f"PDF nicht gefunden: {pdf_path}")

    out_path = convert_page_to_csv(
        pdf_path,
        args.page,
        args.out,
        structure_numbers,
        write_excel_file,
        args.excel,
    )
    safe_print(f"Geschrieben: {out_path}")
    if write_excel_file:
        excel_path = args.excel or args.out.with_suffix(".xlsx")
        safe_print(f"Geschrieben: {excel_path}")


if __name__ == "__main__":
    main()
