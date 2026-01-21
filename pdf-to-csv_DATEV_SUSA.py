#!/usr/bin/env python3
"""
Extracts a DATEV Summen- und Saldenliste (SuSa) PDF into CSV.

Usage:
  python pdf-to-csv_DATEV_SUSA.py --pdf input/your_file.pdf --out output/susa.csv [--start-page N] [--end-page N]
  python pdf-to-csv_DATEV_SUSA.py --batch --input-dir input --output-dir output

Defaults:
- Ohne --pdf wird die erste PDF in ./input genommen.
- Wenn --start-page/--end-page fehlen, wird nach Seiten mit "Summen und Salden" gesucht.
- CSV nutzt Semikolon, Zahlen sind signiert (Soll=+, Haben=-).
- Standardmaessig wird zusaetzlich eine .xlsx geschrieben (mit --no-excel deaktivierbar).
- Ohne Parameter laeuft ein Batch von input nach output.
"""
import argparse
import csv
import re
import sys
from pathlib import Path

import pdfplumber

# Column slices based on measured x positions in the sample SuSa PDF.
# Adjust here if the layout moves in other reports.
COLUMN_BOUNDS = [
    ("Konto", 0, 90),
    ("Beschriftung", 90, 320),
    ("EB-Wert", 320, 410),
    ("Okt 2025 Soll", 410, 500),
    ("Okt 2025 Haben", 500, 580),
    ("Kum Werte Soll", 580, 660),
    ("Kum Werte Haben", 660, 750),
    ("Saldo", 750, 900),
]
SUSA_HEADER_RE = re.compile(r"summen\s*(?:-\s*)?und\s*salden", re.IGNORECASE)


def collect_lines(page, bucket=1):
    lines = {}
    for ch in page.chars:
        y = round(ch["top"] / bucket) * bucket
        lines.setdefault(y, []).append(ch)
    return [(y, lines[y]) for y in sorted(lines.keys())]


def slice_columns(chars):
    chars = sorted(chars, key=lambda c: c["x0"])
    columns = {}
    for name, start, end in COLUMN_BOUNDS:
        columns[name] = "".join(
            ch["text"] for ch in chars if start <= ch["x0"] < end
        ).strip()
    return columns


def parse_amount(text, signed=False):
    if not text:
        return 0.0
    cleaned = text.replace(" ", "")
    sign = 1
    if cleaned.endswith(("S", "H")):
        if signed:
            sign = -1 if cleaned.endswith("H") else 1
        cleaned = cleaned[:-1]
    cleaned = cleaned.replace(".", "").replace(",", ".")
    try:
        return sign * float(cleaned)
    except ValueError:
        return 0.0


def safe_print(message: str):
    try:
        print(message)
    except UnicodeEncodeError:
        encoding = sys.stdout.encoding or "utf-8"
        safe_message = message.encode(encoding, errors="backslashreplace").decode(
            encoding, errors="ignore"
        )
        print(safe_message)


def find_susa_page_range(pdf):
    matches = []
    for index, page in enumerate(pdf.pages, start=1):
        page_text = page.extract_text() or ""
        if SUSA_HEADER_RE.search(page_text):
            matches.append(index)
    if not matches:
        return None, None
    return matches[0], matches[-1]


def parse_pdf(pdf_path: Path, start_page: int | None = None, end_page: int | None = None):
    konto_re = re.compile(r"^\d{3,4}\s*\d{2}$")
    rows = []
    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        if start_page is None and end_page is None:
            start_page, end_page = find_susa_page_range(pdf)
            if start_page is None:
                raise RuntimeError("Keine Seite mit 'Summen und Salden' in der PDF gefunden.")
        if start_page is None:
            start_page = 1
        if end_page is None:
            end_page = total_pages

        start_idx = max(start_page - 1, 0)
        end_idx = min(end_page, total_pages)
        if start_idx >= total_pages:
            raise ValueError(
                f"Startseite {start_page} liegt ausserhalb der PDF mit {total_pages} Seiten."
            )
        if end_idx <= start_idx:
            raise ValueError(
                f"Endseite {end_page} muss groesser als Startseite {start_page} sein."
            )

        for page in pdf.pages[start_idx:end_idx]:
            for _, line_chars in collect_lines(page):
                cols = slice_columns(line_chars)
                konto = cols["Konto"].strip()
                if not konto_re.match(konto):
                    continue
                rows.append(
                    {
                        "Konto": konto,
                        "Beschriftung": cols["Beschriftung"],
                        "EB-Wert": parse_amount(cols["EB-Wert"], signed=True),
                        "Okt 2025 Soll": parse_amount(cols["Okt 2025 Soll"]),
                        "Okt 2025 Haben": parse_amount(cols["Okt 2025 Haben"]),
                        "Kum Werte Soll": parse_amount(cols["Kum Werte Soll"]),
                        "Kum Werte Haben": parse_amount(cols["Kum Werte Haben"]),
                        "Saldo": parse_amount(cols["Saldo"], signed=True),
                    }
                )
    return rows


def write_csv(rows, out_path: Path):
    out_path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = [name for name, _, _ in COLUMN_BOUNDS]
    # Use UTF-8 with BOM so Excel (deutsche Lokalisierung) oeffnet Umlaute korrekt.
    with out_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=";")
        writer.writeheader()
        fmt = lambda v: f"{v:.2f}".replace(".", ",")
        for row in rows:
            writer.writerow(
                {
                    "Konto": row["Konto"],
                    "Beschriftung": row["Beschriftung"],
                    "EB-Wert": fmt(row["EB-Wert"]),
                    "Okt 2025 Soll": fmt(row["Okt 2025 Soll"]),
                    "Okt 2025 Haben": fmt(row["Okt 2025 Haben"]),
                    "Kum Werte Soll": fmt(row["Kum Werte Soll"]),
                    "Kum Werte Haben": fmt(row["Kum Werte Haben"]),
                    "Saldo": fmt(row["Saldo"]),
                }
            )


def write_excel(rows, out_path: Path):
    try:
        import pandas as pd
    except ImportError as exc:
        raise RuntimeError("Pandas ist nicht installiert. Excel-Ausgabe nicht moeglich.") from exc

    out_path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = [name for name, _, _ in COLUMN_BOUNDS]
    df = pd.DataFrame(rows, columns=fieldnames)
    try:
        df.to_excel(out_path, index=False)
    except Exception as exc:
        raise RuntimeError(f"Excel-Ausgabe fehlgeschlagen: {exc}") from exc


def build_output_paths(out_dir: Path, pdf_path: Path, excel_dir: Path | None, write_excel_file: bool):
    csv_path = out_dir / f"{pdf_path.stem}_susa.csv"
    excel_path = None
    if write_excel_file:
        target_dir = excel_dir or out_dir
        excel_path = target_dir / f"{pdf_path.stem}_susa.xlsx"
    return csv_path, excel_path


def convert_batch(
    input_dir: Path,
    out_dir: Path,
    start_page: int | None,
    end_page: int | None,
    write_excel_file: bool,
    excel_dir: Path | None,
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
        csv_path, excel_path = build_output_paths(out_dir, pdf_path, excel_dir, write_excel_file)
        if csv_path.exists() and (not write_excel_file or (excel_path and excel_path.exists())):
            safe_print(f"INFO: Uebersprungen (bereits vorhanden): {csv_path}")
            continue
        try:
            rows = parse_pdf(pdf_path, start_page=start_page, end_page=end_page)
            write_csv(rows, csv_path)
            if write_excel_file and excel_path is not None:
                write_excel(rows, excel_path)
            written.append(csv_path)
            safe_print(f"Geschrieben: {csv_path}")
            if write_excel_file and excel_path is not None:
                safe_print(f"Geschrieben: {excel_path}")
        except Exception as exc:
            skipped.append((pdf_path, str(exc)))
            safe_print(f"Uebersprungen: {pdf_path} ({exc})")
    return written, skipped


def main():
    parser = argparse.ArgumentParser(description="SuSa PDF in CSV/Excel umwandeln")
    parser.add_argument("--pdf", type=Path, default=None, help="Pfad zur PDF")
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
        "--out", type=Path, default=Path("output") / "susa.csv", help="Pfad zur Ausgabe-CSV"
    )
    parser.add_argument(
        "--excel", type=Path, default=None, help="Pfad zur Excel-Ausgabe (Einzelmodus)"
    )
    parser.add_argument(
        "--no-excel", action="store_true", help="Excel-Ausgabe deaktivieren"
    )
    parser.add_argument(
        "--start-page",
        type=int,
        default=None,
        help="1-basierte Startseite (inklusive); ohne Angaben wird gesucht",
    )
    parser.add_argument(
        "--end-page",
        type=int,
        default=None,
        help="1-basierte Endseite (inklusive); ohne Angaben wird gesucht",
    )
    raw_args = sys.argv[1:]
    if not raw_args:
        raw_args = ["--batch"]
    args = parser.parse_args(raw_args)

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
            args.start_page,
            args.end_page,
            write_excel_file,
            args.excel_dir,
        )
        if skipped:
            safe_print(f"{len(skipped)} PDFs uebersprungen.")
        return

    pdf_path = args.pdf
    if pdf_path is None:
        candidates = sorted(Path("input").glob("*.pdf"))
        pdf_path = next(
            (p for p in candidates if "30.11.2025" in p.name),
            candidates[0] if candidates else None,
        )
        if pdf_path is None:
            parser.error("Keine PDF angegeben und keine Datei in ./input gefunden.")
    if not pdf_path.exists():
        parser.error(f"PDF nicht gefunden: {pdf_path}")

    rows = parse_pdf(pdf_path, start_page=args.start_page, end_page=args.end_page)
    write_csv(rows, args.out)
    safe_print(f"Geschrieben: {args.out} ({len(rows)} Zeilen)")
    if write_excel_file:
        excel_path = args.excel or args.out.with_suffix(".xlsx")
        write_excel(rows, excel_path)
        safe_print(f"Geschrieben: {excel_path} ({len(rows)} Zeilen)")


if __name__ == "__main__":
    main()
