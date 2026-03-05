"""Send raw PRN files to the Saturn Card Printer on Windows.

Usage:
    python print.py card1.prn card2.prn card3.prn
    python print.py --printer "Star Card Printer" *.prn
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

DEFAULT_PRINTER = "Star Card Printer"
SCB_MARKER = b"\x1bSCB="


def validate_prn(path: Path) -> bytes:
    """Read a PRN file and do a basic sanity check."""
    data = path.read_bytes()
    if SCB_MARKER not in data:
        raise ValueError(f"{path}: not a valid Saturn PRN file (missing SCB marker)")
    return data


def send_prn(data: bytes, printer_name: str, doc_name: str) -> None:
    """Send raw PRN bytes to a printer via the Windows spooler."""
    import win32print

    hprinter = win32print.OpenPrinter(printer_name)
    try:
        win32print.StartDocPrinter(hprinter, 1, (doc_name, None, "RAW"))
        try:
            win32print.StartPagePrinter(hprinter)
            win32print.WritePrinter(hprinter, data)
            win32print.EndPagePrinter(hprinter)
        finally:
            win32print.EndDocPrinter(hprinter)
    finally:
        win32print.ClosePrinter(hprinter)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Send PRN files to the Saturn Card Printer"
    )
    parser.add_argument("files", nargs="+", help="PRN file(s) to print")
    parser.add_argument(
        "-p", "--printer", default=DEFAULT_PRINTER,
        help=f'Printer name (default: "{DEFAULT_PRINTER}")',
    )
    args = parser.parse_args()

    paths = [Path(f) for f in args.files]

    # Validate all files before sending any
    jobs: list[tuple[Path, bytes]] = []
    for path in paths:
        if not path.exists():
            print(f"Error: {path} not found", file=sys.stderr)
            sys.exit(1)
        try:
            data = validate_prn(path)
        except ValueError as e:
            print(f"Error: {e}", file=sys.stderr)
            sys.exit(1)
        jobs.append((path, data))

    total = len(jobs)
    print(f"Sending {total} job(s) to \"{args.printer}\"")

    for i, (path, data) in enumerate(jobs, 1):
        print(f"  [{i}/{total}] {path.name} ({len(data):,} bytes) ... ", end="", flush=True)
        try:
            send_prn(data, args.printer, path.name)
            print("done")
        except Exception as e:
            print(f"FAILED: {e}", file=sys.stderr)
            sys.exit(1)

    print(f"All {total} job(s) sent successfully.")


if __name__ == "__main__":
    main()
