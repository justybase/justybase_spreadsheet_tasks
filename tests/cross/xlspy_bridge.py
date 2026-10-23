#!/usr/bin/env python3
"""JSON bridge between the Node cross-test suite and the xlspy Python library.

Commands (JSON job on stdin, JSON result on stdout):

  check
      Verify xlspy is importable. No job. Prints {"ok": true}.

  write --format xlsx|xlsb --file <path>
      Job: {"sheets": [{"name": str, "rows": [[cell, ...], ...],
                        "tables": [{"displayName": str, "ref": str,
                                    "columns": [str, ...] | null}] | null}],
            "headerRows": int}
      Creates the file with xlspy's writers.

  read --file <path> [--streaming]
      Job: {"streaming": bool} (optional)
      Reads every sheet with xlspy's ExcelReader.

  replace-sheet-data --format xlsx|xlsb --file <src> --sheet <name> --out <dst>
      Job: {"rows": [[cell, ...], ...], "headers": [str, ...] | null}
      Runs xlspy's XlsxUpdater/XlsbUpdater, saves to --out, then re-reads the
      output so the caller sees the result through xlspy's own reader.

Cell encoding: {"datetime": "<ISO 8601, naive>"} for date cells (same
convention as xlspy's golden JSON sidecars); everything else is JSON-native.
"""
import argparse
import datetime as dt
import json
import sys
import traceback


def decode_cell(value):
    if isinstance(value, dict) and "datetime" in value:
        return dt.datetime.fromisoformat(value["datetime"])
    return value


def decode_rows(rows):
    return [[decode_cell(cell) for cell in row] for row in rows]


def encode_cell(value):
    if isinstance(value, dt.datetime):
        return {"datetime": value.isoformat(timespec="seconds")}
    if isinstance(value, dt.date):
        return {"datetime": dt.datetime(value.year, value.month, value.day).isoformat(timespec="seconds")}
    return value


def read_sheets(path, streaming=False):
    from xlspy import ExcelReader

    with ExcelReader(path) as reader:
        return {
            name: [[encode_cell(cell) for cell in row] for row in reader.read_all(name, streaming=streaming)]
            for name in reader.get_sheet_names()
        }


def cmd_write(args, job):
    from xlspy import ExcelTable, XlsbWriter, XlsxWriter

    cls = XlsxWriter if args.format == "xlsx" else XlsbWriter
    with cls(args.file) as writer:
        header_rows = job.get("headerRows", 1)
        for sheet in job["sheets"]:
            writer.add_sheet(sheet["name"])
            tables = None
            if sheet.get("tables"):
                tables = [
                    ExcelTable(t["displayName"], t["ref"], columns=t.get("columns"))
                    for t in sheet["tables"]
                ]
            writer.write_sheet(decode_rows(sheet["rows"]), tables=tables, header_rows=header_rows)
    return {"ok": True}


def cmd_read(args, job):
    streaming = bool(job.get("streaming", False)) if job else False
    return {"sheets": read_sheets(args.file, streaming=streaming)}


def cmd_replace(args, job):
    from xlspy import XlsbUpdater, XlsxUpdater

    cls = XlsxUpdater if args.format == "xlsx" else XlsbUpdater
    updater = cls(args.file)
    try:
        updater.replace_sheet_data(args.sheet, decode_rows(job["rows"]), headers=job.get("headers"))
        updater.save(args.out)
    finally:
        updater.close()
    return {"sheets": read_sheets(args.out)}


def cmd_check(_args, _job):
    import xlspy  # noqa: F401

    return {"ok": True}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("check")
    p.set_defaults(func=cmd_check)

    p = sub.add_parser("write")
    p.add_argument("--format", choices=("xlsx", "xlsb"), required=True)
    p.add_argument("--file", required=True)
    p.set_defaults(func=cmd_write)

    p = sub.add_parser("read")
    p.add_argument("--file", required=True)
    p.set_defaults(func=cmd_read)

    p = sub.add_parser("replace-sheet-data")
    p.add_argument("--format", choices=("xlsx", "xlsb"), required=True)
    p.add_argument("--file", required=True)
    p.add_argument("--sheet", required=True)
    p.add_argument("--out", required=True)
    p.set_defaults(func=cmd_replace)

    args = parser.parse_args()
    stdin_text = sys.stdin.read()
    try:
        job = json.loads(stdin_text) if stdin_text.strip() else {}
    except json.JSONDecodeError:
        job = {}
    try:
        result = args.func(args, job)
    except Exception:  # noqa: BLE001 - report to the caller, exit non-zero
        traceback.print_exc(file=sys.stderr)
        return 1
    json.dump(result, sys.stdout)
    return 0


if __name__ == "__main__":
    sys.exit(main())
