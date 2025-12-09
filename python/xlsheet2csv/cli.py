# xlsheet2csv - Export XLSX worksheets to CSV files.
# MIT License - see LICENSE file in the repository root for details.

import argparse
import logging
import os
import sys
from datetime import datetime

import pandas as pd


def _ensure_directory(path: str) -> None:
    if path and not os.path.exists(path):
        os.makedirs(path, exist_ok=True)


def _create_output_folder(workbook_path: str, destination_root: str, date_format: str) -> str:
    base_name = os.path.splitext(os.path.basename(workbook_path))[0]
    safe_base = "".join(c if c not in '\\/:*?"<>|' else "_" for c in base_name)
    folder_name = f"{safe_base}_{datetime.now().strftime(date_format)}"
    output_folder = os.path.join(destination_root, folder_name)
    _ensure_directory(output_folder)
    return output_folder


def _configure_logger(log_file: str | None) -> logging.Logger:
    logger = logging.getLogger("xlsheet2csv")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logger.addHandler(console_handler)

    if log_file:
        _ensure_directory(os.path.dirname(log_file))
        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        file_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
        logger.addHandler(file_handler)

    return logger


def export_workbook(workbook_path: str, output_folder: str, logger: logging.Logger, include_sheets=None, exclude_sheets=None) -> dict:
    logger.info("Starting workbook: %s", workbook_path)

    sheets = pd.read_excel(workbook_path, sheet_name=None)
    logger.info("Workbook has %d sheet(s).", len(sheets))

    include_sheets = set(include_sheets or [])
    exclude_sheets = set(exclude_sheets or [])

    exported = []
    sheet_names = []

    for name, df in sheets.items():
        if include_sheets and name not in include_sheets:
            logger.info("Skipping sheet '%s' (not in include list).", name)
            continue
        if exclude_sheets and name in exclude_sheets:
            logger.info("Skipping sheet '%s' (in exclude list).", name)
            continue

        sheet_names.append(name)
        safe_name = "".join(c if c not in '\\/:*?"<>|' else "_" for c in name)
        csv_path = os.path.join(output_folder, f"{safe_name}.csv")

        logger.info("Exporting sheet '%s' to '%s' (rows=%d, cols=%d).", name, csv_path, len(df), len(df.columns))
        df.to_csv(csv_path, index=False)
        exported.append(csv_path)

    logger.info("Completed workbook: %s (sheets exported: %d)", workbook_path, len(exported))

    return {
        "backend": "pandas",
        "workbook_path": workbook_path,
        "output_folder": output_folder,
        "sheets_exported": sheet_names,
        "csv_files": exported,
    }


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Export all sheets in an XLSX file or all XLSX files in a folder to individual CSV files."
    )
    parser.add_argument("source_path", help="Path to a single XLSX file or a folder containing XLSX files.")
    parser.add_argument(
        "-o",
        "--output-root",
        help="Destination root folder for exports. Defaults to 'csv-export' next to the source.",
    )
    parser.add_argument(
        "--backend",
        choices=["pandas"],
        default="pandas",
        help="Backend to use. Currently only 'pandas' is implemented.",
    )
    parser.add_argument(
        "--recurse",
        action="store_true",
        help="Recurse into subdirectories when source_path is a folder.",
    )
    parser.add_argument(
        "--date-format",
        default="%d-%m-%Y_%H%M",
        help="Timestamp format for export folder names (Python strftime format). Default matches dd-MM-yyyy_HHmm.",
    )
    parser.add_argument(
        "--include",
        nargs="*",
        help="List of sheet names to include. If not specified, all sheets are considered.",
    )
    parser.add_argument(
        "--exclude",
        nargs="*",
        help="List of sheet names to exclude.",
    )
    parser.add_argument(
        "--log-root",
        help="Optional root directory for log files. If not set, logs are written into each workbook's export folder.",
    )

    args = parser.parse_args()
    source_path = os.path.abspath(args.source_path)

    if not os.path.exists(source_path):
        print(f"Source path not found: {source_path}", file=sys.stderr)
        return 1

    if args.output_root:
        output_root = os.path.abspath(args.output_root)
    else:
        if os.path.isdir(source_path):
            output_root = os.path.join(source_path, "csv-export")
        else:
            output_root = os.path.join(os.path.dirname(source_path), "csv-export")

    _ensure_directory(output_root)

    if os.path.isdir(source_path):
        if args.recurse:
            xlsx_files = []
            for root, _, files in os.walk(source_path):
                for name in files:
                    if name.lower().endswith(".xlsx") and not name.startswith("~$"):
                        xlsx_files.append(os.path.join(root, name))
        else:
            xlsx_files = [
                os.path.join(source_path, name)
                for name in os.listdir(source_path)
                if name.lower().endswith(".xlsx") and not name.startswith("~$")
            ]
    else:
        xlsx_files = [source_path]

    if not xlsx_files:
        print("No .xlsx files found.", file=sys.stderr)
        return 1

    for idx, workbook_path in enumerate(sorted(xlsx_files), start=1):
        output_folder = _create_output_folder(workbook_path, output_root, args.date_format)

        if args.log_root:
            _ensure_directory(args.log_root)
            base_name = os.path.splitext(os.path.basename(workbook_path))[0]
            safe_name = "".join(c if c not in '\\/:*?"<>|' else "_" for c in base_name)
            log_file = os.path.join(args.log_root, f"{safe_name}_{datetime.now().strftime(args.date_format)}.log")
        else:
            log_file = os.path.join(output_folder, "export.log")

        logger = _configure_logger(log_file)
        logger.info("[%d/%d] Processing workbook '%s'", idx, len(xlsx_files), workbook_path)

        if args.backend == "pandas":
            export_workbook(
                workbook_path,
                output_folder,
                logger,
                include_sheets=args.include,
                exclude_sheets=args.exclude,
            )
        else:
            logger.error("Unsupported backend: %s", args.backend)
            return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
