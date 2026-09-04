"""Command-line interface for xlsheet2csv."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from . import __version__
from .converter import ConversionError, ConversionPolicy, InputLimits, convert_path

DELIMITERS = {
    "comma": ",",
    "semicolon": ";",
    "tab": "\t",
    "pipe": "|",
}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Extract XLSX worksheets to collision-free CSV files with integrity manifests."
    )
    parser.add_argument("source", type=Path, help="An .xlsx file or a directory containing .xlsx files.")
    parser.add_argument("-o", "--output-root", type=Path, help="Output root. Defaults to csv-export beside the source.")
    parser.add_argument("--recurse", action="store_true", help="Search source subdirectories recursively.")
    parser.add_argument(
        "--include",
        action="append",
        default=[],
        metavar="SHEET",
        help="Include an exact sheet name; repeat as needed. Matching is case-insensitive.",
    )
    parser.add_argument(
        "--exclude",
        action="append",
        default=[],
        metavar="SHEET",
        help="Exclude an exact sheet name; repeat as needed. Matching is case-insensitive.",
    )
    parser.add_argument(
        "--formulas",
        choices=("values", "formulas"),
        default="values",
        help="Use cached values or formula text. Formulas are never calculated.",
    )
    parser.add_argument(
        "--formula-safety",
        choices=("escape", "preserve"),
        default="escape",
        help="Escape formula-like text for spreadsheet consumers, or preserve it exactly.",
    )
    parser.add_argument("--hidden-sheets", choices=("include", "exclude"), default="include")
    parser.add_argument("--encoding", choices=("utf-8", "utf-8-sig"), default="utf-8")
    parser.add_argument("--delimiter", choices=tuple(DELIMITERS), default="comma")
    parser.add_argument("--max-file-mib", type=int, default=512)
    parser.add_argument("--max-expanded-mib", type=int, default=2048)
    parser.add_argument("--max-compression-ratio", type=float, default=250.0)
    parser.add_argument("--json", action="store_true", help="Write the run result as JSON to stdout.")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return parser


def default_output_root(source: Path) -> Path:
    resolved = source.expanduser().resolve()
    return (resolved if resolved.is_dir() else resolved.parent) / "csv-export"


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.max_file_mib <= 0 or args.max_expanded_mib <= 0 or args.max_compression_ratio <= 0:
        parser.error("Input limits must be greater than zero.")

    policy = ConversionPolicy(
        encoding=args.encoding,
        delimiter=DELIMITERS[args.delimiter],
        formulas=args.formulas,
        formula_safety=args.formula_safety,
        hidden_sheets=args.hidden_sheets,
    )
    limits = InputLimits(
        max_file_bytes=args.max_file_mib * 1024 * 1024,
        max_expanded_bytes=args.max_expanded_mib * 1024 * 1024,
        max_compression_ratio=args.max_compression_ratio,
    )
    output_root = args.output_root or default_output_root(args.source)

    try:
        manifest, exit_code, manifest_path = convert_path(
            args.source,
            output_root,
            recurse=args.recurse,
            policy=policy,
            limits=limits,
            include_sheets=args.include,
            exclude_sheets=args.exclude,
        )
    except (ConversionError, OSError, ValueError) as error:
        print(f"xlsheet2csv: {error}", file=sys.stderr)
        return 1

    if args.json:
        print(json.dumps(manifest, sort_keys=True))
    else:
        print(
            f"{manifest['status']}: {manifest['success_count']} succeeded, "
            f"{manifest['failure_count']} failed; manifest={manifest_path}"
        )
        for failure in manifest["failures"]:
            print(
                f"failed: {failure['source_relative_path']}: {failure['error_type']}: {failure['message']}",
                file=sys.stderr,
            )
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
