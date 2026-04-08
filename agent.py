#!/usr/bin/env python
"""Repository agent for Google Sheets upload tasks."""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path
from typing import Any

TYPE_RULES = [
    ("공정 진행정보", {"관리번호", "포장계획일", "진행률"}),
    ("수주내역정보", {"수주번호", "주문일자", "확정납기", "단품코드", "수주량"}),
    ("수주관리", {"대리점", "CRM 고객코드", "수주건명"}),
]


def classify_columns(columns: list[str]) -> str:
    colset = {str(col).strip() for col in columns}
    for label, required in TYPE_RULES:
        if required.issubset(colset):
            return label
    return "미분류"


def load_dataframe(path: Path) -> Any:
    try:
        import pandas as pd

        return pd.read_excel(path, dtype=object)
    except ImportError:
        raise RuntimeError(
            "Pandas is required to classify files. Install requirements with: python -m pip install -r requirements-gsheets.txt"
        )


def list_files(input_dir: Path, pattern: str) -> int:
    files = sorted(path for path in input_dir.glob(pattern) if path.is_file())
    if not files:
        print(f"No files found in {input_dir} with pattern {pattern}")
        return 0

    print(f"Found {len(files)} file(s) in {input_dir} matching '{pattern}':")
    for path in files:
        label = "unclassified"
        try:
            df = load_dataframe(path)
            label = classify_columns(list(df.columns))
        except Exception as exc:
            label = f"error: {exc}"

        print(f"- {path.name}: {label}")
    return 0


def run_upload_script(arguments: list[str]) -> int:
    command = [sys.executable, str(Path(__file__).parent / "upload_xls_to_gsheets.py")] + arguments
    result = subprocess.run(command)
    return result.returncode


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Agent for managing Google Sheets upload tasks in this repository."
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    parser_list = subparsers.add_parser("list", help="List Excel files and their inferred upload type.")
    parser_list.add_argument("--input-dir", default=".", help="Directory containing source files.")
    parser_list.add_argument("--pattern", default="*.xls", help="Glob pattern for source files.")

    parser_preview = subparsers.add_parser("preview", help="Preview the upload plan without uploading.")
    parser_preview.add_argument("--credentials", required=True, help="Path to service account JSON file.")
    parser_preview.add_argument("--spreadsheet-id", required=True, help="Target Google Spreadsheet ID.")
    parser_preview.add_argument("--input-dir", default=".", help="Directory containing source files.")
    parser_preview.add_argument("--pattern", default="*.xls", help="Glob pattern for source files.")
    parser_preview.add_argument("--uploaded-at", help="Override upload timestamp for worksheet names.")
    parser_preview.add_argument("--max-rows-per-batch", type=int, default=1000, help="Rows per Google Sheets update call.")

    parser_upload = subparsers.add_parser("upload", help="Upload Excel files to Google Sheets.")
    parser_upload.add_argument("--credentials", required=True, help="Path to service account JSON file.")
    parser_upload.add_argument("--spreadsheet-id", required=True, help="Target Google Spreadsheet ID.")
    parser_upload.add_argument("--input-dir", default=".", help="Directory containing source files.")
    parser_upload.add_argument("--pattern", default="*.xls", help="Glob pattern for source files.")
    parser_upload.add_argument("--uploaded-at", help="Override upload timestamp for worksheet names.")
    parser_upload.add_argument("--max-rows-per-batch", type=int, default=1000, help="Rows per Google Sheets update call.")

    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    if args.command == "list":
        return list_files(Path(args.input_dir), args.pattern)

    if args.command in {"preview", "upload"}:
        script_args = [
            "--credentials",
            args.credentials,
            "--spreadsheet-id",
            args.spreadsheet_id,
            "--input-dir",
            args.input_dir,
            "--pattern",
            args.pattern,
            "--max-rows-per-batch",
            str(args.max_rows_per_batch),
        ]
        if args.uploaded_at:
            script_args += ["--uploaded-at", args.uploaded_at]
        if args.command == "preview":
            script_args.append("--dry-run")

        return run_upload_script(script_args)

    parser.print_help()
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
