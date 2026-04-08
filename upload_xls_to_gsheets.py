import argparse
import math
import re
from datetime import datetime
from pathlib import Path

import gspread
import pandas as pd
from google.oauth2.service_account import Credentials


SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

TYPE_RULES = [
    (
        "재고현황",
        {"제품구분", "단품코드", "현재고", "재고금액"},
    ),
    (
        "공정 진행정보",
        {"관리번호", "포장계획일", "진행률"},
    ),
    (
        "수주내역정보",
        {"수주번호", "주문일자", "확정납기", "단품코드", "수주량"},
    ),
    (
        "수주관리",
        {"대리점", "CRM 고객코드", "수주건명"},
    ),
]


def normalize_col_name(name: str) -> str:
    text = str(name).strip()
    # Grid exports sometimes include sort markers like '확정납기▼'
    text = re.sub(r"[▲▼]", "", text)
    text = re.sub(r"\s+", "", text)
    return text


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Upload local Excel files to Google Sheets with timestamped worksheet names."
    )
    parser.add_argument("--credentials", required=True, help="Path to service account JSON.")
    parser.add_argument("--spreadsheet-id", required=True, help="Target Google Spreadsheet ID.")
    parser.add_argument(
        "--input-dir",
        default=".",
        help="Directory containing source files. Defaults to current directory.",
    )
    parser.add_argument(
        "--pattern",
        default="*.xls",
        help="Glob pattern for source files. Defaults to *.xls",
    )
    parser.add_argument(
        "--uploaded-at",
        help="Override upload timestamp for worksheet names. Example: 2026-03-26_1745",
    )
    parser.add_argument(
        "--max-rows-per-batch",
        type=int,
        default=1000,
        help="How many rows to upload per API call.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print classification and row counts without uploading.",
    )
    parser.add_argument(
        "--cleanup-daily",
        action="store_true",
        help="Keep only one worksheet per (type, date). The latest HHMM is kept.",
    )
    parser.add_argument(
        "--cleanup-apply",
        action="store_true",
        help="Actually delete old worksheets when --cleanup-daily is set.",
    )
    return parser.parse_args()


def authorize(credentials_path: str) -> gspread.Client:
    creds = Credentials.from_service_account_file(credentials_path, scopes=SCOPES)
    return gspread.authorize(creds)


def discover_files(input_dir: Path, pattern: str) -> list[Path]:
    return sorted(path for path in input_dir.glob(pattern) if path.is_file())


def classify_columns(columns: list[str]) -> str:
    colset = {normalize_col_name(col) for col in columns}
    for label, required in TYPE_RULES:
        normalized_required = {normalize_col_name(col) for col in required}
        if normalized_required.issubset(colset):
            return label
    return "미분류"


def sanitize_worksheet_title(name: str) -> str:
    cleaned = re.sub(r"[\[\]\:\*\?\/\\]", "_", name).strip()
    return cleaned[:100] or "Sheet"


def parse_worksheet_stamp(title: str):
    match = re.match(r"^(?P<label>.+)_(?P<date>\d{4}-\d{2}-\d{2})_(?P<time>\d{4})$", title)
    if not match:
        return None
    return match.group("label"), match.group("date"), match.group("time")


def stringify(value) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, pd.Timestamp):
        return value.isoformat(sep=" ")
    if isinstance(value, float) and math.isnan(value):
        return ""
    return str(value)


def dataframe_to_values(df: pd.DataFrame) -> list[list[str]]:
    normalized = df.astype(object).where(pd.notna(df), "")
    header = [stringify(col) for col in normalized.columns]
    rows = [[stringify(cell) for cell in row] for row in normalized.itertuples(index=False, name=None)]
    return [header, *rows]


def get_or_create_worksheet(spreadsheet, title: str, rows: int, cols: int):
    try:
        worksheet = spreadsheet.worksheet(title)
        worksheet.clear()
        worksheet.resize(rows=max(rows, 1), cols=max(cols, 1))
        return worksheet
    except gspread.WorksheetNotFound:
        return spreadsheet.add_worksheet(title=title, rows=max(rows, 1), cols=max(cols, 1))


def upload_values(worksheet, values: list[list[str]], max_rows_per_batch: int) -> None:
    for start in range(0, len(values), max_rows_per_batch):
        chunk = values[start : start + max_rows_per_batch]
        cell = f"A{start + 1}"
        worksheet.update(range_name=cell, values=chunk)


def plan_daily_cleanup(spreadsheet):
    candidates = []
    known_labels = {label for label, _ in TYPE_RULES}
    for ws in spreadsheet.worksheets():
        parsed = parse_worksheet_stamp(ws.title)
        if not parsed:
            continue
        label, day, hhmm = parsed
        if label not in known_labels:
            continue
        candidates.append({"worksheet": ws, "title": ws.title, "label": label, "day": day, "hhmm": hhmm})

    grouped = {}
    for item in candidates:
        grouped.setdefault((item["label"], item["day"]), []).append(item)

    keep = []
    delete = []
    for _, items in grouped.items():
        sorted_items = sorted(items, key=lambda x: x["hhmm"], reverse=True)
        keep.append(sorted_items[0])
        delete.extend(sorted_items[1:])
    return keep, delete


def read_excel_file(path: Path) -> pd.DataFrame:
    return pd.read_excel(path, dtype=object)


def preprocess_dataframe(df: pd.DataFrame, label: str) -> pd.DataFrame:
    processed = df.copy()

    # The production progress export uses visually merged cells,
    # so repeated values must be filled from the previous row.
    if label == "공정 진행정보":
        processed = processed.ffill()

    return processed


def main() -> None:
    args = parse_args()
    input_dir = Path(args.input_dir)
    files = discover_files(input_dir, args.pattern)

    if not files:
        raise SystemExit(f"No files found in {input_dir} with pattern {args.pattern}")

    uploaded_at = args.uploaded_at or datetime.now().strftime("%Y-%m-%d_%H%M")

    jobs = []
    for path in files:
        raw_df = read_excel_file(path)
        label = classify_columns(list(raw_df.columns))
        df = preprocess_dataframe(raw_df, label)
        worksheet_title = sanitize_worksheet_title(f"{label}_{uploaded_at}")
        jobs.append(
            {
                "path": path,
                "label": label,
                "rows": len(df),
                "cols": len(df.columns),
                "worksheet_title": worksheet_title,
                "df": df,
            }
        )

    print("upload plan:")
    for job in jobs:
        print(
            f"- file={job['path'].name} "
            f"type={job['label']} "
            f"rows={job['rows']} cols={job['cols']} "
            f"worksheet={job['worksheet_title']}"
        )

    if args.dry_run and not args.cleanup_daily:
        return

    client = authorize(args.credentials)
    spreadsheet = client.open_by_key(args.spreadsheet_id)

    if not args.dry_run:
        for job in jobs:
            values = dataframe_to_values(job["df"])
            rows = len(values)
            cols = max((len(row) for row in values), default=1)
            worksheet = get_or_create_worksheet(
                spreadsheet=spreadsheet,
                title=job["worksheet_title"],
                rows=rows,
                cols=cols,
            )
            upload_values(worksheet, values, args.max_rows_per_batch)
            print(f"  uploaded -> {job['worksheet_title']}")

    if args.cleanup_daily:
        keep, delete = plan_daily_cleanup(spreadsheet)
        print("\ncleanup plan (keep latest per type/day):")
        for item in sorted(keep, key=lambda x: (x["label"], x["day"])):
            print(f"  keep   -> {item['title']}")
        for item in sorted(delete, key=lambda x: (x["label"], x["day"], x["hhmm"])):
            print(f"  delete -> {item['title']}")

        if args.cleanup_apply:
            for item in delete:
                spreadsheet.del_worksheet(item["worksheet"])
                print(f"  deleted -> {item['title']}")
        else:
            print("  (preview only) add --cleanup-apply to delete listed worksheets.")

    print(f"done: {spreadsheet.url}")


if __name__ == "__main__":
    main()
