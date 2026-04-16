import argparse
import math
import random
import re
import time
from datetime import datetime, timedelta
from pathlib import Path

import gspread
import pandas as pd
from google.oauth2.service_account import Credentials


SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SPREADSHEET_CELL_LIMIT = 10_000_000
IMMUTABLE_WORKSHEET_TITLES = {"북미키워드"}

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
    def positive_int(value: str) -> int:
        parsed = int(value)
        if parsed <= 0:
            raise argparse.ArgumentTypeError("must be a positive integer")
        return parsed

    def non_negative_int(value: str) -> int:
        parsed = int(value)
        if parsed < 0:
            raise argparse.ArgumentTypeError("must be a non-negative integer")
        return parsed

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
        type=positive_int,
        default=3000,
        help="How many rows to upload per API call.",
    )
    parser.add_argument(
        "--max-write-requests-per-minute",
        type=positive_int,
        default=35,
        help="Throttle Google Sheets write calls to this per-minute rate for stability.",
    )
    parser.add_argument(
        "--max-write-retries",
        type=non_negative_int,
        default=8,
        help="How many times to retry a write chunk on transient API errors (429/5xx).",
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
    parser.add_argument(
        "--skip-read-errors",
        action="store_true",
        help="Skip files that fail to read and continue with the rest.",
    )
    parser.add_argument(
        "--cleanup-protect-days",
        type=non_negative_int,
        default=0,
        help=(
            "Protect worksheets from the most recent N days during automatic capacity cleanup. "
            "Default is 0 (no additional day protection)."
        ),
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

    # Inventory sample format fallback:
    # e.g. 품목코드/색상/재고구분/현재고 + daily forecast columns.
    inventory_sample_required = {
        normalize_col_name("품목코드"),
        normalize_col_name("색상"),
        normalize_col_name("재고구분"),
        normalize_col_name("현재고"),
    }
    if inventory_sample_required.issubset(colset):
        return "재고현황"

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


def is_retryable_api_error(exc: gspread.APIError) -> bool:
    status_code = getattr(getattr(exc, "response", None), "status_code", None)
    if status_code in {429, 500, 502, 503, 504}:
        return True
    message = str(exc).lower()
    return "quota" in message or "rate" in message or "timeout" in message


def upload_values(
    worksheet,
    values: list[list[str]],
    max_rows_per_batch: int,
    max_write_requests_per_minute: int,
    max_write_retries: int,
) -> None:
    write_interval_seconds = 60.0 / max(max_write_requests_per_minute, 1)
    next_allowed_ts = 0.0

    for start in range(0, len(values), max_rows_per_batch):
        chunk = values[start : start + max_rows_per_batch]
        cell = f"A{start + 1}"
        attempt = 0
        while True:
            now = time.monotonic()
            if now < next_allowed_ts:
                time.sleep(next_allowed_ts - now)

            try:
                worksheet.update(range_name=cell, values=chunk)
                next_allowed_ts = time.monotonic() + write_interval_seconds
                break
            except gspread.APIError as exc:
                if (not is_retryable_api_error(exc)) or attempt >= max_write_retries:
                    raise
                status_code = getattr(getattr(exc, "response", None), "status_code", "unknown")
                backoff = min(90.0, (2 ** attempt)) + random.uniform(0.2, 0.8)
                print(
                    f"  write retry -> row={start + 1} status={status_code} "
                    f"attempt={attempt + 1}/{max_write_retries} wait={backoff:.1f}s"
                )
                time.sleep(backoff)
                attempt += 1


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


def run_daily_cleanup(spreadsheet, apply: bool, heading: str) -> None:
    keep, delete = plan_daily_cleanup(spreadsheet)
    delete = [item for item in delete if item["title"] not in IMMUTABLE_WORKSHEET_TITLES]
    print(f"\n{heading}")
    for item in sorted(keep, key=lambda x: (x["label"], x["day"])):
        print(f"  keep   -> {item['title']}")
    for item in sorted(delete, key=lambda x: (x["label"], x["day"], x["hhmm"])):
        print(f"  delete -> {item['title']}")

    if apply:
        for item in delete:
            spreadsheet.del_worksheet(item["worksheet"])
            print(f"  deleted -> {item['title']}")
    else:
        print("  (preview only) add --cleanup-apply to delete listed worksheets.")


def worksheet_cell_count(worksheet) -> int:
    return max(int(worksheet.row_count), 1) * max(int(worksheet.col_count), 1)


def required_cells_for_job(job: dict) -> int:
    required_rows = max(int(job["rows"]) + 1, 1)  # +1 for header row
    required_cols = max(int(job["cols"]), 1)
    return required_rows * required_cols


def projected_cell_growth(spreadsheet, jobs: list[dict]) -> int:
    worksheets_by_title = {ws.title: ws for ws in spreadsheet.worksheets()}
    growth = 0
    for job in jobs:
        title = job["worksheet_title"]
        required = required_cells_for_job(job)
        existing = worksheets_by_title.get(title)
        if existing is None:
            growth += required
            continue
        growth += max(required - worksheet_cell_count(existing), 0)
    return growth


def free_cells_for_upload(
    spreadsheet,
    jobs: list[dict],
    protect_days: int = 0,
    cell_limit: int = SPREADSHEET_CELL_LIMIT,
) -> None:
    total_cells = sum(worksheet_cell_count(ws) for ws in spreadsheet.worksheets())
    growth = projected_cell_growth(spreadsheet, jobs)
    projected = total_cells + growth
    if projected <= cell_limit:
        return

    need_to_free = projected - cell_limit
    known_labels = {label for label, _ in TYPE_RULES}
    protected_titles = {job["worksheet_title"] for job in jobs}
    protected_titles.update(IMMUTABLE_WORKSHEET_TITLES)
    # Keep at least one latest worksheet per known label.
    latest_title_per_label: dict[str, str] = {}
    for ws in spreadsheet.worksheets():
        parsed = parse_worksheet_stamp(ws.title)
        if not parsed:
            continue
        label, day, hhmm = parsed
        if label not in known_labels:
            continue
        stamp = f"{day}_{hhmm}"
        existing_title = latest_title_per_label.get(label)
        if existing_title is None:
            latest_title_per_label[label] = ws.title
            continue
        existing_parsed = parse_worksheet_stamp(existing_title)
        if existing_parsed is None:
            latest_title_per_label[label] = ws.title
            continue
        existing_stamp = f"{existing_parsed[1]}_{existing_parsed[2]}"
        if stamp > existing_stamp:
            latest_title_per_label[label] = ws.title
    protected_titles.update(latest_title_per_label.values())

    protected_day_from = (datetime.now().date() - timedelta(days=protect_days)) if protect_days > 0 else None
    candidates = []
    for ws in spreadsheet.worksheets():
        parsed = parse_worksheet_stamp(ws.title)
        if not parsed:
            continue
        label, day, hhmm = parsed
        if label not in known_labels:
            continue
        if ws.title in protected_titles:
            continue
        if protected_day_from is not None:
            sheet_day = datetime.strptime(day, "%Y-%m-%d").date()
            if sheet_day >= protected_day_from:
                continue
        candidates.append({"worksheet": ws, "title": ws.title, "day": day, "hhmm": hhmm, "cells": worksheet_cell_count(ws)})

    candidates.sort(key=lambda x: (x["day"], x["hhmm"], x["title"]))
    if not candidates:
        raise SystemExit(
            f"Insufficient spreadsheet cell capacity: need to free at least {need_to_free} cells, "
            "but no managed worksheets are available to delete."
        )

    print(
        "\nauto capacity cleanup before upload:"
        f" current_cells={total_cells} projected_growth={growth} limit={cell_limit} protect_days={protect_days}"
    )
    freed = 0
    for item in candidates:
        spreadsheet.del_worksheet(item["worksheet"])
        freed += item["cells"]
        print(f"  deleted(capacity) -> {item['title']} cells={item['cells']}")
        if freed >= need_to_free:
            break

    if freed < need_to_free:
        raise SystemExit(
            f"Insufficient spreadsheet cell capacity after cleanup: "
            f"needed={need_to_free}, freed={freed}. "
            f"Delete more old worksheets, lower --cleanup-protect-days (current={protect_days}), or use a new spreadsheet."
        )


def read_excel_file(path: Path) -> list[tuple[str, pd.DataFrame]]:
    workbook = pd.read_excel(path, sheet_name=None, dtype=object)
    if not isinstance(workbook, dict):
        return [("Sheet1", workbook)]
    return [(str(sheet_name), sheet_df) for sheet_name, sheet_df in workbook.items()]


def preprocess_dataframe(df: pd.DataFrame, label: str) -> pd.DataFrame:
    processed = df.copy()

    # The production progress export uses visually merged cells,
    # so repeated values must be filled from the previous row.
    if label == "공정 진행정보":
        processed = processed.ffill()

    # Inventory sample headers can differ from dashboard join keys.
    # Normalize key columns so dashboard code can reuse the uploaded sheet directly.
    if label == "재고현황":
        normalized_col_map = {normalize_col_name(col): col for col in processed.columns}
        rename_map = {}
        alias_candidates = [
            ("품목코드", "단품코드"),
            ("(기간)총입고예정", "기간총입고"),
            ("(기간)총출고예정", "기간총출고"),
        ]
        for src_alias, target_name in alias_candidates:
            src_norm = normalize_col_name(src_alias)
            if src_norm in normalized_col_map:
                rename_map[normalized_col_map[src_norm]] = target_name
        if rename_map:
            processed = processed.rename(columns=rename_map)

    return processed


def same_columns_in_order(left: pd.DataFrame, right: pd.DataFrame) -> bool:
    return list(left.columns) == list(right.columns)


def main() -> None:
    args = parse_args()
    input_dir = Path(args.input_dir)
    files = discover_files(input_dir, args.pattern)

    if not files:
        raise SystemExit(f"No files found in {input_dir} with pattern {args.pattern}")

    uploaded_at = args.uploaded_at or datetime.now().strftime("%Y-%m-%d_%H%M")

    raw_jobs = []
    read_errors = []
    for path in files:
        try:
            sheets = read_excel_file(path)
        except Exception as exc:
            if args.skip_read_errors:
                read_errors.append((path, exc))
                continue
            raise SystemExit(f"Failed to read '{path.name}': {exc}") from exc
        for sheet_name, raw_df in sheets:
            if raw_df is None or len(raw_df.columns) == 0:
                continue
            label = classify_columns(list(raw_df.columns))
            df = preprocess_dataframe(raw_df, label)
            raw_jobs.append(
                {
                    "path": path,
                    "sheet_name": sheet_name,
                    "source_name": f"{path.name}#{sheet_name}",
                    "label": label,
                    "rows": len(df),
                    "cols": len(df.columns),
                    "df": df,
                }
            )

    if not raw_jobs:
        raise SystemExit("No readable files to upload.")

    # Merge files by label when their header columns are identical in the same order.
    jobs_by_label = {}
    for job in raw_jobs:
        label = job["label"]
        if label not in jobs_by_label:
            jobs_by_label[label] = {
                "label": label,
                "paths": [job["path"]],
                "source_names": [job["source_name"]],
                "df": job["df"],
            }
            continue

        existing = jobs_by_label[label]
        existing_df = existing["df"]
        incoming_df = job["df"]
        if not same_columns_in_order(existing_df, incoming_df):
            raise SystemExit(
                f"Cannot merge files for label '{label}' because column headers differ. "
                f"first={existing['source_names'][0]}, second={job['source_name']}"
            )
        existing["paths"].append(job["path"])
        existing["source_names"].append(job["source_name"])
        existing["df"] = pd.concat([existing_df, incoming_df], ignore_index=True)

    jobs = []
    for label, merged in sorted(jobs_by_label.items(), key=lambda x: x[0]):
        merged_df = merged["df"]
        worksheet_title = sanitize_worksheet_title(f"{label}_{uploaded_at}")
        jobs.append(
            {
                "path": merged["paths"][0],
                "source_paths": merged["paths"],
                "source_names": merged["source_names"],
                "label": label,
                "rows": len(merged_df),
                "cols": len(merged_df.columns),
                "worksheet_title": worksheet_title,
                "df": merged_df,
            }
        )

    seen_titles = set()
    duplicate_titles = set()
    for job in jobs:
        title = job["worksheet_title"]
        if title in seen_titles:
            duplicate_titles.add(title)
        seen_titles.add(title)
    if duplicate_titles:
        duplicates_text = ", ".join(sorted(duplicate_titles))
        raise SystemExit(
            "Detected duplicate worksheet titles in this run. "
            f"Adjust --uploaded-at or inputs and retry. titles={duplicates_text}"
        )

    print("upload plan:")
    for job in jobs:
        source_names = ",".join(job["source_names"])
        print(
            f"- files={source_names} "
            f"type={job['label']} "
            f"rows={job['rows']} cols={job['cols']} "
            f"worksheet={job['worksheet_title']}"
        )
    for path, exc in read_errors:
        print(f"- skipped file={path.name} reason={exc}")

    if args.dry_run and not args.cleanup_daily:
        return

    client = authorize(args.credentials)
    spreadsheet = client.open_by_key(args.spreadsheet_id)

    # Free up cells first when cleanup deletion is requested.
    # This avoids hitting the 10M-cell spreadsheet limit during new sheet creation.
    if args.cleanup_daily and args.cleanup_apply:
        run_daily_cleanup(
            spreadsheet=spreadsheet,
            apply=True,
            heading="cleanup plan before upload (free cells first):",
        )
        free_cells_for_upload(
            spreadsheet=spreadsheet,
            jobs=jobs,
            protect_days=args.cleanup_protect_days,
        )

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
            upload_values(
                worksheet=worksheet,
                values=values,
                max_rows_per_batch=args.max_rows_per_batch,
                max_write_requests_per_minute=args.max_write_requests_per_minute,
                max_write_retries=args.max_write_retries,
            )
            print(f"  uploaded -> {job['worksheet_title']}")

    if args.cleanup_daily:
        run_daily_cleanup(
            spreadsheet=spreadsheet,
            apply=args.cleanup_apply,
            heading="cleanup plan after upload (keep latest per type/day):",
        )

    print(f"done: {spreadsheet.url}")


if __name__ == "__main__":
    main()
