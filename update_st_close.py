import datetime as dt
import argparse
import os
import sys
import time
from pathlib import Path
from typing import Dict, Optional, Tuple

import akshare as ak  # type: ignore
import pandas as pd


def ensure_std_streams() -> None:
    """
    In PyInstaller --windowed mode, stdout/stderr can be None.
    Some third-party libs (e.g. tqdm in AKShare) require writable streams.
    """
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w")


ensure_std_streams()


def get_effective_trade_date_and_source() -> Tuple[dt.date, str]:
    """
    Determine the trade date and which price field to use from AKShare.

    Rules:
    - If today is NOT a trading day:
        - Use the latest past trading day as trade_date
        - Use 'latest' price from实时行情 (等价于该最近交易日收盘价)
    - If today IS a trading day:
        - If current time >= 15:00 (认为已经收盘):
            - trade_date = today
            - use 'latest' price
        - Else (盘中):
            - trade_date = previous trading day
            - use 'prev_close' (昨收)
    """
    today = dt.date.today()

    try:
        trade_dates_df = ak.tool_trade_date_hist_sina()
    finally:
        # Avoid hitting remote endpoints too frequently.
        time.sleep(3)
    trade_dates = pd.to_datetime(trade_dates_df["trade_date"]).dt.date

    past_trade_dates = [d for d in trade_dates if d <= today]
    if not past_trade_dates:
        raise RuntimeError("No past trading dates available from AKShare.")

    last_trade_date = past_trade_dates[-1]

    # Today is not a trading day: use the most recent trading day, and latest price
    if last_trade_date != today:
        return last_trade_date, "latest"

    # Today is a trading day
    now_time = dt.datetime.now().time()
    close_time = dt.time(15, 0)

    if now_time >= close_time:
        # After market close: use today's latest price (收盘价)
        return today, "latest"

    # Market not closed yet -> use yesterday's收盘价
    if len(past_trade_dates) >= 2:
        return past_trade_dates[-2], "prev_close"

    # Edge case: only one trading day
    return today, "latest"


def normalize_code(value) -> Optional[str]:
    """
    Normalize a '代码' cell value from Excel into a 6-digit A-share code string.
    """
    if pd.isna(value):
        return None

    s = str(value).strip()
    if not s:
        return None

    # Remove any decimal part (e.g. '600000.0' -> '600000')
    if "." in s:
        s = s.split(".")[0]

    # Keep only digits
    digits = "".join(ch for ch in s if ch.isdigit())
    if not digits:
        return None

    # If longer than 6 digits, keep the last 6
    if len(digits) > 6:
        digits = digits[-6:]

    return digits.zfill(6)


def normalize_header_row(df: pd.DataFrame) -> pd.DataFrame:
    """
    Some sheets (like sheet 3) have the real header not in the first row.
    Try to find a row that contains '代码' and promote it to header.
    """
    normalized_columns = [str(c).strip() for c in df.columns]
    if "代码" in normalized_columns:
        df = df.copy()
        df.columns = normalized_columns
        return df

    for idx in range(len(df)):
        row = df.iloc[idx]
        if any("代码" in str(v).strip() for v in row):
            # Use this row as header
            new_columns = [str(v).strip() if not pd.isna(v) else "" for v in row]
            new_df = df.iloc[idx + 1 :].copy()
            new_df.columns = new_columns
            new_df.reset_index(drop=True, inplace=True)
            return new_df

    return df


def find_code_column(df: pd.DataFrame) -> Optional[str]:
    """
    Find the best matching column name for stock code.
    """
    exact = [c for c in df.columns if str(c).strip() == "代码"]
    if exact:
        return exact[0]

    contains = [c for c in df.columns if "代码" in str(c).strip()]
    if contains:
        return contains[0]

    return None


def is_st_name(value) -> bool:
    """
    Determine whether a stock name indicates ST by checking if it contains 'st'.
    """
    if pd.isna(value):
        return False
    return "st" in str(value).strip().lower()


def find_name_column(df: pd.DataFrame, code_col: str) -> Optional[str]:
    """
    Prefer the column immediately to the right of code_col; fallback to columns containing '名称'.
    """
    columns = list(df.columns)
    try:
        idx = columns.index(code_col)
    except ValueError:
        idx = -1

    if idx >= 0 and idx + 1 < len(columns):
        return columns[idx + 1]

    exact = [c for c in columns if str(c).strip() == "名称"]
    if exact:
        return exact[0]

    contains = [c for c in columns if "名称" in str(c).strip()]
    if contains:
        return contains[0]

    return None


def build_close_price_cache(price_source: str) -> Dict[str, Dict[str, Optional[object]]]:
    """
    Fetch all A-share latest data once, and build {code: close_price} cache.

    price_source:
      - 'latest'    -> use columns like '最新价' / '今收'
      - 'prev_close'-> use columns like '昨收'
    """
    print(f"Fetching all A-share quotes via AKShare Sina API, source={price_source!r} ...")
    try:
        spot = ak.stock_zh_a_spot()
    except Exception as e:  # noqa: BLE001
        raise RuntimeError(f"Failed to fetch spot data from AKShare: {e}") from e
    finally:
        # Force a cooldown after each AKShare call.
        time.sleep(3)

    if spot is None or spot.empty:
        raise RuntimeError("Empty spot data from AKShare.")

    if "代码" not in spot.columns:
        raise RuntimeError("Column '代码' not found in spot data.")

    # Sina spot data mainly exposes 最新价 / 昨收.
    price_col_candidates_latest = ["最新价", "今收", "收盘", "现价"]
    price_col_candidates_prev = ["昨收", "前收盘", "前收盘价"]

    if price_source == "latest":
        candidates = price_col_candidates_latest
    else:
        candidates = price_col_candidates_prev

    price_col: Optional[str] = None
    for col in candidates:
        if col in spot.columns:
            price_col = col
            break

    if price_col is None:
        raise RuntimeError(
            f"None of price columns {candidates} found in spot data. "
            f"Available columns: {list(spot.columns)}"
        )

    cache: Dict[str, Dict[str, Optional[object]]] = {}
    spot_name_col = "名称" if "名称" in spot.columns else None
    use_cols = ["代码", price_col] + ([spot_name_col] if spot_name_col else [])

    for _, row in spot[use_cols].iterrows():
        code_raw = row["代码"]
        price_raw = row[price_col]
        name_raw = row[spot_name_col] if spot_name_col else None

        code = normalize_code(code_raw)
        if code is None:
            continue

        st_flag = is_st_name(name_raw)

        if pd.isna(price_raw):
            cache[code] = {"close_price": None, "is_st": st_flag}
            continue

        try:
            cache[code] = {"close_price": float(price_raw), "is_st": st_flag}
        except Exception:  # noqa: BLE001
            cache[code] = {"close_price": None, "is_st": st_flag}

    if not cache:
        raise RuntimeError("Built empty close price cache from spot data.")

    print(f"Built close price cache for {len(cache)} codes.")
    return cache


def process_excel_file(excel_path: Path, output_path: Optional[Path] = None) -> Path:
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    print(f"Loading Excel: {excel_path.resolve()}")

    trade_date, price_source = get_effective_trade_date_and_source()
    trade_date_str = trade_date.isoformat()

    print(f"Effective trade date: {trade_date_str}, price_source={price_source}")

    # Step 0: fetch all close prices into memory once
    close_cache: Dict[str, Dict[str, Optional[object]]] = build_close_price_cache(price_source)

    # Read all sheets into a dict: {sheet_name: DataFrame}
    sheets: Dict[str, pd.DataFrame] = pd.read_excel(
        excel_path, sheet_name=None
    )

    updated_sheets: Dict[str, pd.DataFrame] = {}

    for sheet_name, df in sheets.items():
        print(f"\nProcessing sheet: {sheet_name}")
        df = normalize_header_row(df)

        code_col = find_code_column(df)
        if code_col is None:
            print("  Column '代码' not found (even after header normalization), skipping this sheet.")
            updated_sheets[sheet_name] = df
            continue

        name_col = find_name_column(df, code_col)
        codes_series = df[code_col]
        name_series = df[name_col] if name_col is not None else pd.Series([None] * len(df))
        norm_codes = codes_series.map(normalize_code)

        new_col_values = []
        note_values = []
        for file_name, norm in zip(name_series, norm_codes):
            if norm is None:
                new_col_values.append("-")
                note_values.append("")
                continue

            if norm not in close_cache:
                # Price not found in pre-fetched cache
                new_col_values.append("-")
                note_values.append("")
            else:
                quote = close_cache[norm]
                price = quote.get("close_price")
                new_col_values.append(price if price is not None else "-")

                sina_is_st = bool(quote.get("is_st"))
                file_is_st = is_st_name(file_name)
                if file_is_st != sina_is_st:
                    # File has ST but Sina does not -> remove ST; otherwise add ST.
                    note_values.append("摘帽" if file_is_st else "戴帽")
                else:
                    note_values.append("")

        new_col_name = f"{trade_date_str} 收盘价"
        note_col_name = f"{trade_date_str} 说明"
        df[new_col_name] = new_col_values
        df[note_col_name] = note_values

        updated_sheets[sheet_name] = df
        print(f"  Added columns: {new_col_name}, {note_col_name}")

    # Save to a new Excel file to avoid overwriting the original
    if output_path is None:
        output_path = excel_path.with_name(f"{excel_path.stem}_updated.xlsx")
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in updated_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"\nDone. Updated Excel saved to: {output_path.resolve()}")
    print(f"Trade date used: {trade_date_str}")
    return output_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Update close price columns in ST Excel files.")
    parser.add_argument(
        "-i",
        "--input",
        default="st.xlsx",
        help="Input Excel path. Default: st.xlsx",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="Output Excel path. Default: <input_stem>_updated.xlsx",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_path = Path(args.input)
    output_path = Path(args.output) if args.output else None
    process_excel_file(input_path, output_path)


if __name__ == "__main__":
    main()
