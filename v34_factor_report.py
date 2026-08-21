"""Experimental V3.4 factor observer for the Google Sheet Watchlist.

This script intentionally does not emit BUY/SELL signals.  V3.4 currently has
the point-in-time data layer and 29 factors, but not the finished regime,
backtest, strategy-weighting, or runtime layers.  The output is therefore a
research snapshot: seven equal-weight style averages plus the legacy score.
"""

from __future__ import annotations

import argparse
import json
import math
import os
from datetime import datetime
from statistics import fmean
from typing import Any
from zoneinfo import ZoneInfo

import gspread
import requests
from dotenv import load_dotenv
from gspread.exceptions import WorksheetNotFound

from stock_strategies.context import build_context
from stock_strategies.factors import FACTOR_REGISTRY, compute_factor


TAIPEI = ZoneInfo("Asia/Taipei")
OUTPUT_SHEET = "V3.4 Factors"

# Each completed V3.4 school gets one equal vote.  Legacy is displayed for
# comparison but excluded from composite_7style so its six factors do not
# outweigh the newer schools, most of which contain three or four factors.
STYLE_PREFIXES = [
    ("value", "價值"),
    ("growth", "成長"),
    ("momentum", "動能"),
    ("chips", "籌碼"),
    ("revenue", "營收"),
    ("reversal", "反轉"),
    ("breakout", "突破"),
]
LEGACY_PREFIX = "legacy"

HEADERS = [
    "as_of",
    "rank",
    "stock_id",
    "name",
    "category",
    "composite_7style",
    "value",
    "growth",
    "momentum",
    "chips",
    "revenue",
    "reversal",
    "breakout",
    "legacy",
    "missing_count",
    "missing_items",
    "factor_errors",
    "status",
]


def env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"Missing environment variable: {name}")
    return value


def enabled(value: Any) -> bool:
    return str(value).strip().upper() in {"TRUE", "1", "YES", "Y", "ON"}


def normalize_stock_id(value: Any) -> str:
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text.zfill(4) if text.isdigit() and len(text) < 4 else text


def safe_score(value: Any) -> float | None:
    try:
        score = float(value)
    except (TypeError, ValueError):
        return None
    if not math.isfinite(score):
        return None
    return max(0.0, min(1.0, score))


def mean_percent(values: list[float]) -> float | None:
    return round(fmean(values) * 100, 1) if values else None


def evaluate_stock(row: dict[str, Any], as_of: str) -> dict[str, Any]:
    stock_id = normalize_stock_id(row.get("stock_id", ""))
    name = str(row.get("name", "")).strip()
    category = str(row.get("category", "")).strip()
    if not stock_id:
        raise ValueError("blank stock_id")

    ctx = build_context(stock_id, as_of)
    factor_values: dict[str, float] = {}
    factor_errors: list[str] = []

    for factor_name in sorted(FACTOR_REGISTRY):
        try:
            score = safe_score(compute_factor(factor_name, ctx, {}))
            if score is not None:
                factor_values[factor_name] = score
            else:
                factor_errors.append(f"{factor_name}:non_numeric")
        except Exception as exc:  # isolate one factor instead of losing a stock
            factor_errors.append(f"{factor_name}:{type(exc).__name__}")

    style_scores: dict[str, float | None] = {}
    for prefix, _label in STYLE_PREFIXES:
        values = [
            value
            for factor_name, value in factor_values.items()
            if factor_name.startswith(f"{prefix}.")
        ]
        style_scores[prefix] = mean_percent(values)

    legacy_values = [
        value
        for factor_name, value in factor_values.items()
        if factor_name.startswith(f"{LEGACY_PREFIX}.")
    ]
    legacy_score = mean_percent(legacy_values)
    available_styles = [
        score for score in style_scores.values() if score is not None
    ]
    composite = round(fmean(available_styles), 1) if available_styles else None

    missing = [str(item) for item in (getattr(ctx, "meta", {}) or {}).get("missing", [])]
    if not available_styles:
        status = "ERROR_NO_STYLE_SCORE"
    elif factor_errors:
        status = "PARTIAL_FACTOR_ERRORS"
    elif missing:
        status = "OK_WITH_MISSING_DATA"
    else:
        status = "OK"

    return {
        "as_of": as_of,
        "rank": None,
        "stock_id": stock_id,
        "name": name,
        "category": category,
        "composite_7style": composite,
        **style_scores,
        "legacy": legacy_score,
        "missing_count": len(missing),
        "missing_items": ",".join(missing),
        "factor_errors": ",".join(factor_errors),
        "status": status,
    }


def error_result(row: dict[str, Any], as_of: str, exc: Exception) -> dict[str, Any]:
    result = {header: None for header in HEADERS}
    result.update(
        {
            "as_of": as_of,
            "stock_id": normalize_stock_id(row.get("stock_id", "")),
            "name": str(row.get("name", "")).strip(),
            "category": str(row.get("category", "")).strip(),
            "missing_count": 0,
            "status": f"ERROR_{type(exc).__name__}: {str(exc)[:120]}",
        }
    )
    return result


def open_spreadsheet() -> gspread.Spreadsheet:
    credentials = json.loads(env("GOOGLE_CREDS_JSON"))
    client = gspread.service_account_from_dict(credentials)
    return client.open_by_key(env("GOOGLE_SHEET_ID"))


def read_watchlist(spreadsheet: gspread.Spreadsheet) -> list[dict[str, Any]]:
    rows = spreadsheet.worksheet("Watchlist").get_all_records()
    selected = [row for row in rows if enabled(row.get("enabled"))]
    if not selected:
        raise RuntimeError("Watchlist has no enabled rows")
    return selected


def rank_results(results: list[dict[str, Any]]) -> list[dict[str, Any]]:
    ranked = sorted(
        results,
        key=lambda item: (
            item.get("composite_7style") is not None,
            item.get("composite_7style") or -1,
        ),
        reverse=True,
    )
    rank = 0
    for item in ranked:
        if item.get("composite_7style") is not None:
            rank += 1
            item["rank"] = rank
    return ranked


def write_snapshot(
    spreadsheet: gspread.Spreadsheet,
    results: list[dict[str, Any]],
    sheet_name: str,
) -> None:
    try:
        worksheet = spreadsheet.worksheet(sheet_name)
    except WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(
            title=sheet_name,
            rows=max(100, len(results) + 10),
            cols=len(HEADERS),
        )

    table = [HEADERS] + [[item.get(header, "") for header in HEADERS] for item in results]
    worksheet.clear()
    worksheet.resize(rows=max(100, len(table) + 5), cols=len(HEADERS))
    worksheet.update(range_name="A1", values=table, value_input_option="RAW")
    try:
        worksheet.freeze(rows=1)
        worksheet.format(
            "A1:R1",
            {
                "backgroundColor": {"red": 0.18, "green": 0.31, "blue": 0.47},
                "textFormat": {
                    "bold": True,
                    "foregroundColor": {"red": 1, "green": 1, "blue": 1},
                },
            },
        )
    except Exception as exc:
        # Presentation must never invalidate an otherwise successful snapshot.
        print(f"[WARN] worksheet formatting skipped: {type(exc).__name__}: {exc}")


def strongest_styles(item: dict[str, Any], limit: int = 2) -> str:
    candidates = [
        (label, item.get(prefix))
        for prefix, label in STYLE_PREFIXES
        if item.get(prefix) is not None
    ]
    candidates.sort(key=lambda pair: pair[1], reverse=True)
    return " / ".join(f"{label}{score:.1f}" for label, score in candidates[:limit])


def send_telegram(results: list[dict[str, Any]], as_of: str) -> None:
    scored = [item for item in results if item.get("composite_7style") is not None]
    lines = [
        f"🧪 V3.4 因子觀察報告 {as_of}",
        f"分析 {len(results)} 檔｜可評分 {len(scored)} 檔",
        "",
        "🏆 七流派觀察排名",
    ]
    for item in scored[:5]:
        lines.append(
            f"{item['rank']}. {item['stock_id']} {item['name']} "
            f"{item['composite_7style']:.1f}｜{strongest_styles(item)}"
        )
    errors = len(results) - len(scored)
    if errors:
        lines.extend(["", f"⚠️ {errors} 檔無法完成評分，請查看 Google Sheet 狀態欄。"])
    lines.extend(
        [
            "",
            "ℹ️ 這是研究排名，不是 BUY/SELL。V3.4 的回測、regime 與正式權重尚未完成。",
        ]
    )

    response = requests.post(
        f"https://api.telegram.org/bot{env('TELEGRAM_BOT_TOKEN')}/sendMessage",
        json={"chat_id": env("TELEGRAM_CHAT_ID"), "text": "\n".join(lines)},
        timeout=30,
    )
    response.raise_for_status()
    payload = response.json()
    if not payload.get("ok"):
        raise RuntimeError(payload.get("description", "Telegram send failed"))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="V3.4 Watchlist factor observer")
    parser.add_argument(
        "--as-of",
        default=datetime.now(TAIPEI).date().isoformat(),
        help="Point-in-time factor date (YYYY-MM-DD)",
    )
    parser.add_argument("--sheet-name", default=OUTPUT_SHEET)
    parser.add_argument("--no-telegram", action="store_true")
    return parser.parse_args()


def main() -> int:
    load_dotenv(".env")
    args = parse_args()
    spreadsheet = open_spreadsheet()
    watchlist = read_watchlist(spreadsheet)

    results: list[dict[str, Any]] = []
    for index, row in enumerate(watchlist, start=1):
        stock_id = normalize_stock_id(row.get("stock_id", ""))
        print(f"[{index}/{len(watchlist)}] evaluating {stock_id}", flush=True)
        try:
            results.append(evaluate_stock(row, args.as_of))
        except Exception as exc:
            print(f"[WARN] {stock_id}: {type(exc).__name__}: {exc}", flush=True)
            results.append(error_result(row, args.as_of, exc))

    results = rank_results(results)
    write_snapshot(spreadsheet, results, args.sheet_name)
    print(f"[OK] wrote {len(results)} rows to worksheet: {args.sheet_name}")

    scored_count = sum(item.get("composite_7style") is not None for item in results)
    if scored_count == 0:
        raise RuntimeError("No stocks produced a V3.4 style score")

    if not args.no_telegram:
        send_telegram(results, args.as_of)
        print("[OK] Telegram summary sent")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
