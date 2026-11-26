#!/usr/bin/env python3
"""Fetch the six configured public Chartink screeners and create Excel/JSON reports.

All screening conditions and result membership remain controlled by the saved
Chartink screeners. This script does not alter scan clauses or place orders.
"""

import asyncio
import json
import re
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
from playwright.async_api import async_playwright

# Five stock screeners followed by the sector/index screener.
SCREENS = [
    ("ADVANCEGFSCASHSTOCKS", "https://chartink.com/screener/advancegfscashstocks", "stocks"),
    ("ADVANCEGFSSTOCKSFUTURE", "https://chartink.com/screener/advancegfsstocksfuture", "stocks"),
    ("VLAUEBUYSTOCKSCASH", "https://chartink.com/screener/vlauebuystockscash", "stocks"),
    ("GFSSTOCKCASH", "https://chartink.com/screener/gfsstockcash", "stocks"),
    ("GFSWITHSECTOR", "https://chartink.com/screener/gfswithsector", "stocks"),
    ("SECTORINDICES", "https://chartink.com/screener/gfs-60610539", "indices"),
]

STOCK_COLUMNS = [
    "Sr", "Stock Name", "Daily RSI", "Weekly RSI", "Monthly RSI",
    "Close", "%_change", "Volume",
]
INDEX_COLUMNS = [
    "Sr", "Index Name", "Daily RSI", "Weekly RSI", "Monthly RSI",
    "Close", "%_change",
]

OUT_XLSX = Path("gfs_dashboard_latest.xlsx")
OUT_JSON = Path("gfs_dashboard_latest.json")


def pick(record, *keys):
    """Return the first non-empty value among known Chartink field aliases."""
    lowered = {str(k).lower(): v for k, v in record.items()}
    for key in keys:
        value = lowered.get(key.lower())
        if value not in (None, ""):
            return value
    return ""


def stock_view(records):
    rows = []
    for i, rec in enumerate(records, 1):
        rows.append({
            "Sr": i,
            "Stock Name": pick(rec, "name", "stock_name", "company_name", "nsecode", "symbol"),
            "Daily RSI": pick(rec, "daily_rsi", "rsi", "rsi_14", "rsi14"),
            "Weekly RSI": pick(rec, "weekly_rsi", "rsi_weekly", "weeklyrsi"),
            "Monthly RSI": pick(rec, "monthly_rsi", "rsi_monthly", "monthlyrsi"),
            "Close": pick(rec, "close", "price"),
            "%_change": pick(rec, "per_chg", "percent_change", "pchange", "% change"),
            "Volume": pick(rec, "volume"),
        })
    return rows


def index_view(records):
    rows = []
    for i, rec in enumerate(records, 1):
        rows.append({
            "Sr": i,
            "Index Name": pick(rec, "name", "stock_name", "company_name", "nsecode", "symbol"),
            "Daily RSI": pick(rec, "daily_rsi", "rsi", "rsi_14", "rsi14"),
            "Weekly RSI": pick(rec, "weekly_rsi", "rsi_weekly", "weeklyrsi"),
            "Monthly RSI": pick(rec, "monthly_rsi", "rsi_monthly", "monthlyrsi"),
            "Close": pick(rec, "close", "price"),
            "%_change": pick(rec, "per_chg", "percent_change", "pchange", "% change"),
        })
    return rows


async def run_screen(page, title, url, kind):
    captured = []

    async def on_response(response):
        try:
            if response.request.method.upper() != "POST":
                return
            if "screener/process" not in response.url.lower():
                return
            content_type = (response.headers.get("content-type") or "").lower()
            if "json" not in content_type:
                return
            body = await response.json()
            if isinstance(body, dict) and isinstance(body.get("data"), list):
                captured.append(body)
        except Exception:
            # Ignore unrelated or unreadable network responses.
            pass

    page.on("response", on_response)
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=60000)
        # Allow Chartink's page scripts to submit the screener request.
        for _ in range(12):
            if captured:
                break
            await page.wait_for_timeout(5000)
    except Exception as exc:
        return {
            "title": title, "url": url, "kind": kind,
            "status": "ERROR", "error": str(exc),
            "records": [], "raw_fields": [],
        }
    finally:
        page.remove_listener("response", on_response)

    if not captured:
        return {
            "title": title, "url": url, "kind": kind,
            "status": "NO_LIVE_RESPONSE_CAPTURED",
            "records": [], "raw_fields": [],
        }

    records = [r for r in captured[-1].get("data", []) if isinstance(r, dict)]
    fields = list(dict.fromkeys(k for record in records for k in record.keys()))
    return {
        "title": title, "url": url, "kind": kind, "status": "OK",
        "records": records, "raw_fields": fields,
    }


async def main():
    extracted_at = datetime.now(ZoneInfo("Asia/Kolkata")).isoformat(timespec="seconds")

    async with async_playwright() as playwright:
        browser = await playwright.chromium.launch(headless=True)
        context = await browser.new_context(
            viewport={"width": 1440, "height": 1000},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/131.0 Safari/537.36"
            ),
        )
        page = await context.new_page()
        results = []
        for title, url, kind in SCREENS:
            print(f"Running {title} ...", flush=True)
            result = await run_screen(page, title, url, kind)
            result["extracted_at"] = extracted_at
            results.append(result)
            print(
                f"{result['status']} | rows={len(result['records'])} "
                f"| fields={len(result['raw_fields'])}",
                flush=True,
            )
        await browser.close()

    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as writer:
        dashboard_rows = [
            {"Field": "Extracted at", "Value": extracted_at},
            {"Field": "Timezone", "Value": "Asia/Kolkata"},
            {"Field": "Data source", "Value": "Chartink saved screeners only"},
            {"Field": "Configured screeners", "Value": len(SCREENS)},
        ]
        for result in results:
            dashboard_rows.append({
                "Field": result["title"],
                "Value": f"{result['status']} — {len(result['records'])} rows",
            })
        pd.DataFrame(dashboard_rows).to_excel(
            writer, sheet_name="Dashboard", index=False
        )

        for result in results:
            is_index = result["kind"] == "indices"
            rows = index_view(result["records"]) if is_index else stock_view(result["records"])
            columns = INDEX_COLUMNS if is_index else STOCK_COLUMNS
            sheet = re.sub(r"[:\\/?*\[\]]", "_", result["title"])[:31]
            pd.DataFrame(rows, columns=columns).to_excel(
                writer, sheet_name=sheet, index=False
            )

            # Keep the original Chartink fields for auditing and mapping fixes.
            raw_sheet = (sheet[:25] + "_RAW")[:31]
            pd.DataFrame(result["records"]).to_excel(
                writer, sheet_name=raw_sheet, index=False
            )

    payload = {
        "extracted_at": extracted_at,
        "timezone": "Asia/Kolkata",
        "source": "Chartink saved screeners only",
        "screeners": results,
        "stock_columns": STOCK_COLUMNS,
        "index_columns": INDEX_COLUMNS,
    }
    OUT_JSON.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    print(f"\nSaved {OUT_XLSX.resolve()}")
    print(f"Saved {OUT_JSON.resolve()}")
    failed = [r["title"] for r in results if r["status"] != "OK"]
    if failed:
        raise SystemExit(
            "One or more Chartink screeners did not return data: "
            + ", ".join(failed)
            + ". Partial Excel/JSON files were still generated for diagnosis."
        )


if __name__ == "__main__":
    asyncio.run(main())
