#!/usr/bin/env python3
"""
FOE Duty Dashboard — Flask API Backend
=======================================
Serves the React dashboard and provides API endpoints to:
  1. List all personnel names from the Excel workbook (validated via cols A/B/C)
  2. List available month sheets for date range selection
  3. Generate duty reports for a selected person with error checking
  4. Live-pull data from Google Sheets with caching

The Excel data is fetched from a public Google Sheet on startup and cached
locally. It auto-refreshes every 60 minutes, or on-demand via /api/refresh.

Run:
    python server.py

Then open http://localhost:5000 in your browser.
"""

import os
import re
import time
import threading
import urllib.request
import urllib.error
from datetime import datetime, date
from pathlib import Path

from flask import Flask, jsonify, request, send_from_directory
from flask_cors import CORS

# Import the consolidated reporter's core functions
from foe_duty_reporter import (
    parse_month,
    make_report,
    all_hits,
    summary_counts,
    month_label,
    date_label,
    normalize_text,
    build_month_sheet_map,
    iter_months,
    find_person_row,
    MONTHS,
    MONTH_ABBR,
)

from openpyxl import load_workbook

# ── Configuration ─────────────────────────────────────────────────────────────
GOOGLE_SHEET_ID = "161I1PqF8qVl_ASUHiPUWYl5ortwthlJdggLufZUoG4o"
GOOGLE_EXPORT_URL = (
    f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/export?format=xlsx"
)

# Cached file location (stored in project directory)
CACHE_FILE = Path(__file__).parent / "_cached_foe_data.xlsx"
CACHE_TTL_SECONDS = 3600  # 1 hour

# Fallback: local Excel file if it exists
LOCAL_FALLBACK = Path(__file__).parent / "FOE 2026.xlsx"

DEFAULT_START = "FEB25"

# Path to the React production build
DASHBOARD_DIR = Path(__file__).parent / "dashboard" / "dist"

# ── Flask App ─────────────────────────────────────────────────────────────────
app = Flask(__name__, static_folder=str(DASHBOARD_DIR), static_url_path="")
CORS(app)

# Sheet name pattern: 3-letter month + 2-digit year
SHEET_NAME_RE = re.compile(r'^[A-Z]{3}\d{2}$')

# ── Data Cache State ─────────────────────────────────────────────────────────
_cache_lock = threading.Lock()
_cache_info = {
    "last_downloaded_at": None,   # ISO string of last successful download
    "last_downloaded_ts": 0,      # Unix timestamp of last download
    "source": None,               # "google_sheets" | "local_fallback" | "cached"
    "error": None,                # Last download error, if any
}

_parsed_cache = {
    "names": None,
    "months": None,
    "loaded_ts": 0,
}

# Initialize cache info from disk if it exists on startup
if CACHE_FILE.exists():
    try:
        mtime = os.path.getmtime(CACHE_FILE)
        _cache_info["last_downloaded_ts"] = mtime
        _cache_info["last_downloaded_at"] = datetime.fromtimestamp(mtime).isoformat()
        _cache_info["source"] = "cached"
    except Exception:
        pass


def download_from_google_sheets(force: bool = False) -> dict:
    """
    Download the Google Sheet as .xlsx and save to CACHE_FILE.
    Returns a status dict with download info.

    If force=False, skips download if cache is fresh (< TTL).
    If force=True, always re-downloads.
    """
    global _cache_info

    with _cache_lock:
        # Check if cache is still fresh
        if not force and CACHE_FILE.exists():
            age = time.time() - _cache_info["last_downloaded_ts"]
            if age < CACHE_TTL_SECONDS:
                return {
                    "status": "cached",
                    "message": f"Using cached data ({int(age)}s old, TTL={CACHE_TTL_SECONDS}s)",
                    "downloaded_at": _cache_info["last_downloaded_at"],
                    "source": _cache_info["source"],
                }

    # Download from Google Sheets
    print(f"  [DATA] Downloading sheet from Google Sheets...")
    try:
        req = urllib.request.Request(
            GOOGLE_EXPORT_URL,
            headers={"User-Agent": "FOE-Duty-Dashboard/1.0"},
        )
        with urllib.request.urlopen(req, timeout=30) as response:
            data = response.read()

        # Verify it's a valid xlsx (starts with PK zip header)
        if not data[:2] == b"PK":
            raise ValueError("Downloaded file is not a valid .xlsx (may be an error page)")

        # Write to temporary file first, then atomically rename
        # This prevents 'Bad magic number' corruption when multiple workers/processes download at once
        tmp_file = CACHE_FILE.with_suffix('.xlsx.tmp')
        tmp_file.write_bytes(data)
        tmp_file.replace(CACHE_FILE)
        
        with _cache_lock:
            now = datetime.now()
            _cache_info["last_downloaded_at"] = now.isoformat()
            _cache_info["last_downloaded_ts"] = time.time()
            _cache_info["source"] = "google_sheets"
            _cache_info["error"] = None

        size_kb = len(data) / 1024
        print(f"  [DATA] OK - Downloaded {size_kb:.0f} KB -> {CACHE_FILE.name}")
        
        with _cache_lock:
            # Invalidate memory cache
            _parsed_cache["loaded_ts"] = 0

        return {
            "status": "ok",
            "message": f"Fresh data downloaded from Google Sheets ({size_kb:.0f} KB)",
            "downloaded_at": _cache_info["last_downloaded_at"],
            "source": "google_sheets",
        }

    except Exception as e:
        error_msg = str(e)
        print(f"  [DATA] FAIL - Download failed: {error_msg}")

        with _cache_lock:
            _cache_info["error"] = error_msg

        # Fallback: use cached file if it exists
        if CACHE_FILE.exists():
            return {
                "status": "fallback_cached",
                "message": f"Download failed. Using cached data from {_cache_info['last_downloaded_at'] or 'unknown time'}.",
                "downloaded_at": _cache_info["last_downloaded_at"],
                "source": "cached",
                "error": error_msg,
            }

        # Fallback: use local Excel file if available
        if LOCAL_FALLBACK.exists():
            with _cache_lock:
                _cache_info["source"] = "local_fallback"
            return {
                "status": "fallback_local",
                "message": f"Download failed. Using local file: {LOCAL_FALLBACK.name}",
                "downloaded_at": None,
                "source": "local_fallback",
                "error": error_msg,
            }

        # No data available at all
        return {
            "status": "error",
            "message": f"Download failed and no cached or local data available: {error_msg}",
            "downloaded_at": None,
            "source": None,
            "error": error_msg,
        }


def get_workbook_path() -> str:
    """
    Returns the path to the workbook file, auto-downloading if stale.
    Ensures the cache is fresh before returning.
    """
    # Try auto-refresh if cache is stale
    result = download_from_google_sheets(force=False)

    if CACHE_FILE.exists():
        return str(CACHE_FILE)

    if LOCAL_FALLBACK.exists():
        return str(LOCAL_FALLBACK)

    raise FileNotFoundError(
        "No workbook data available. Could not download from Google Sheets "
        "and no local file found."
    )


def get_data_status() -> dict:
    """Returns the current data freshness status for the frontend."""
    with _cache_lock:
        info = dict(_cache_info)

    age_seconds = None
    if info["last_downloaded_ts"]:
        age_seconds = int(time.time() - info["last_downloaded_ts"])

    return {
        "last_refreshed": info["last_downloaded_at"],
        "source": info["source"],
        "age_seconds": age_seconds,
        "ttl_seconds": CACHE_TTL_SECONDS,
        "is_stale": age_seconds is not None and age_seconds > CACHE_TTL_SECONDS,
        "error": info["error"],
    }


# ── Row Validation Helpers ────────────────────────────────────────────────────

def is_valid_number(value) -> bool:
    """Check if a value looks like a row number (1, 2, 3, ...)."""
    if value is None:
        return False
    try:
        n = int(value)
        return n > 0
    except (ValueError, TypeError):
        return False


def is_valid_rank(value) -> bool:
    """Check if column B has a non-empty rank string."""
    if value is None:
        return False
    text = str(value).strip()
    return len(text) > 0


def is_valid_name(value) -> bool:
    """Check if column C has a plausible person name (multi-word uppercase)."""
    if value is None or not isinstance(value, str):
        return False
    name = value.strip().upper()
    if " " not in name:
        return False
    if len(name) < 4:
        return False
    return True


# ── Data Extraction Functions ─────────────────────────────────────────────────

def load_data_into_cache(workbook_path: str):
    """Parses names and months from the Excel file and caches them in memory."""
    global _parsed_cache
    print(f"  [CACHE] Parsing Excel file to memory...")
    try:
        names = extract_all_names(workbook_path)
        months = get_available_months(workbook_path)
        with _cache_lock:
            _parsed_cache["names"] = names
            _parsed_cache["months"] = months
            _parsed_cache["loaded_ts"] = time.time()
        print(f"  [CACHE] -> Loaded {len(names)} names and {len(months)} months")
    except Exception as e:
        print(f"  [CACHE] FAIL - Failed to parse: {e}")
        raise e

def extract_all_names(workbook_path: str) -> list[str]:
    """
    Scans all monthly sheets in the workbook and extracts unique personnel
    names. Validates each row by checking:
      - Column A: must be a number (1, 2, 3, ...)
      - Column B: must have a rank (non-empty)
      - Column C: must be a multi-word string (person name)
    """
    wb = load_workbook(workbook_path, read_only=True, data_only=True)
    names = set()

    for sheet_name in wb.sheetnames:
        if not SHEET_NAME_RE.match(sheet_name):
            continue

        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=4, min_col=1, max_col=3, values_only=True):
            col_a = row[0] if len(row) >= 1 else None
            col_b = row[1] if len(row) >= 2 else None
            col_c = row[2] if len(row) >= 3 else None

            if not is_valid_number(col_a):
                continue
            if not is_valid_rank(col_b):
                continue
            if not is_valid_name(col_c):
                continue

            name = col_c.strip().upper()
            names.add(name)

    wb.close()
    return sorted(names)


def get_available_months(workbook_path: str) -> list[dict]:
    """
    Returns a sorted list of available month sheets from the workbook.
    Each entry has: { "key": "FEB25", "label": "Feb 2025", "date": "2025-02-01" }
    """
    wb = load_workbook(workbook_path, read_only=True, data_only=True)
    month_map = build_month_sheet_map(wb)
    wb.close()

    months = []
    for month_date, sheet_name in sorted(month_map.items()):
        abbr = MONTH_ABBR.get(month_date.month, "")
        yr = month_date.year % 100
        months.append({
            "key": f"{abbr}{yr:02d}",
            "label": month_label(month_date),
            "date": month_date.strftime("%Y-%m-%d"),
        })

    return months


def check_person_in_range(
    workbook_path: str,
    target_name: str,
    start: date,
    end: date,
) -> dict:
    """
    Check if the person exists in any sheet within the selected range.
    Also validates the found row has a valid number (col A) and rank (col B).
    """
    wb = load_workbook(workbook_path, read_only=True, data_only=True)
    month_map = build_month_sheet_map(wb)

    found_in = []
    missing_in = []

    for month in iter_months(start, end):
        sheet_name = month_map.get(month, "")
        if not sheet_name:
            missing_in.append(month_label(month))
            continue

        ws = wb[sheet_name]
        row_num = find_person_row(ws, target_name)

        if row_num is not None:
            row_data = None
            for r in ws.iter_rows(min_row=row_num, max_row=row_num, min_col=1, max_col=3, values_only=True):
                row_data = r
                break

            if row_data and is_valid_number(row_data[0]) and is_valid_rank(row_data[1]):
                found_in.append(month_label(month))
            else:
                missing_in.append(month_label(month))
        else:
            missing_in.append(month_label(month))

    wb.close()

    return {
        "found": len(found_in) > 0,
        "found_in": found_in,
        "missing_in": missing_in,
        "total_sheets_checked": len(found_in) + len(missing_in),
    }


def build_duty_payload(
    workbook_path: str,
    target_name: str,
    target_gen: str = "",
    start_str: str = DEFAULT_START,
    end_str: str = "CURRENT",
) -> dict:
    """
    Runs the duty report for a given person and returns the JSON payload dict.
    """
    start = parse_month(start_str)
    end = parse_month(end_str, allow_current=True)

    results = make_report(
        workbook_path=workbook_path,
        target_name=target_name,
        target_gen=target_gen,
        start=start,
        end=end,
    )

    hits = all_hits(results)
    weekdays, fridays, weekends = summary_counts(hits)
    total_guard = sum(h.duty_type == "Guard" for h in hits)
    total_bds = sum(h.duty_type == "BDS" for h in hits)

    return {
        "metadata": {
            "name": target_name.upper(),
            "gen": target_gen,
            "generated_at": datetime.now().isoformat(),
            "range": f"{month_label(start)} to {month_label(end)}",
            "start": start.strftime("%Y-%m-%d"),
            "end": end.strftime("%Y-%m-%d"),
        },
        "summary": {
            "total_duties": len(hits),
            "total_guard": total_guard,
            "total_bds": total_bds,
            "weekdays": weekdays,
            "fridays": fridays,
            "weekends": weekends,
        },
        "duties": [
            {
                "number": idx,
                "date": hit.duty_date.strftime("%Y-%m-%d"),
                "date_label": date_label(hit.duty_date),
                "month_label": month_label(hit.duty_date),
                "month_key": hit.duty_date.strftime("%b %y").upper(),
                "day": hit.day_name,
                "day_type": hit.category,
                "duty_type": hit.duty_type,
                "duty_display": hit.duty_display,
                "sheet": hit.sheet,
                "row": hit.row,
            }
            for idx, hit in enumerate(hits, start=1)
        ],
    }


# ── API Routes ────────────────────────────────────────────────────────────────

@app.route("/api/names", methods=["GET"])
def get_names():
    """Return all unique, validated personnel names from the workbook."""
    try:
        wb_path = get_workbook_path()
        
        # Check if memory cache is valid
        with _cache_lock:
            ts = _cache_info["last_downloaded_ts"]
            loaded = _parsed_cache["loaded_ts"]
            names = _parsed_cache["names"]
            
        if not names or loaded < ts:
            load_data_into_cache(wb_path)
            with _cache_lock:
                names = _parsed_cache["names"]

        data_status = get_data_status()
        return jsonify({"names": names, "data_status": data_status})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/months", methods=["GET"])
def get_months():
    """Return all available month sheets from the workbook."""
    try:
        wb_path = get_workbook_path()
        
        # Check if memory cache is valid
        with _cache_lock:
            ts = _cache_info["last_downloaded_ts"]
            loaded = _parsed_cache["loaded_ts"]
            months = _parsed_cache["months"]
            
        if not months or loaded < ts:
            load_data_into_cache(wb_path)
            with _cache_lock:
                months = _parsed_cache["months"]
                
        return jsonify({"months": months})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/validate", methods=["POST"])
def validate_person():
    """
    Check if a person exists within the selected date range.
    Body: {"name": "...", "start": "FEB25", "end": "CURRENT"}
    """
    body = request.get_json(silent=True) or {}
    name = body.get("name", "").strip()
    if not name:
        return jsonify({"error": "No name provided."}), 400

    start_str = body.get("start", DEFAULT_START)
    end_str = body.get("end", "CURRENT")

    try:
        wb_path = get_workbook_path()
        start = parse_month(start_str)
        end = parse_month(end_str, allow_current=True)

        if start > end:
            return jsonify({
                "error": "Start month cannot be after end month."
            }), 400

        result = check_person_in_range(wb_path, name, start, end)

        if not result["found"]:
            return jsonify({
                "valid": False,
                "error": f"{name} was not found in any sheet between {month_label(start)} and {month_label(end)}.",
                "details": result,
            }), 200

        return jsonify({
            "valid": True,
            "details": result,
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/duties", methods=["POST"])
def get_duties():
    """Generate a duty report for the specified person."""
    body = request.get_json(silent=True) or {}
    name = body.get("name", "").strip()
    if not name:
        return jsonify({"error": "No name provided."}), 400

    gen = body.get("gen", "")
    start = body.get("start", DEFAULT_START)
    end = body.get("end", "CURRENT")

    try:
        wb_path = get_workbook_path()

        start_date = parse_month(start)
        end_date = parse_month(end, allow_current=True)
        if start_date > end_date:
            return jsonify({"error": "Start month cannot be after end month."}), 400

        payload = build_duty_payload(wb_path, name, gen, start, end)
        payload["data_status"] = get_data_status()
        return jsonify(payload)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/refresh", methods=["POST"])
def refresh_data():
    """
    Force re-download the Google Sheet data.
    Returns download status + data freshness info.
    """
    try:
        result = download_from_google_sheets(force=True)
        result["data_status"] = get_data_status()
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/status", methods=["GET"])
def data_status():
    """Return the current data freshness status."""
    return jsonify(get_data_status())


# ── Serve React SPA ──────────────────────────────────────────────────────────

@app.route("/", defaults={"path": ""})
@app.route("/<path:path>")
def serve_react(path):
    """Serve React build files, falling back to index.html for SPA routing."""
    file_path = DASHBOARD_DIR / path
    if path and file_path.exists():
        return send_from_directory(str(DASHBOARD_DIR), path)
    return send_from_directory(str(DASHBOARD_DIR), "index.html")


# ── Main ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print(f"\n  Google Sheet: {GOOGLE_SHEET_ID}")
    print(f"  Cache File:   {CACHE_FILE}")
    print(f"  Cache TTL:    {CACHE_TTL_SECONDS}s ({CACHE_TTL_SECONDS // 60} min)")
    print(f"  Dashboard:    {DASHBOARD_DIR}")

    # Download sheet data at startup (will skip if cache file is < 1 hour old)
    print(f"\n  Checking Google Sheets data freshness...")
    startup_result = download_from_google_sheets(force=False)
    print(f"  -> {startup_result['message']}\n")

    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "false").lower() == "true"
    print(f"  Starting server at http://localhost:{port}\n")
    app.run(host="0.0.0.0", port=port, debug=debug)
