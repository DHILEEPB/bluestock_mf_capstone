"""
live_nav_fetch.py
=================
Mutual Fund Analytics Capstone Project -- Day 1
-----------------------------------------------
Purpose : Fetch live / latest NAV data for a list of scheme codes from the
          public MFAPI (https://mfapi.in) REST API, validate scheme codes
          against the AMFI master list, and persist results to
          data/processed/live_nav_<date>.csv and .parquet.

Author  : Capstone Project -- Senior Data Engineer
Created : 2026-06-12
Standard: PEP8 compliant

API Reference:
    Latest NAV  : GET https://api.mfapi.in/mf/{scheme_code}
    All schemes : GET https://api.mfapi.in/mf
    AMFI master : https://www.amfiindia.com/spages/NAVAll.txt

Usage
-----
Run from the project root:
    python scripts/live_nav_fetch.py

    Optional CLI flags:
        --scheme-codes 119551 120503 122639   # space-separated scheme codes
        --output-name  my_nav_snapshot        # custom output file stem
        --validate-amfi                       # cross-check codes vs AMFI master
        --threads 10                          # parallel worker count (default 8)
"""

import sys
import csv
import time
import logging
import logging.handlers
import argparse
import concurrent.futures
from io import StringIO
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import requests
import pandas as pd
import numpy as np

# ---------------------------------------------------------------------------
# Project paths
# ---------------------------------------------------------------------------

PROJECT_ROOT  = Path(__file__).resolve().parent.parent
PROCESSED_DIR = PROJECT_ROOT / "data" / "processed"
LOGS_DIR      = PROJECT_ROOT / "logs"

for _dir in (PROCESSED_DIR, LOGS_DIR):
    _dir.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

LOG_FILE = LOGS_DIR / f"live_nav_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

logger = logging.getLogger("live_nav_fetch")
logger.setLevel(logging.DEBUG)

_fmt = logging.Formatter(
    fmt="%(asctime)s  [%(levelname)-8s]  %(name)s -- %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)

_fh = logging.handlers.RotatingFileHandler(
    LOG_FILE, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8"
)
_fh.setLevel(logging.DEBUG)
_fh.setFormatter(_fmt)

_ch = logging.StreamHandler(open(sys.stdout.fileno(), mode="w", encoding="utf-8", closefd=False))
_ch.setLevel(logging.INFO)
_ch.setFormatter(_fmt)

logger.addHandler(_fh)
logger.addHandler(_ch)


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

MFAPI_BASE_URL   = "https://api.mfapi.in/mf"
AMFI_MASTER_URL  = "https://www.amfiindia.com/spages/NAVAll.txt"
REQUEST_TIMEOUT  = 15          # seconds per HTTP request
MAX_RETRIES      = 3           # retry count on transient failures
RETRY_BACKOFF    = 2.0         # seconds between retries (exponential)
DEFAULT_THREADS  = 8           # parallel workers

# A curated list of popular scheme codes used when none are provided by CLI
DEFAULT_SCHEME_CODES: List[int] = [
    119551,   # Axis Bluechip Fund -- Growth
    120503,   # Mirae Asset Large Cap Fund -- Growth
    122639,   # SBI Small Cap Fund -- Growth
    125497,   # HDFC Mid-Cap Opportunities Fund -- Growth
    118989,   # ICICI Prudential Technology Fund -- Growth
    119598,   # Kotak Bluechip Fund -- Growth
    120465,   # Nippon India Large Cap Fund -- Growth
    120594,   # DSP Top 100 Equity Fund -- Growth
    120716,   # Canara Robeco Bluechip Equity Fund -- Growth
    135781,   # Parag Parikh Flexi Cap Fund -- Growth
]


# ---------------------------------------------------------------------------
# HTTP utilities
# ---------------------------------------------------------------------------


def _get_with_retry(url: str, timeout: int = REQUEST_TIMEOUT,
                    max_retries: int = MAX_RETRIES) -> Optional[requests.Response]:
    """
    Perform an HTTP GET request with exponential-backoff retry logic.

    Parameters
    ----------
    url        : str -- Target URL.
    timeout    : int -- Request timeout in seconds.
    max_retries: int -- Maximum number of retry attempts.

    Returns
    -------
    Optional[requests.Response]
        Response object on success, or None after all retries are exhausted.
    """
    for attempt in range(1, max_retries + 1):
        try:
            response = requests.get(url, timeout=timeout)
            response.raise_for_status()
            return response
        except requests.exceptions.HTTPError as exc:
            # 404 means the scheme code does not exist -- no point retrying
            if exc.response is not None and exc.response.status_code == 404:
                logger.warning("404 Not Found: %s", url)
                return None
            logger.warning("HTTP error (attempt %d/%d): %s -- %s",
                           attempt, max_retries, url, exc)
        except requests.exceptions.ConnectionError as exc:
            logger.warning("Connection error (attempt %d/%d): %s -- %s",
                           attempt, max_retries, url, exc)
        except requests.exceptions.Timeout:
            logger.warning("Timeout (attempt %d/%d): %s", attempt, max_retries, url)
        except requests.exceptions.RequestException as exc:
            logger.error("Unrecoverable request error: %s -- %s", url, exc)
            return None

        if attempt < max_retries:
            sleep_secs = RETRY_BACKOFF ** attempt
            logger.debug("Retrying in %.1f seconds ...", sleep_secs)
            time.sleep(sleep_secs)

    logger.error("All %d attempts failed for: %s", max_retries, url)
    return None


# ---------------------------------------------------------------------------
# AMFI master validation
# ---------------------------------------------------------------------------


def fetch_amfi_master() -> Optional[pd.DataFrame]:
    """
    Download the AMFI NAV master list and parse it into a DataFrame.

    The AMFI master file is pipe-delimited text with the format:
        Scheme Code;ISIN Div Payout/ Growth;ISIN Div Reinvestment;
        Scheme Name;Net Asset Value;Date

    Returns
    -------
    Optional[pd.DataFrame]
        Parsed master DataFrame with columns [scheme_code, scheme_name, nav, nav_date],
        or None on failure.
    """
    logger.info("Fetching AMFI master list from %s ...", AMFI_MASTER_URL)
    response = _get_with_retry(AMFI_MASTER_URL)

    if response is None:
        logger.error("Could not fetch AMFI master list.")
        return None

    # Parse -- the file contains category headers (non-numeric first fields)
    # that we must skip. Each data row has a numeric Scheme Code as field 0.
    valid_rows = []
    for line in response.text.splitlines():
        parts = line.strip().split(";")
        if len(parts) < 6:
            continue
        # Only data rows have a numeric scheme code
        if parts[0].strip().isdigit():
            valid_rows.append({
                "scheme_code" : int(parts[0].strip()),
                "isin_growth" : parts[1].strip(),
                "isin_reinv"  : parts[2].strip(),
                "scheme_name" : parts[3].strip(),
                "nav"         : parts[4].strip(),
                "nav_date"    : parts[5].strip(),
            })

    if not valid_rows:
        logger.error("AMFI master list parsed 0 valid rows -- check format.")
        return None

    df = pd.DataFrame(valid_rows)
    df["nav"] = pd.to_numeric(df["nav"], errors="coerce")
    logger.info("AMFI master: %d schemes loaded.", len(df))
    return df


def validate_codes_against_amfi(
    scheme_codes: List[int],
    amfi_master: pd.DataFrame
) -> Tuple[List[int], List[int]]:
    """
    Split scheme_codes into valid (present in AMFI master) and invalid lists.

    Parameters
    ----------
    scheme_codes : List[int]        -- Codes to validate.
    amfi_master  : pd.DataFrame     -- Master DataFrame from fetch_amfi_master().

    Returns
    -------
    Tuple[List[int], List[int]]
        (valid_codes, invalid_codes)
    """
    master_codes = set(amfi_master["scheme_code"].dropna().astype(int))
    valid   = [c for c in scheme_codes if c in master_codes]
    invalid = [c for c in scheme_codes if c not in master_codes]

    if invalid:
        logger.warning("Invalid/unknown AMFI codes (not in master): %s", invalid)
    logger.info("Validation: %d valid | %d invalid out of %d codes.",
                len(valid), len(invalid), len(scheme_codes))
    return valid, invalid


# ---------------------------------------------------------------------------
# NAV fetch -- single scheme
# ---------------------------------------------------------------------------


def fetch_nav_for_scheme(scheme_code: int) -> Optional[Dict]:
    """
    Fetch the latest NAV record for a single scheme code from mfapi.in.

    API response structure (JSON):
    {
        "meta": {
            "fund_house": "...",
            "scheme_type": "...",
            "scheme_category": "...",
            "scheme_code": 119551,
            "scheme_name": "..."
        },
        "data": [
            {"date": "12-06-2024", "nav": "123.4500"},
            ...
        ],
        "status": "SUCCESS"
    }

    Parameters
    ----------
    scheme_code : int -- AMFI scheme code.

    Returns
    -------
    Optional[Dict]
        Flattened dictionary with latest NAV details, or None on failure.
    """
    url      = f"{MFAPI_BASE_URL}/{scheme_code}"
    response = _get_with_retry(url)

    if response is None:
        return None

    try:
        payload = response.json()
    except ValueError as exc:
        logger.error("[%d] JSON decode error: %s", scheme_code, exc)
        return None

    if payload.get("status") != "SUCCESS":
        logger.warning("[%d] API status != SUCCESS: %s", scheme_code, payload.get("status"))
        return None

    meta     = payload.get("meta", {})
    nav_data = payload.get("data", [])

    if not nav_data:
        logger.warning("[%d] No NAV data returned by API.", scheme_code)
        return None

    latest = nav_data[0]   # mfapi returns data newest-first

    return {
        "scheme_code"    : scheme_code,
        "scheme_name"    : meta.get("scheme_name", ""),
        "fund_house"     : meta.get("fund_house", ""),
        "scheme_type"    : meta.get("scheme_type", ""),
        "scheme_category": meta.get("scheme_category", ""),
        "nav"            : float(latest.get("nav", np.nan)),
        "nav_date"       : latest.get("date", ""),
        "fetched_at"     : datetime.now().isoformat(timespec="seconds"),
        "api_status"     : "SUCCESS",
    }


# ---------------------------------------------------------------------------
# Parallel batch fetcher
# ---------------------------------------------------------------------------


def fetch_nav_batch(
    scheme_codes: List[int],
    max_workers: int = DEFAULT_THREADS,
) -> pd.DataFrame:
    """
    Fetch live NAV for a batch of scheme codes using a thread pool.

    Parameters
    ----------
    scheme_codes : List[int]  -- List of AMFI scheme codes.
    max_workers  : int        -- Maximum parallel threads.

    Returns
    -------
    pd.DataFrame
        DataFrame with one row per scheme code, including failed fetches
        (marked with api_status = 'FAILED').
    """
    total    = len(scheme_codes)
    results  = []
    failed   = []

    logger.info("Fetching NAV for %d schemes using %d threads ...", total, max_workers)

    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_code = {
            executor.submit(fetch_nav_for_scheme, code): code
            for code in scheme_codes
        }

        completed = 0
        for future in concurrent.futures.as_completed(future_to_code):
            code      = future_to_code[future]
            completed += 1

            try:
                record = future.result()
                if record:
                    results.append(record)
                    logger.debug("[%d] [OK] NAV=%.4f  Date=%s",
                                 code, record["nav"], record["nav_date"])
                else:
                    failed.append(code)
                    results.append({
                        "scheme_code" : code,
                        "scheme_name" : "",
                        "fund_house"  : "",
                        "scheme_type" : "",
                        "scheme_category": "",
                        "nav"         : np.nan,
                        "nav_date"    : "",
                        "fetched_at"  : datetime.now().isoformat(timespec="seconds"),
                        "api_status"  : "FAILED",
                    })
            except Exception as exc:
                logger.error("[%d] Unexpected error: %s", code, exc)
                failed.append(code)

            # Progress log every 10 completions
            if completed % 10 == 0 or completed == total:
                logger.info("Progress: %d / %d schemes fetched.", completed, total)

    df = pd.DataFrame(results)
    if not df.empty:
        # Enforce types
        df["nav"]         = pd.to_numeric(df["nav"], errors="coerce")
        df["nav_date"]    = pd.to_datetime(df["nav_date"], dayfirst=True, errors="coerce")
        df["scheme_code"] = df["scheme_code"].astype("Int64")
        df.sort_values("scheme_code", inplace=True)
        df.reset_index(drop=True, inplace=True)

    n_success = len(df[df["api_status"] == "SUCCESS"]) if not df.empty else 0
    n_failed  = len(failed)
    logger.info("Batch complete: %d success | %d failed.", n_success, n_failed)

    if failed:
        logger.warning("Failed scheme codes: %s", failed)

    return df


# ---------------------------------------------------------------------------
# Persist results
# ---------------------------------------------------------------------------


def save_nav_snapshot(df: pd.DataFrame, stem: str) -> Dict[str, Path]:
    """
    Save the NAV snapshot to data/processed/ as CSV and Parquet.

    Parameters
    ----------
    df   : pd.DataFrame -- NAV snapshot DataFrame.
    stem : str          -- Output filename stem (without extension).

    Returns
    -------
    Dict[str, Path]
        Mapping of format -> saved file path.
    """
    paths = {}

    # CSV
    csv_path = PROCESSED_DIR / f"{stem}.csv"
    df.to_csv(csv_path, index=False, encoding="utf-8")
    paths["csv"] = csv_path
    logger.info("Saved CSV -> %s", csv_path)

    # Parquet
    parquet_path = PROCESSED_DIR / f"{stem}.parquet"
    try:
        df.to_parquet(parquet_path, index=False, engine="pyarrow")
        paths["parquet"] = parquet_path
        logger.info("Saved Parquet -> %s", parquet_path)
    except ImportError:
        logger.warning("pyarrow not installed -- skipping Parquet output.")
    except Exception as exc:
        logger.error("Parquet save failed: %s", exc)

    return paths


# ---------------------------------------------------------------------------
# CLI argument parser
# ---------------------------------------------------------------------------


def _build_arg_parser() -> argparse.ArgumentParser:
    """
    Build and return the CLI argument parser.

    Returns
    -------
    argparse.ArgumentParser
    """
    parser = argparse.ArgumentParser(
        prog="live_nav_fetch.py",
        description="Fetch live mutual fund NAV data from mfapi.in",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "--scheme-codes",
        nargs="+",
        type=int,
        default=DEFAULT_SCHEME_CODES,
        metavar="CODE",
        help="Space-separated list of AMFI scheme codes to fetch.",
    )
    parser.add_argument(
        "--output-name",
        type=str,
        default=f"live_nav_{datetime.now().strftime('%Y%m%d')}",
        help="Output filename stem (written to data/processed/).",
    )
    parser.add_argument(
        "--validate-amfi",
        action="store_true",
        default=True,
        help="Cross-check scheme codes against AMFI master list before fetching.",
    )
    parser.add_argument(
        "--threads",
        type=int,
        default=DEFAULT_THREADS,
        help="Number of parallel worker threads.",
    )
    return parser


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main() -> int:
    """
    Main entry point.

    Workflow:
        1. Parse CLI arguments.
        2. (Optional) Fetch AMFI master and validate scheme codes.
        3. Fetch live NAV in parallel for all valid codes.
        4. Save results to data/processed/.
        5. Print summary.

    Returns
    -------
    int
        0 on success, 1 on error.
    """
    logger.info("=" * 60)
    logger.info("  Mutual Fund Analytics -- Live NAV Fetch Pipeline")
    logger.info("  Started : %s", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    logger.info("=" * 60)

    parser = _build_arg_parser()
    args   = parser.parse_args()

    scheme_codes: List[int] = args.scheme_codes
    logger.info("Scheme codes requested: %d", len(scheme_codes))

    # Step 1 -- Optional AMFI validation
    amfi_invalid: List[int] = []
    if args.validate_amfi:
        amfi_master = fetch_amfi_master()
        if amfi_master is not None:
            scheme_codes, amfi_invalid = validate_codes_against_amfi(
                scheme_codes, amfi_master
            )
        else:
            logger.warning("AMFI validation skipped (master fetch failed).")

    if not scheme_codes:
        logger.error("No valid scheme codes to fetch. Exiting.")
        return 1

    # Step 2 -- Parallel NAV fetch
    nav_df = fetch_nav_batch(scheme_codes, max_workers=args.threads)

    if nav_df.empty:
        logger.error("No NAV data fetched. Check network connectivity.")
        return 1

    # Step 3 -- Append rows for AMFI-invalid codes (for audit trail)
    if amfi_invalid:
        invalid_rows = pd.DataFrame([{
            "scheme_code"    : code,
            "scheme_name"    : "",
            "fund_house"     : "",
            "scheme_type"    : "",
            "scheme_category": "",
            "nav"            : np.nan,
            "nav_date"       : pd.NaT,
            "fetched_at"     : datetime.now().isoformat(timespec="seconds"),
            "api_status"     : "AMFI_INVALID",
        } for code in amfi_invalid])
        nav_df = pd.concat([nav_df, invalid_rows], ignore_index=True)

    # Step 4 -- Save
    saved_paths = save_nav_snapshot(nav_df, args.output_name)

    # Step 5 -- Summary report to console
    _print_summary(nav_df, saved_paths)

    logger.info("=" * 60)
    logger.info("  Completed : %s", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    logger.info("=" * 60)
    return 0


def _print_summary(df: pd.DataFrame, saved_paths: Dict[str, Path]) -> None:
    """
    Print a human-readable NAV summary table and fetch statistics.

    Parameters
    ----------
    df          : pd.DataFrame     -- NAV snapshot DataFrame.
    saved_paths : Dict[str, Path]  -- Output file paths.

    Returns
    -------
    None
    """
    logger.info("-" * 60)
    logger.info("NAV SNAPSHOT SUMMARY")
    logger.info("-" * 60)

    success_df = df[df["api_status"] == "SUCCESS"].copy()

    if not success_df.empty:
        display_cols = ["scheme_code", "scheme_name", "nav", "nav_date", "fund_house"]
        display_cols = [c for c in display_cols if c in success_df.columns]
        logger.info("\n%s", success_df[display_cols].to_string(index=False))

    logger.info("")
    logger.info("Total codes processed : %d", len(df))
    logger.info("  [OK] SUCCESS           : %d", (df["api_status"] == "SUCCESS").sum())
    logger.info("  [FAIL] FAILED            : %d", (df["api_status"] == "FAILED").sum())
    logger.info("  [WARN] AMFI_INVALID      : %d", (df["api_status"] == "AMFI_INVALID").sum())

    if not success_df.empty and "nav" in success_df.columns:
        logger.info("")
        logger.info("NAV Statistics (successful fetches):")
        logger.info("  Min NAV : INR %.4f", success_df["nav"].min())
        logger.info("  Max NAV : INR %.4f", success_df["nav"].max())
        logger.info("  Avg NAV : INR %.4f", success_df["nav"].mean())

    logger.info("")
    logger.info("Output files:")
    for fmt, path in saved_paths.items():
        logger.info("  [%-7s] %s", fmt.upper(), path)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    sys.exit(main())
