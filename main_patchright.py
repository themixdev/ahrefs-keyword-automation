"""
Ahrefs Content Explorer listicle exporter - Patchright edition.

Same job as main_fixed.py (one Content Explorer export per keyword), but:
  * Log in ONCE. The browser profile is saved in ./patchright_profile, so later
    runs start already logged in and can run unattended.
  * Downloads are captured directly by the browser (no Temp-folder guessing).
  * "No results" is detected and logged separately from real errors, and is
    NOT retried on the next run.
  * Filters are set in one place (FILTERS below).
  * Every keyword's outcome is written to run_results.csv as it happens,
    so a crash or Ctrl+C never loses progress.

Usage (from the Keyword_automation folder):
  python main_patchright.py                                 # uses keyword_list.xlsx
  python main_patchright.py MM-188_keyword_list_run_order.xlsx --out Downloads_MM188
  python main_patchright.py my_list.xlsx --limit 15         # quick test on 15 keywords
  python main_patchright.py my_list.xlsx --retry-empty      # also retry keywords that had no results before

The keyword file needs a column named "list" (same as before).
"""

import argparse
import csv
import os
import random
import re
import sys
import time
from datetime import datetime
from urllib.parse import quote

import pandas as pd
from patchright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ----------------------------------------------------------------------------
# Settings - change these, not the code below
# ----------------------------------------------------------------------------
FILTERS = {
    "dr_min": 30,            # Domain Rating from
    "dr_max": 90,            # Domain Rating to
    "traffic_min": 1000,     # website organic traffic from
    "published_days": 365,   # published in the last N days
    "words_min": 500,        # article length from N words
}
TITLE_WORDS = ["best", "top"]   # title must contain one of these AND the keyword

PROFILE_DIR = "patchright_profile"   # saved login lives here - do not share this folder
RESULTS_FILE = "run_results.csv"     # one line per keyword: ok / no_results / rate_limited / error
PAUSE_BETWEEN_KEYWORDS = (6, 12)     # random seconds between keywords
PAGE_TIMEOUT = 60                    # seconds to wait for results to show
DOWNLOAD_TIMEOUT = 120               # seconds to wait for the export file
RATE_LIMIT_COOLDOWN = (90, 180)      # random seconds to wait after a rate limit
MAX_ATTEMPTS = 3                     # tries per keyword for rate limits / slow pages
STOP_AFTER_ERRORS_IN_A_ROW = 8       # stop the run if this many keywords fail in a row

NO_RESULTS_MARKERS = ["no pages found", "no results found", "nothing found"]
RATE_LIMIT_MARKERS = ["rate limit exceeded", '"code":10029', "too many requests"]
BLOCK_MARKERS = ["verify you are human", "unusual traffic", "access denied"]
EXPORT_HEADER = "Content Title"
AUTH_URL_WORDS = ["login", "signin", "sign-in", "confirm", "verif", "2fa", "two-factor", "otp"]
BASE_URL = "https://app.ahrefs.com"
BROWSER_CHANNEL = "chrome"           # your installed Google Chrome; set to None to use Patchright's own Chromium


# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------
def log(msg):
    print(f"{datetime.now():%H:%M:%S}  {msg}", flush=True)


def safe_name(raw):
    """Same file naming as main_fixed.py, so existing exports are recognised."""
    return re.sub(r'[\\/:*?"<>|]+', '-', re.sub(r"\s+", "_", raw))


def build_url(keyword):
    words = "%20OR%20".join(f"%22{w}%22" for w in TITLE_WORDS)
    f = FILTERS
    return (
        f"{BASE_URL}/content-explorer/results/en/score_desc?field=title"
        f"&topic=({words})%20AND%20{quote(keyword, safe='')}"
        "&one-article-per-domain=1&exclude-homepages=1&exclude-subdomains=1"
        "&filter-broken-pages=all&publish_type=all"
        f"&filter-words={f['words_min']}-&published={f['published_days']}&first_seen=all"
        f"&filter-subdomain-orgtraffic={f['traffic_min']}-"
        f"&filter-domain_rating={f['dr_min']}-{f['dr_max']}"
        "&chart-range=last_6_months&page=1&results-per-page=50"
        "&report-languages=1%2CtotalTraffic%2Cdesc&report-websites=1%2CtotalTraffic%2Cdesc"
        "&report-authors=1%2CtotalTraffic%2Cdesc&safe-search=1"
    )


def is_valid_export(path):
    """True if the file is a real Content Explorer export (has the Content Title column)."""
    try:
        with open(path, "rb") as fh:
            head = fh.read(4000)
        enc = "utf-16" if head[:2] in (b"\xff\xfe", b"\xfe\xff") else "utf-8-sig"
        lines = head.decode(enc, errors="ignore").splitlines()
        return bool(lines) and EXPORT_HEADER in lines[0]
    except OSError:
        return False


def count_rows(path):
    try:
        with open(path, "rb") as fh:
            data = fh.read()
        enc = "utf-16" if data[:2] in (b"\xff\xfe", b"\xfe\xff") else "utf-8-sig"
        return max(len(data.decode(enc, errors="ignore").strip().splitlines()) - 1, 0)
    except OSError:
        return 0


def load_previous_results():
    """keyword -> last status from earlier runs."""
    last = {}
    if os.path.exists(RESULTS_FILE):
        with open(RESULTS_FILE, newline="", encoding="utf-8") as fh:
            for row in csv.DictReader(fh):
                last[row["keyword"]] = row["status"]
    return last


def record(keyword, status, rows="", note=""):
    new = not os.path.exists(RESULTS_FILE)
    with open(RESULTS_FILE, "a", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        if new:
            w.writerow(["time", "keyword", "status", "rows", "note"])
        w.writerow([datetime.now().isoformat(timespec="seconds"), keyword, status, rows, note])


def body_text(page):
    try:
        return page.locator("body").inner_text(timeout=3000).lower()
    except Exception:
        return ""


def page_state(page):
    """What the results page is showing right now."""
    text = body_text(page)
    if any(m in text for m in BLOCK_MARKERS):
        return "blocked"
    if any(m in text for m in RATE_LIMIT_MARKERS):
        return "rate_limited"
    if any(m in text for m in NO_RESULTS_MARKERS):
        return "no_results"
    try:
        if page.get_by_role("button", name="Export", exact=True).first.is_visible():
            return "ready"
    except Exception:
        pass
    if any(w in page.url.lower() for w in AUTH_URL_WORDS):
        return "logged_out"
    return "loading"


def wait_for_state(page, timeout):
    end = time.time() + timeout
    state = "loading"
    while time.time() < end:
        state = page_state(page)
        if state != "loading":
            return state
        time.sleep(1)
    return state


def is_logged_in(page):
    """Logged in = Content Explorer actually opens, not a login / email-confirmation page."""
    page.goto(f"{BASE_URL}/content-explorer", wait_until="domcontentloaded")
    time.sleep(5)
    url = page.url.lower()
    return "/content-explorer" in url and not any(w in url for w in AUTH_URL_WORDS)


def ensure_login(page):
    if is_logged_in(page):
        log("Already logged in to Ahrefs (saved session).")
        return
    page.goto(f"{BASE_URL}/user/login", wait_until="domcontentloaded")
    log("Please log in to Ahrefs in the browser window that opened.")
    log("If Ahrefs emails you a confirmation link or code: copy the LINK into THIS browser window's")
    log("address bar (or type the code here), not your normal Chrome - otherwise this window stays logged out.")
    for attempt in range(1, 4):
        input(">>> When you see the Ahrefs dashboard in this window, press Enter here... ")
        if is_logged_in(page):
            log("Logged in. The session is saved, next runs will skip this step.")
            return
        log(f"Not logged in yet (the page is still a login/confirmation page). Try again ({attempt}/3).")
    sys.exit("Still not logged in - stopping. Run the script again and log in.")


def export_one(page, keyword, out_path):
    """Returns (status, rows, note)."""
    url = build_url(keyword)
    state = "loading"
    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=PAGE_TIMEOUT * 1000)
        except PWTimeout:
            log(f"   page load timed out (attempt {attempt}/{MAX_ATTEMPTS})")
            continue

        state = wait_for_state(page, PAGE_TIMEOUT)
        if state == "no_results":
            return "no_results", 0, ""
        if state == "logged_out":
            return "logged_out", "", "session expired"
        if state in ("rate_limited", "blocked"):
            wait = random.uniform(*RATE_LIMIT_COOLDOWN) * (2 if state == "blocked" else 1)
            log(f"   {state} - cooling down {wait:.0f}s (attempt {attempt}/{MAX_ATTEMPTS})")
            time.sleep(wait)
            continue
        if state != "ready":
            log(f"   results did not load (attempt {attempt}/{MAX_ATTEMPTS})")
            continue

        try:
            page.get_by_role("button", name="Export", exact=True).first.click()
            time.sleep(random.uniform(1.5, 3))
            # work inside the export dialog only, so "All" never matches something else on the page
            dialogs = page.locator("[role=dialog]")
            scope = dialogs.last if dialogs.count() else page
            for option in ("One page per domain", "All"):
                opt = scope.get_by_text(option, exact=False).first
                if opt.count() and opt.is_visible():
                    opt.click()
                    time.sleep(random.uniform(0.5, 1.5))
                    break
            confirm = scope.get_by_role("button", name="Export", exact=True).last
            with page.expect_download(timeout=DOWNLOAD_TIMEOUT * 1000) as dl_info:
                confirm.click()
            download = dl_info.value
            tmp_path = out_path + ".part"
            download.save_as(tmp_path)
            if not is_valid_export(tmp_path):
                os.replace(tmp_path, out_path.replace(".csv", "_NOT_AN_EXPORT.csv"))
                return "error", "", "downloaded file is not a Content Explorer export"
            os.replace(tmp_path, out_path)
            return "ok", count_rows(out_path), ""
        except PWTimeout:
            log(f"   export dialog or download timed out (attempt {attempt}/{MAX_ATTEMPTS})")
            page.keyboard.press("Escape")
        except Exception as err:
            log(f"   export failed: {type(err).__name__} (attempt {attempt}/{MAX_ATTEMPTS})")
            page.keyboard.press("Escape")

    return ("rate_limited" if state in ("rate_limited", "blocked") else "error"), "", f"gave up, last state: {state}"


# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Ahrefs Content Explorer listicle exporter (Patchright)")
    ap.add_argument("keyword_file", nargs="?", default="keyword_list.xlsx")
    ap.add_argument("--out", default="Downloads", help="folder for the exported CSVs")
    ap.add_argument("--limit", type=int, default=0, help="only process the first N keywords (for testing)")
    ap.add_argument("--retry-empty", action="store_true", help="also retry keywords that had no results before")
    args = ap.parse_args()

    os.makedirs(args.out, exist_ok=True)
    keywords = [str(k).strip() for k in pd.read_excel(args.keyword_file)["list"].tolist()]
    keywords = [k for k in keywords if k and k.lower() != "nan"]
    keywords = list(dict.fromkeys(keywords))   # drop duplicates, keep order
    if args.limit:
        keywords = keywords[: args.limit]
    previous = load_previous_results()

    log(f"Keyword file: {args.keyword_file}  ({len(keywords)} keywords)  ->  {args.out}/")
    log("Filters: DR {dr_min}-{dr_max}, traffic {traffic_min}+, last {published_days} days, "
        "{words_min}+ words".format(**FILTERS))

    counts = {"ok": 0, "no_results": 0, "error": 0, "rate_limited": 0, "skipped": 0}
    errors_in_a_row = 0

    with sync_playwright() as p:
        opts = dict(user_data_dir=PROFILE_DIR, headless=False, no_viewport=True, accept_downloads=True)
        try:
            context = p.chromium.launch_persistent_context(channel=BROWSER_CHANNEL, **opts) \
                if BROWSER_CHANNEL else p.chromium.launch_persistent_context(**opts)
        except Exception as err:
            log(f"Could not start Google Chrome ({type(err).__name__}); trying Patchright's Chromium instead.")
            context = p.chromium.launch_persistent_context(**opts)
        page = context.pages[0] if context.pages else context.new_page()
        try:
            ensure_login(page)
            for i, kw in enumerate(keywords, 1):
                out_path = os.path.join(args.out, safe_name(kw) + ".csv")
                if os.path.exists(out_path) and is_valid_export(out_path):
                    counts["skipped"] += 1
                    continue
                if previous.get(kw) == "no_results" and not args.retry_empty:
                    counts["skipped"] += 1
                    continue

                log(f"[{i}/{len(keywords)}] {kw}")
                status, rows, note = export_one(page, kw, out_path)

                if status == "logged_out":
                    log("Ahrefs session expired - stopping. Run again and log in; finished keywords are kept.")
                    break

                record(kw, status, rows, note)
                counts[status] = counts.get(status, 0) + 1
                log(f"   -> {status}{f' ({rows} rows)' if status == 'ok' else ''}{f' - {note}' if note else ''}")

                errors_in_a_row = errors_in_a_row + 1 if status in ("error", "rate_limited") else 0
                if errors_in_a_row >= STOP_AFTER_ERRORS_IN_A_ROW:
                    log(f"{errors_in_a_row} failures in a row - stopping so the account is not hammered. "
                        "Check the browser, then run again (finished keywords are skipped).")
                    break

                time.sleep(random.uniform(*PAUSE_BETWEEN_KEYWORDS))
        except KeyboardInterrupt:
            log("Stopped by you (Ctrl+C). Progress is saved; run again to continue.")
        finally:
            context.close()

    log("Done. ok={ok}  no_results={no_results}  errors={error}  rate_limited={rate_limited}  "
        "skipped={skipped}".format(**counts))
    # Rerun list = real failures only (not empty keywords)
    last = load_previous_results()
    retry = [k for k in keywords if last.get(k) in ("error", "rate_limited")]
    if retry:
        pd.DataFrame({"list": retry}).to_excel("retry_keywords.xlsx", index=False)
        log(f"{len(retry)} keywords to retry -> retry_keywords.xlsx "
            f"(run: python main_patchright.py retry_keywords.xlsx --out {args.out})")


if __name__ == "__main__":
    main()
