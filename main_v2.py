"""
Ahrefs Content Explorer bulk exporter — unattended edition.

Key improvements over the original main.py:
  * Cookie persistence: log in by hand ONCE, session is reused every run
    (far less likely to be flagged than logging in each time).
  * Auto-retry on the "Failed / Reload" state — no human needed.
  * Block/captcha detection -> long randomized cooldown, then auto-resume.
  * Resume: keywords already downloaded are skipped, so a crash/block
    never makes you redo work.
  * Randomized human-like delays to reduce detection.
  * Real logging to run.log so you can see what happened while away.
  * Explicit Chrome download dir + proper "wait for download to finish".

Usage:
  1. First run: a browser opens on the Ahrefs login page. Log in manually
     (solve any 2FA/captcha). Then press Enter in the console ONCE.
     Cookies are saved to ahrefs_cookies.pkl.
  2. Every run after that reuses the cookies — fully unattended.
     If cookies expire, it pauses and asks you to log in again.
"""

import os
import re
import glob
import time
import random
import pickle
import shutil
import logging
from datetime import datetime
import pandas as pd
import undetected_chromedriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ----------------------------------------------------------------------------
# Config — tweak these
# ----------------------------------------------------------------------------
COOKIES_FILE      = "ahrefs_cookies.pkl"
LOGIN_FILE        = "ahrefs_login.txt"   # line 1 = email, line 2 = password
KEYWORD_FILE      = "keyword_list_new.xlsx"   # client-tagged list (cols: list, client)
KEYWORD_COLUMN    = "list"

# Proxy — leave "" (empty) to use your normal connection (default, works as-is).
# To use a proxy later, paste the string from your provider, e.g.:
#   PROXY = "http://username:password@host:port"
PROXY = ""
CHROME_VERSION    = 151         # set to your Chrome's major version (Help > About Google Chrome)
MAX_RELOAD_RETRIES = 4          # how many times to click Reload/retry before giving up on a keyword
BLOCK_COOLDOWN_SEC = (600, 1200)  # random cooldown range (10-20 min) when a block is detected
PER_KEYWORD_PAUSE  = (8, 20)      # random pause between keywords (human-like pacing)

# Text that means the query legitimately returned nothing — skip, don't retry.
NO_RESULTS_MARKERS = ["No pages found", "No results found", "Nothing found"]
# Text that indicates a transient failure with a reload/retry button.
RELOAD_MARKERS = ["Reload", "Try again", "Retry", "Something went wrong"]
# Text that indicates you've been rate-limited / blocked / challenged.
BLOCK_MARKERS  = ["unusual traffic", "verify you are human", "captcha",
                  "rate limit", "too many requests", "access denied", "cf-challenge"]

# ----------------------------------------------------------------------------
# Logging
# ----------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    handlers=[logging.FileHandler("run.log", encoding="utf-8"),
              logging.StreamHandler()],
)
log = logging.getLogger("ahrefs")


def rand(lo_hi):
    """Randomized sleep from a (low, high) tuple — makes pacing look human."""
    lo, hi = lo_hi
    time.sleep(random.uniform(lo, hi))


# ----------------------------------------------------------------------------
# Downloads helpers
# ----------------------------------------------------------------------------
def wait_for_download(download_dir, timeout=90):
    """Wait until a new .csv finishes downloading (no .crdownload left)."""
    end = time.time() + timeout
    while time.time() < end:
        crdownloads = glob.glob(os.path.join(download_dir, "*.crdownload"))
        csvs = glob.glob(os.path.join(download_dir, "*.csv"))
        if csvs and not crdownloads:
            return max(csvs, key=os.path.getctime)
        time.sleep(1)
    return None


def rename_csv(download_dir, dest_dir, keyword):
    src = wait_for_download(download_dir)
    if not src:
        log.warning("No CSV appeared for keyword: %s", keyword)
        return False
    try:
        shutil.move(src, os.path.join(dest_dir, keyword + ".csv"))
        return True
    except Exception as err:
        log.warning("Rename failed for %s: %s", keyword, err)
        return False


# ----------------------------------------------------------------------------
# Login / cookies
# ----------------------------------------------------------------------------
def save_cookies(driver):
    with open(COOKIES_FILE, "wb") as f:
        pickle.dump(driver.get_cookies(), f)
    log.info("Saved %d cookies to %s", len(driver.get_cookies()), COOKIES_FILE)


def load_cookies(driver):
    # Missing or empty/corrupt file -> behave as "no cookies" instead of crashing.
    if not os.path.exists(COOKIES_FILE) or os.path.getsize(COOKIES_FILE) == 0:
        return False
    try:
        with open(COOKIES_FILE, "rb") as f:
            cookies = pickle.load(f)
    except (EOFError, pickle.UnpicklingError):
        log.warning("Cookie file was empty/corrupt — ignoring it and logging in fresh.")
        return False
    if not cookies:
        return False
    driver.get("https://app.ahrefs.com/")
    rand((2, 4))
    for c in cookies:
        c.pop("sameSite", None)  # selenium chokes on some sameSite values
        try:
            driver.add_cookie(c)
        except Exception:
            pass
    return True


def is_logged_in(driver):
    driver.get("https://app.ahrefs.com/dashboard")
    rand((3, 6))
    return "login" not in driver.current_url.lower()


def read_credentials():
    """Read email (line 1) and password (line 2) from LOGIN_FILE."""
    with open(LOGIN_FILE, "r", encoding="utf-8-sig") as f:
        lines = [ln.strip() for ln in f if ln.strip()]
    if len(lines) < 2:
        raise RuntimeError(f"{LOGIN_FILE} must have email on line 1 and password on line 2.")
    return lines[0], lines[1]


def ensure_login(driver):
    """Reuse saved cookies; if invalid, auto-fill credentials then pause for verification."""
    if load_cookies(driver) and is_logged_in(driver):
        log.info("Reused saved session — logged in.")
        return

    log.info("No valid session. Auto-filling login from %s.", LOGIN_FILE)
    email, password = read_credentials()

    driver.get("https://app.ahrefs.com/user/login")
    driver.maximize_window()
    rand((8, 12))  # let the form render

    try:
        email_field = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.NAME, "email"))
        )
        password_field = driver.find_element(By.NAME, "password")
        email_field.clear()
        email_field.send_keys(email)
        rand((1, 2))
        password_field.clear()
        password_field.send_keys(password)
        rand((1, 2))
        driver.find_element(By.XPATH, "//button[@type='submit']").click()
        log.info("Submitted login for %s.", email)
    except Exception as err:
        log.warning("Auto-fill hit a snag (%s). Fill it in the browser by hand.",
                    type(err).__name__)

    # Pause so you can open the email verification link / finish 2FA, then press Enter.
    print("\n" + "=" * 70)
    print(">>> If Ahrefs asks for email verification or 2FA:")
    print(">>>   open the verification link, complete it in this browser,")
    print(">>>   then come back here and press Enter to start scraping.")
    print("=" * 70)
    input(">>> Press Enter once you're logged in... ")

    if is_logged_in(driver):
        save_cookies(driver)
        log.info("Login captured and saved — future runs skip this step.")
    else:
        raise RuntimeError("Still not logged in after verification step — aborting.")


# ----------------------------------------------------------------------------
# Page-state detection  (based on VISIBLE elements/text, not raw page source —
# the app's JS bundle contains words like "Reload" even on a healthy page)
# ----------------------------------------------------------------------------
def visible_text(driver):
    try:
        return driver.find_element(By.TAG_NAME, "body").text.lower()
    except Exception:
        return ""


def is_blocked(driver):
    """True only if a block/captcha message is actually shown to the user."""
    text = visible_text(driver)
    title = (driver.title or "").lower()
    return any(m.lower() in text or m.lower() in title for m in BLOCK_MARKERS)


def has_no_results(driver):
    """True if Ahrefs shows an empty-result state ('No pages found')."""
    text = visible_text(driver)
    return any(m.lower() in text for m in NO_RESULTS_MARKERS)


def find_reload_button(driver):
    """Return a VISIBLE reload/retry button element, or None. Ignores hidden matches."""
    for marker in RELOAD_MARKERS:
        try:
            els = driver.find_elements(
                By.XPATH,
                f"//button[contains(normalize-space(.), '{marker}')] | "
                f"//a[contains(normalize-space(.), '{marker}')]"
            )
            for el in els:
                if el.is_displayed():
                    return el
        except Exception:
            continue
    return None


def cooldown_for_block(keyword):
    wait = random.uniform(*BLOCK_COOLDOWN_SEC)
    log.warning("BLOCK detected on '%s'. Cooling down %.0f min, then resuming.",
                keyword, wait / 60)
    time.sleep(wait)


# ----------------------------------------------------------------------------
# Export one keyword (with retry + block handling)
# ----------------------------------------------------------------------------
def export_keyword(driver, keyword, url_encoded, download_dir, dest_dir):
    # Date range: 2024-01-01 through today (end date auto-updates each run).
    start_pub   = "2024-01-01T08%3A00%3A00Z"
    start_seen  = "2024-01-01T20%3A00%3A00Z"
    today_pub   = datetime.now().strftime("%Y-%m-%d") + "T08%3A00%3A00Z"
    today_seen  = datetime.now().strftime("%Y-%m-%d") + "T20%3A00%3A00Z"

    web_page = (
        "https://app.ahrefs.com/content-explorer/results/en/score_desc?"
        "field=title&topic=title%3A(%22best%22%20OR%20%22top%22%20OR%20%22alternatives%22)"
        "%20AND%20title%3A%22" + url_encoded + "%22%20-title%3A(%22vs%22%20OR%20%22versus%22)"
        "&one-article-per-domain=1&filter-broken-pages=live&publish_type=all&filter-words=600-"
        "&published=" + start_pub + "%2C" + today_pub +
        "&first_seen=" + start_seen + "%2C" + today_seen +
        "&filter-subdomain-orgtraffic=1000-&filter-domain_rating=40-89"
        "&chart-range=last_6_months&page=1&results-per-page=50"
        "&report-languages=1%2CtotalTraffic%2Cdesc&report-websites=1%2CtotalTraffic%2Cdesc"
        "&report-authors=1%2CtotalTraffic%2Cdesc&safe-search=1"
    )

    for attempt in range(1, MAX_RELOAD_RETRIES + 1):
        driver.get(web_page)
        rand((3, 6))

        # 1) Blocked? Cool down and retry the same keyword.
        if is_blocked(driver):
            cooldown_for_block(keyword)
            continue

        # 1b) Legitimately no results ("No pages found") — skip fast, no retries.
        if has_no_results(driver):
            log.info("SKIP  '%s' — no pages found (nothing to export).", keyword)
            return False

        # 2) Wait for the Export button — that's the real "results loaded" signal.
        #    If it never appears, THEN look for a visible reload button and click it.
        try:
            export_btn = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Export']"))
            )
        except Exception:
            reload_btn = find_reload_button(driver)
            if reload_btn:
                log.info("Attempt %d/%d: real failed/reload panel on '%s' — clicking it.",
                         attempt, MAX_RELOAD_RETRIES, keyword)
                try:
                    reload_btn.click()
                except Exception:
                    pass
                rand((5, 10))
            else:
                log.info("Attempt %d/%d: no Export button and no reload on '%s' "
                         "(no results, or slow load) — retrying.",
                         attempt, MAX_RELOAD_RETRIES, keyword)
                rand((4, 8))
            continue

        # 3) Export button is present — run the export flow.
        try:
            export_btn.click()
            rand((3, 6))

            # dialog radio option (one-per-domain preferred, else All)
            for opt in ("One page per domain", "All"):
                try:
                    driver.find_element(By.XPATH, f"//span[contains(text(), '{opt}')]").click()
                    rand((1, 2))
                    break
                except Exception:
                    continue

            # confirm export (2nd Export button in the dialog)
            driver.find_elements(By.XPATH, "//button[normalize-space()='Export']")[1].click()
            rand((6, 12))

            if rename_csv(download_dir, dest_dir, keyword):
                log.info("OK  '%s' exported (attempt %d).", keyword, attempt)
                return True
            log.info("Export clicked but no file for '%s', retrying.", keyword)

        except Exception as err:
            log.info("Attempt %d/%d failed for '%s': %s",
                     attempt, MAX_RELOAD_RETRIES, keyword, type(err).__name__)
            rand((4, 8))

    log.error("GAVE UP on '%s' after %d attempts.", keyword, MAX_RELOAD_RETRIES)
    return False


# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------
def main():
    parent_dir     = os.getcwd()
    download_dir   = os.path.join(parent_dir, "Downloads")  # Chrome downloads here
    dest_dir       = os.path.join(parent_dir, "Temp")       # renamed CSVs land here

    # Do NOT wipe dest_dir — we want resume. Only ensure dirs exist.
    os.makedirs(download_dir, exist_ok=True)
    os.makedirs(dest_dir, exist_ok=True)

    # Which keywords are already done? (resume support)
    done = {os.path.splitext(f)[0] for f in os.listdir(dest_dir) if f.endswith(".csv")}
    log.info("%d keywords already downloaded — will skip them.", len(done))

    # Chrome with an explicit, isolated download directory.
    opts = undetected_chromedriver.ChromeOptions()
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "profile.default_content_settings.popups": 0,
    }
    opts.add_experimental_option("prefs", prefs)
    if PROXY:
        opts.add_argument(f"--proxy-server={PROXY}")
        log.info("Using proxy: %s", PROXY.split("@")[-1])  # log host only, not credentials
    else:
        log.info("No proxy — using normal connection.")
    driver = undetected_chromedriver.Chrome(options=opts, version_main=CHROME_VERSION)

    try:
        ensure_login(driver)

        keyw = pd.read_excel(KEYWORD_FILE)
        key_list = keyw[KEYWORD_COLUMN].dropna().tolist()
        log.info("Loaded %d keywords.", len(key_list))

        for i, raw in enumerate(key_list, 1):
            keyword_safe = re.sub(r"\s+", "_", str(raw))   # for filename
            if keyword_safe in done:
                log.info("[%d/%d] skip (done): %s", i, len(key_list), keyword_safe)
                continue
            url_encoded = re.sub(r"\s+", "%20", str(raw))  # for the URL
            log.info("[%d/%d] processing: %s", i, len(key_list), keyword_safe)

            export_keyword(driver, keyword_safe, url_encoded, download_dir, dest_dir)

            # Re-save cookies occasionally so a long run keeps the session fresh.
            if i % 25 == 0:
                save_cookies(driver)

            rand(PER_KEYWORD_PAUSE)  # human-like gap between keywords

    finally:
        try:
            save_cookies(driver)
        except Exception:
            pass
        driver.quit()
        log.info("Done. See run.log for the full record.")


if __name__ == "__main__":
    main()
