import os
import random
import re
import json
import sys
import traceback
import urllib.parse
import urllib.request
import uuid
from datetime import datetime
from urllib.parse import quote

import pandas as pd
from patchright.sync_api import sync_playwright

# ---------------- Settings ----------------
KEYWORD_FILE = "keyword_list.xlsx"   # Excel file with a column called "list"
OUT_FOLDER = "Downloads_maintest"    # CSV exports go here, one per keyword
LIMIT = None                         # how many keywords to run (None = all, 3 = quick test)
BREAK_EVERY = 50                     # take a long break after this many keywords...
BREAK_SECONDS = (120, 240)           # ...lasting 2-4 minutes (random)
STOP_AFTER_FAILS = 8                 # stop if this many keywords fail in a row (rate limit / logout)
TELEGRAM_SETTINGS_FILE = "telegram_settings.txt"   # your bot token (see the file for setup)

# Armen's filters: DR 30-90, traffic 1000+, last 365 days, 500+ words
DR_MIN, DR_MAX = 30, 90
TRAFFIC_MIN = 1000
PUBLISHED_DAYS = 365
WORDS_MIN = 500

# Text on the page that means Ahrefs is suspicious of us or limiting us
SUSPICIOUS_TEXT = ["verify you are human", "unusual traffic", "access denied",
                   "rate limit exceeded", "too many requests"]


class StopRun(Exception):
    """Raised when the run should stop on purpose (suspicious activity, logout, too many failures)."""


# ---------------- Telegram reports ----------------
def read_telegram_settings():
    """Reads telegram_settings.txt: lines like 'bot_token = 123:ABC'. Returns a dict (may be incomplete)."""
    settings = {}
    if os.path.exists(TELEGRAM_SETTINGS_FILE):
        for line in open(TELEGRAM_SETTINGS_FILE, encoding="utf-8"):
            if "=" in line and not line.strip().startswith("#"):
                key, value = line.split("=", 1)
                settings[key.strip()] = value.strip()
    return settings


def telegram_api(token, method, fields=None, photo=None):
    """Calls the Telegram Bot API. Uses only Python's built-in urllib, so nothing extra to install."""
    url = f"https://api.telegram.org/bot{token}/{method}"
    fields = fields or {}
    if photo:   # sending a file needs a 'multipart' request
        boundary = uuid.uuid4().hex
        body = b""
        for name, value in fields.items():
            body += (f"--{boundary}\r\nContent-Disposition: form-data; name=\"{name}\"\r\n\r\n"
                     f"{value}\r\n").encode()
        with open(photo, "rb") as f:
            body += (f"--{boundary}\r\nContent-Disposition: form-data; name=\"photo\"; "
                     f"filename=\"last_screen.png\"\r\nContent-Type: image/png\r\n\r\n").encode()
            body += f.read() + f"\r\n--{boundary}--\r\n".encode()
        headers = {"Content-Type": f"multipart/form-data; boundary={boundary}"}
    else:
        body = urllib.parse.urlencode(fields).encode()
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
    request = urllib.request.Request(url, data=body, headers=headers)
    with urllib.request.urlopen(request, timeout=30) as response:
        return json.loads(response.read())


def find_chat_id(token):
    """After you press Start in your bot, Telegram knows your chat id - fetch it and save it to the settings file."""
    updates = telegram_api(token, "getUpdates").get("result", [])
    for update in reversed(updates):
        chat = (update.get("message") or {}).get("chat")
        if chat:
            chat_id = str(chat["id"])
            text = open(TELEGRAM_SETTINGS_FILE, encoding="utf-8").read()
            text = re.sub(r"^chat_id *=.*$", f"chat_id = {chat_id}", text, flags=re.M)
            open(TELEGRAM_SETTINGS_FILE, "w", encoding="utf-8").write(text)
            return chat_id
    return None


def list_chats(token):
    """Prints everyone who has written to the bot recently, with their chat id."""
    updates = telegram_api(token, "getUpdates").get("result", [])
    seen = {}
    for update in updates:
        chat = (update.get("message") or {}).get("chat")
        if chat:
            name = " ".join(x for x in (chat.get("first_name"), chat.get("last_name")) if x) or chat.get("title", "")
            username = f"@{chat['username']}" if chat.get("username") else ""
            seen[chat["id"]] = f"{name} {username}".strip()
    if not seen:
        print("Nobody has written to the bot recently. Ask the person to open the bot, press Start and send 'hi'.")
    for chat_id, who in seen.items():
        print(f"{chat_id}   {who}")


def send_report(subject, body, screenshot=None):
    """Sends the report to your Telegram. If it fails, just prints why - it never stops the script."""
    settings = read_telegram_settings()
    token = settings.get("bot_token")
    if not token:
        print(f"(No Telegram report sent - fill in {TELEGRAM_SETTINGS_FILE} to get reports.)")
        return
    try:
        chat_ids_text = settings.get("chat_id") or find_chat_id(token)
        if not chat_ids_text:
            print("Telegram: open your bot in Telegram and press Start (or send it any message), then try again.")
            return
        chat_ids = [c.strip() for c in chat_ids_text.split(",") if c.strip()]   # several people: comma-separated
        text = f"{subject}\n\n{body}"
        if len(text) > 4000:                       # Telegram's limit is 4096 characters per message
            text = text[:3900] + "\n\n... (cut - see the terminal for the full text)"
        for chat_id in chat_ids:
            try:
                telegram_api(token, "sendMessage", {"chat_id": chat_id, "text": text})
                if screenshot and os.path.exists(screenshot):
                    telegram_api(token, "sendPhoto", {"chat_id": chat_id, "caption": "Browser when it stopped"},
                                 photo=screenshot)
                print(f"Report sent to Telegram chat {chat_id}.")
            except Exception as error:      # one person failing (e.g. never pressed Start) doesn't stop the others
                print(f"Could not send to Telegram chat {chat_id}: {error}")
    except Exception as error:
        print(f"Could not send the Telegram report: {error}")


# ---------------- Ahrefs helpers ----------------
def build_url(keyword):
    """The Content Explorer search URL: title has ("best" OR "top") AND the keyword, plus filters."""
    return (
        "https://app.ahrefs.com/content-explorer/results/en/score_desc?field=title"
        f"&topic=(%22best%22%20OR%20%22top%22)%20AND%20{quote(keyword, safe='')}"
        "&one-article-per-domain=1&exclude-homepages=1&exclude-subdomains=1"
        "&filter-broken-pages=all&publish_type=all"
        f"&filter-words={WORDS_MIN}-&published={PUBLISHED_DAYS}&first_seen=all"
        f"&filter-subdomain-orgtraffic={TRAFFIC_MIN}-"
        f"&filter-domain_rating={DR_MIN}-{DR_MAX}"
        "&chart-range=last_6_months&page=1&results-per-page=50&safe-search=1"
    )


def file_name(keyword):
    """'eor providers' -> 'eor_providers.csv' (no characters Windows can't use in file names)."""
    return re.sub(r'[\\/:*?"<>|]+', "-", re.sub(r"\s+", "_", keyword)) + ".csv"


def wait_for_results(page, seconds=60):
    """Returns 'ready', 'empty', 'suspicious', 'logged_out' or 'timeout'."""
    export_button = page.get_by_role("button", name="Export", exact=True).first
    for i in range(seconds):
        page.wait_for_timeout(1000)
        text = page.locator("body").inner_text().lower()
        if any(words in text for words in SUSPICIOUS_TEXT):
            return "suspicious"
        if "login" in page.url or "verify" in page.url:
            return "logged_out"
        if "no pages found" in text or "no results found" in text or "nothing found" in text:
            return "empty"
        if export_button.is_visible():
            return "ready"
    return "timeout"


def export_csv(page, save_path):
    """Click Export, choose options in the export window, and save the download."""
    page.get_by_role("button", name="Export", exact=True).first.click()
    page.wait_for_timeout(2000)
    dialog = page.locator("[role=dialog]").last          # the export pop-up window
    for option in ["One page per domain", "All"]:
        choice = dialog.get_by_text(option).first
        if choice.count() and choice.is_visible():
            choice.click()
            page.wait_for_timeout(1000)
            break
    with page.expect_download(timeout=60000) as download_info:
        dialog.get_by_role("button", name="Export", exact=True).last.click()
    download_info.value.save_as(save_path)


def log_in(page):
    page.goto("https://app.ahrefs.com/user/login")
    print("Log in to Ahrefs in the browser window.")
    print("Confirmation email? Right-click the link -> Copy link address -> paste it in THIS window.")
    # page.wait_for_timeout keeps the browser working while we wait (input() / time.sleep would freeze it)
    for check in range(900):                 # 900 x 2 seconds = 30 minutes
        page.wait_for_timeout(2000)
        if "/dashboard" in page.url:
            print("Logged in! Starting keywords.")
            return
    raise StopRun("Didn't reach the Ahrefs dashboard within 30 minutes of opening the login page.")


def run_keywords(page, keywords, counts):
    """One search + export per keyword. Updates `counts` as it goes."""
    fails_in_a_row = 0     # resets to 0 whenever a keyword works
    searched = 0           # keywords actually searched this run (skipped ones don't count)

    for number, keyword in enumerate(keywords, start=1):
        tag = f"[{number}/{len(keywords)}] {keyword}"
        save_path = os.path.join(OUT_FOLDER, file_name(keyword))
        if os.path.exists(save_path):
            print(f"{tag}: already done, skipping")
            counts["skipped"] += 1
            continue

        page.goto(build_url(keyword))
        result = wait_for_results(page)
        searched += 1
        counts["last_keyword"] = keyword
        worked = False

        if result == "suspicious":
            raise StopRun(f"Ahrefs showed a warning (rate limit / 'verify you are human') at keyword '{keyword}'.")
        if result == "logged_out":
            raise StopRun(f"Ahrefs logged you out at keyword '{keyword}'.")

        if result == "ready":
            # Ahrefs sometimes says "Something went wrong" on export - try up to 3 times
            for attempt in range(1, 4):
                try:
                    export_csv(page, save_path)
                    print(f"{tag}: saved {save_path}")
                    counts["saved"] += 1
                    worked = True
                    break
                except Exception as error:
                    print(f"{tag}: export failed ({type(error).__name__}), try {attempt}/3")
                    page.keyboard.press("Escape")
                    page.wait_for_timeout(5000)
                    if attempt < 3:
                        page.goto(build_url(keyword))      # reload the results and try again
                        if wait_for_results(page) != "ready":
                            break
        elif result == "empty":
            print(f"{tag}: no results")
            counts["no_results"] += 1
            worked = True          # the search itself worked, there's just nothing to export
        else:
            print(f"{tag}: page didn't load in 60 seconds")

        if not worked and result != "empty":
            counts["failed"].append(keyword)

        # Safety stop: many failures in a row usually means a rate limit or a logout
        fails_in_a_row = 0 if worked else fails_in_a_row + 1
        if fails_in_a_row >= STOP_AFTER_FAILS:
            raise StopRun(f"{STOP_AFTER_FAILS} keywords failed in a row (last one: '{keyword}').")

        # Long break every BREAK_EVERY keywords, short random pause otherwise
        if searched % BREAK_EVERY == 0:
            pause = random.randint(*BREAK_SECONDS)
            print(f"Done {searched} keywords this run - taking a {pause // 60} min {pause % 60} s break.")
            page.wait_for_timeout(pause * 1000)
        else:
            page.wait_for_timeout(random.randint(6000, 12000))   # random pause so Ahrefs doesn't see a bot rhythm


def summary(counts, keywords, started):
    minutes = int((datetime.now() - started).total_seconds() // 60)
    failed = counts["failed"]
    lines = [
        f"Keyword file: {KEYWORD_FILE} ({len(keywords)} keywords)",
        f"Started: {started:%Y-%m-%d %H:%M}   Duration: {minutes // 60} h {minutes % 60} min",
        "",
        f"Saved:            {counts['saved']}",
        f"No results:       {counts['no_results']}",
        f"Failed:           {len(failed)}",
        f"Already done:     {counts['skipped']}",
        f"Last keyword:     {counts['last_keyword']}",
        f"Exports folder:   {os.path.abspath(OUT_FOLDER)}",
    ]
    if failed:
        lines += ["", "Failed keywords (run again to retry them):"] + [f"  - {k}" for k in failed[:50]]
        if len(failed) > 50:
            lines.append(f"  ... and {len(failed) - 50} more")
    return "\n".join(lines)


# ---------------- Main ----------------
if __name__ == "__main__":
    if "--telegram-chats" in sys.argv:
        list_chats(read_telegram_settings().get("bot_token", ""))
        raise SystemExit

    if "--test-telegram" in sys.argv:
        send_report("Ahrefs keyword automation: test", "If you can read this, Telegram reports work.")
        raise SystemExit

    started = datetime.now()
    keywords = pd.read_excel(KEYWORD_FILE)["list"].dropna().astype(str).str.strip().tolist()
    limit = LIMIT
    if "--limit" in sys.argv:                       # e.g.  python maintest.py --limit 1
        limit = int(sys.argv[sys.argv.index("--limit") + 1])
    if limit:
        # count only keywords that aren't done yet, so "--limit 1" always searches one new keyword
        not_done = [k for k in keywords if not os.path.exists(os.path.join(OUT_FOLDER, file_name(k)))]
        keywords = not_done[:limit]
    os.makedirs(OUT_FOLDER, exist_ok=True)
    print(f"{len(keywords)} keywords from {KEYWORD_FILE}")

    counts = {"saved": 0, "no_results": 0, "skipped": 0, "failed": [], "last_keyword": "-"}
    status, details, screenshot = "Finished", "", None

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir="my_profile",
            channel="chrome",
            headless=False,
            no_viewport=True,
            accept_downloads=True,
        )
        page = context.pages[0]
        try:
            log_in(page)
            run_keywords(page, keywords, counts)
        except StopRun as reason:                 # stopped on purpose: suspicious activity, logout, failures
            status, details = "STOPPED", str(reason)
        except KeyboardInterrupt:                 # you pressed Ctrl+C
            status, details = "Stopped by you", "You pressed Ctrl+C. Run again to continue."
        except Exception:                         # anything unexpected = crash
            status, details = "CRASHED", traceback.format_exc()

        if status != "Finished":
            try:                                  # picture of what the browser showed when it stopped
                screenshot = "last_screen.png"
                page.screenshot(path=screenshot)
            except Exception:
                screenshot = None

        report = summary(counts, keywords, started)
        if details:
            report = f"Reason:\n{details}\n\n{report}"
        print("\n" + report)
        send_report(f"Ahrefs keyword automation: {status}", report, screenshot)

        try:
            context.close()
        except Exception:
            pass
