# Ahrefs Content Explorer keyword automation

Searches Ahrefs Content Explorer for listicles ("best"/"top" + keyword in the title)
for every keyword in an Excel file and exports each result as a CSV.

## Main script: `maintest.py` (Patchright)

```
pip install -r requirements.txt
python maintest.py              # all keywords in keyword_list.xlsx (column "list")
python maintest.py --limit 1    # quick test: the first keyword not done yet
python maintest.py --test-telegram     # test the Telegram report
python maintest.py --telegram-chats    # list chat ids that wrote to the bot
```

- Log in to Ahrefs by hand in the browser window it opens; the run starts once the dashboard shows.
- Exports go to `Downloads_maintest/<keyword>.csv`. Keywords that already have a CSV are skipped, so a re-run continues where it stopped.
- Filters (DR 30-90, traffic 1000+, last 365 days, 500+ words), pauses, breaks every 50 keywords and the stop after 8 failures in a row are set at the top of the file.
- A Telegram report is sent when the run finishes, stops or crashes. Needs `telegram_settings.txt` (not in git):

```
bot_token = <token from @BotFather>
chat_id = <your chat id>, <more chat ids>
```

## Other scripts

`main_patchright.py`, `main_fixed.py`, `main_v2.py`, `main.py` are earlier versions; `fetch_company_keywords.py` and `domain_position.py` are helper tools.

Secrets (`telegram_settings.txt`, `ahrefs_login.txt`, cookies, browser profiles) and data (downloads, keyword spreadsheets) are excluded by `.gitignore`.
