import os
import re
import time
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import undetected_chromedriver

# List of domains
domains = [
    "5paisa.com", "alp.consulting", "authoritylighthouse.com", "coursiv.com",
    "databridgemarketresearch.com", "dofollow.com", "finnpartners.com", "futurice.com",
    "gameboost.com", "getmelinks.com", "intelivita.co.uk", "kotakneo.com",
    "lazard.com", "linecorp.com", "loox.app", "lyzr.ai",
    "marknteladvisors.com", "matchoffice.com", "oanda.com", "perfios.ai",
    "perimattic.com", "poonawallafincorp.com", "rankviz.co.uk", "reputation-score.netreputation.com",
    "savvylawfirmmarketing.com", "shopify.ca", "socialinsider.io", "sortlist.co.uk",
    "theblueground.com", "tkxel.com", "uknowva.com", "uptodate.com",
    "vegavid.com", "venveo.com", "visualdx.com", "voka.io", "zintego.com"
]

def extract_keywords_from_page(driver):
    """Extract keywords from Ahrefs organic keywords table"""
    keywords = []
    try:
        # Wait for table to load
        WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, "//tr[@data-row-key]"))
        )
        time.sleep(2)

        # Get all keyword rows
        rows = driver.find_elements(By.XPATH, "//tr[@data-row-key]")

        for row in rows[:50]:  # Get top 50 keywords
            try:
                # Get keyword text (usually first td)
                cells = row.find_elements(By.TAG_NAME, "td")
                if cells:
                    keyword = cells[0].text.strip()
                    if keyword:
                        keywords.append(keyword)
            except:
                continue

        return keywords
    except:
        return keywords

def main():
    driver = undetected_chromedriver.Chrome(version_main=151)
    driver.maximize_window()

    print("Log in to Ahrefs manually...")
    driver.get("https://app.ahrefs.com/user/login")

    # Wait for manual login
    input("Press Enter after you've logged in...")

    results = []

    for i, domain in enumerate(domains, 1):
        print(f"[{i}/{len(domains)}] Fetching keywords for {domain}...")

        try:
            # Navigate to Site Explorer > Organic Keywords
            url = f"https://app.ahrefs.com/site-explorer/overview?target={domain}"
            driver.get(url)
            time.sleep(3)

            # Click on Organic Keywords section
            try:
                organic_kw_btn = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, 'organic-keywords')]"))
                )
                organic_kw_btn.click()
                time.sleep(3)
            except:
                pass

            # Extract keywords from current page
            keywords = extract_keywords_from_page(driver)

            results.append({
                'Domain': domain,
                'Top Keywords': ' | '.join(keywords[:20]) if keywords else 'N/A',
                'Keyword Count': len(keywords)
            })

            time.sleep(1)

        except Exception as e:
            print(f"Error fetching {domain}: {str(e)}")
            results.append({
                'Domain': domain,
                'Top Keywords': 'Error',
                'Keyword Count': 0
            })
            time.sleep(2)

    driver.quit()

    # Save to CSV
    df = pd.DataFrame(results)
    output_path = os.path.join(os.path.dirname(__file__), "company_keywords_ahrefs.xlsx")

    # Write to Excel with formatting
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Keywords')

        # Format
        workbook = writer.book
        worksheet = writer.sheets['Keywords']

        for col in worksheet.columns:
            worksheet.column_dimensions[col[0].column_letter].width = 50

    print(f"\nDone! Saved to {output_path}")

if __name__ == '__main__':
    main()
