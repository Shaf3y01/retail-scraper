from selenium import webdriver 
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
import os
import re
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

# === Chrome Setup ===
options = Options()
options.add_argument("--headless=new")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36")

# === Input: Excel with categories and URLs ===
input_excel = "2b-target-links.xlsx"
df = pd.read_excel(input_excel, header=1)
df.columns = df.columns.str.strip()
df = df.dropna(subset=["Category", "URL"])
category_links = list(zip(df["Category"], df["URL"]))

# === Output Directory ===
output_dir = "2b-products"
os.makedirs(output_dir, exist_ok=True)

# === Excel Styling ===
def style_excel_file(path):
    wb = load_workbook(path)
    ws = wb.active

    header_fill = PatternFill(start_color="003366", end_color="003366", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    body_font = Font(color="000000")
    center_align = Alignment(horizontal="center", vertical="center")
    border = Border(bottom=Side(border_style="thin", color="000000"))

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = center_align
            cell.border = border
            if cell.row == 1:
                cell.fill = header_fill
                cell.font = header_font
            else:
                cell.font = body_font

    for column in ws.columns:
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in column)
        ws.column_dimensions[column[0].column_letter].width = max_length + 2

    wb.save(path)

# === Helpers ===
def normalize_price(text):
    if not text:
        return None
    return int(re.sub(r"[^\d]", "", text))

def extract_sku(name):
    if not name:
        return ""
    upper_name = name.upper()
    pattern = r'\b(?:[A-Z0-9]{1,10})(?:[ .\-_]{0,1}[A-Z0-9]{1,10}){0,2}\b'
    matches = re.findall(pattern, upper_name)
    for candidate in reversed(matches):
        if any(c.isalpha() for c in candidate) and len(candidate.replace(" ", "")) >= 2:
            return candidate.strip()
    return ""

def normalize_sku(sku):
    return re.sub(r'[-_\s]', '', sku).lower() if sku else ""

# === Start Browser ===
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 10)

# === Main Loop ===
print("🚀 Starting 2B Scraper...")
for category, url in category_links:
    print(f"\n➡️ Scraping Category: {category}\n🔗 {url}")
    driver.get(url)
    time.sleep(2)

    # Scroll to load all products
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height

    products = driver.find_elements(By.CSS_SELECTOR, "div.product-item-info")
    print(f"✅ Found {len(products)} products.")

    data = []
    for product in products:
        try:
            title_el = product.find_element(By.CSS_SELECTOR, "a.product-item-link")
            title = title_el.text.strip()
            product_url = title_el.get_attribute("href").strip()

            # === Extract prices using specific selectors ===
            try:
                new_price_el = product.find_element(By.CSS_SELECTOR, ".special-price .price")
                new_price = normalize_price(new_price_el.text)
            except:
                try:
                    new_price_el = product.find_element(By.CSS_SELECTOR, ".price-box .price")
                    new_price = normalize_price(new_price_el.text)
                except:
                    new_price = None

            try:
                old_price_el = product.find_element(By.CSS_SELECTOR, ".old-price .price")
                old_price = normalize_price(old_price_el.text)
            except:
                old_price = None

            product_code = extract_sku(title)
            normalized_code = normalize_sku(product_code)

            data.append({
                "Item Name": title,
                "Old Price": old_price,
                "New Price": new_price,
                "Product URL": product_url,
                "Product Code": product_code,
                "Normalized Code": normalized_code
            })

        except Exception as e:
            print(f"⚠️ Skipped product due to error: {e}")
            continue

    if data:
        timestamp = datetime.now().strftime("%Y-%m-%d")
        safe_category = re.sub(r"[^\w\s-]", "", category).replace(" ", "_")
        output_file = os.path.join(output_dir, f"2b_{safe_category}_{timestamp}.xlsx")
        df_out = pd.DataFrame(data)
        df_out.to_excel(output_file, index=False, engine='openpyxl')
        style_excel_file(output_file)
        print(f"💾 Saved {len(data)} products to {output_file}")
    else:
        print("⚠️ No product data collected.")

# === Done ===
driver.quit()
print("🏁 All categories processed for 2B.")
print("✅ Scraping completed successfully!")
print(f"📂 Output files saved in: {output_dir}")