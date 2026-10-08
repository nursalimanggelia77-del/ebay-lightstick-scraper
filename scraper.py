import urllib.request
import xml.etree.ElementTree as ET
import pandas as pd
import re
from datetime import datetime
import os

# 1. Daftar Kata Kunci Pencarian (10 Group K-Pop)
KEYWORDS = [
    "bts lightstick", "blackpink lightstick", "exo lightstick", 
    "seventeen lightstick", "ateez lightstick", "h2h lightstick", 
    "straykidz lightstick", "aespa lightstick", "riize lightstick", 
    "baby monster lightstick"
]

output_filename = "DATA_ebay_lightstick_RAW_ONLY.xlsx"

# 2. Cek apakah file Excel sudah ada sebelumnya (untuk append data harian)
if os.path.exists(output_filename):
    try:
        existing_df = pd.read_excel(output_filename)
        existing_rows = existing_df.to_dict('records')
    except Exception:
        existing_rows = []
else:
    existing_rows = []

new_rows = []
timestamp_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept-Language": "en-US,en;q=0.9"
}

print(f"=== Memulai Scraping eBay RSS Feed [{timestamp_str}] ===")

for keyword in KEYWORDS:
    print(f"\n[+] Scraping RSS Feed kata kunci: {keyword}")
    encoded_kw = keyword.replace(' ', '+')
    # Menggunakan Official eBay RSS Feed URL
    rss_url = f"https://www.ebay.com/sch/i.html?_nkw={encoded_kw}&_rss=1"
    
    try:
        req = urllib.request.Request(rss_url, headers=headers)
        with urllib.request.urlopen(req, timeout=15) as response:
            xml_data = response.read()
            root = ET.fromstring(xml_data)
            
            # Navigasi ke item di XML RSS
            channel = root.find('channel')
            if channel is not None:
                items = channel.findall('item')
                print(f"  -> Ditemukan {len(items)} produk dari RSS Feed")
                
                for item in items:
                    title = item.findtext('title', default='').strip()
                    if not title or "Shop on eBay" in title:
                        continue
                    
                    link = item.findtext('link', default='')
                    description = item.findtext('description', default='')
                    
                    # Ekstrak harga dari deskripsi RSS jika ada
                    price_match = re.search(r'\$[\d,]+\.?\d*', description)
                    raw_price = price_match.group(0) if price_match else ""
                    
                    # Simpan data mentah dengan 18 kolom
                    new_rows.append({
                        "수집_타임스탬프": timestamp_str,
                        "검색_키워드": keyword,
                        "BoyGroup_Dummy": "",
                        "BigAgency_Dummy": "",
                        "상품_제목": title,
                        "중고_판매_가격": raw_price,
                        "Price_USD": "",
                        "상품_상태": "Used",
                        "Cond_LikeNew_Dummy": "",
                        "배송료": "Check Listing",
                        "Shipping_Cost_USD": "",
                        "판매자_아이디": "eBay Seller",
                        "판매자_피드백_수": "",
                        "Feedback_Count": "",
                        "판매자_긍정_피드백(%)": "",
                        "Reputation_Y": "",
                        "판매자_국가": "International",
                        "Country_Dummy": ""
                    })
    except Exception as e:
        print(f"  [-] Error saat membaca RSS {keyword}: {e}")

# 3. Gabungkan dan simpan ke file Excel
all_rows = existing_rows + new_rows
df_final = pd.DataFrame(all_rows)
df_final.to_excel(output_filename, index=False)

print(f"\n[=== Selesai! Total {len(df_final)} baris data tersimpan di '{output_filename}' ===]")
